package com.example.officeconvert.extract;

import com.example.officeconvert.convert.ConversionResult;
import com.example.officeconvert.convert.EmbeddedObject;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.Ole10Native;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.springframework.stereotype.Component;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayDeque;
import java.util.Deque;

/**
 * 从旧版二进制 Office（OLE2 容器）中提取嵌入附件（需求 7.6）。
 * 支持形态：
 * - \u0001Ole10Native 流（图标式嵌入，可恢复原始文件名）；
 * - CONTENTS 流（Package 对象，原始名称不可用时生成稳定名称并标记）。
 */
@Component
public class OleEmbeddingExtractor {

    public void extract(POIFSFileSystem fs, ConversionResult res) {
        Deque<DirectoryEntry> stack = new ArrayDeque<>();
        stack.push(fs.getRoot());
        int n = 0;
        int guard = 0;
        while (!stack.isEmpty() && guard++ < 2000) {
            DirectoryEntry dir = stack.pop();
            try {
                if (dir.hasEntry("\u0001Ole10Native")) {
                    Ole10Native ole = Ole10Native.createFromEmbeddedOleObject((DirectoryNode) dir);
                    byte[] data = ole.getDataBuffer();
                    String name = ole.getFileName();
                    if (name == null || name.isBlank()) name = ole.getLabel();
                    boolean nameAvailable = name != null && !name.isBlank();
                    res.embedded.add(new EmbeddedObject(data, nameAvailable ? name : null,
                            nameAvailable, "OLE 嵌入对象 #" + (++n), guessExt(nameAvailable ? name : null)));
                    continue;
                }
                if (dir.hasEntry("CONTENTS") && dir.getName().startsWith("\u0001")) {
                    // Package 对象：CONTENTS 流即原始文件内容
                    try (DocumentInputStream dis = new DocumentInputStream(
                            (DocumentEntry) dir.getEntry("CONTENTS"))) {
                        byte[] data = readAll(dis);
                        res.embedded.add(new EmbeddedObject(data, null, false,
                                "OLE Package 对象 #" + (++n), null));
                    }
                    continue;
                }
                for (org.apache.poi.poifs.filesystem.Entry e : dir) {
                    if (e instanceof DirectoryEntry de) stack.push(de);
                }
            } catch (Exception ignore) {
                // 单个对象解析失败不影响其他对象
            }
        }
    }

    public POIFSFileSystem open(InputStream in) throws Exception {
        return new POIFSFileSystem(in);
    }

    private byte[] readAll(DocumentInputStream dis) throws Exception {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        dis.transferTo(bos);
        return bos.toByteArray();
    }

    private String guessExt(String name) {
        if (name == null) return null;
        int dot = name.lastIndexOf('.');
        if (dot >= 0 && dot < name.length() - 1) return name.substring(dot + 1);
        return null;
    }
}
