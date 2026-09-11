package com.example.officeconvert.convert;

import com.example.officeconvert.detect.DetectedType;
import com.example.officeconvert.detect.FileTypeDetector;
import com.example.officeconvert.model.ConversionException;
import com.example.officeconvert.model.ConversionStatus;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.Ole10Native;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.springframework.stereotype.Component;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayDeque;
import java.util.Deque;
import java.util.List;

/**
 * OLE 包装对象解包转换器（修复需求 5.1"其他文件仅提取并保留原件"之外的
 * 常见嵌入形态：以 OLE 2.0 容器包装的真实文件）。
 *
 * 处理两类形态：
 * 1. \u0001Ole10Native 流 —— 图标式嵌入，原始文件名可恢复；
 * 2. CONTENTS 流 —— Package 对象（如 Acrobat 嵌入 PDF、拖拽嵌入的任意文件），
 *    CONTENTS 即原始文件内容。
 *
 * 解包后按内部真实类型重新路由到对应转换器，最多解包 3 层（防递归炸弹）。
 */
@Component
public class OleConverter implements DocumentConverter {

    private static final String SELF = "ole-unwrap";
    private static final int MAX_UNWRAP_DEPTH = 3;

    private final FileTypeDetector detector;
    private final List<DocumentConverter> converters;

    public OleConverter(FileTypeDetector detector, List<DocumentConverter> converters) {
        this.detector = detector;
        this.converters = converters;
    }

    @Override
    public boolean supports(String ext) {
        return "ole".equals(ext);
    }

    @Override
    public String name() {
        return SELF;
    }

    @Override
    public ConversionResult convert(Path file, TaskOptions opts, ConvCtx ctx) throws Exception {
        byte[] data;
        try (InputStream in = Files.newInputStream(file)) {
            data = in.readAllBytes();
        }

        // 逐层解包（防嵌套 OLE 包装）
        byte[] payload = data;
        String originalName = null;
        boolean unwrapped = false;
        int depth = 0;
        while (depth < MAX_UNWRAP_DEPTH) {
            ctx.checkDeadline();
            UnwrapResult r = unwrapOnce(payload);
            if (r == null) break;
            payload = r.bytes;
            if (r.name != null) originalName = r.name;
            unwrapped = true;
            depth++;
            DetectedType probe = detector.detect(new ByteArrayInputStream(payload), null);
            if (!"ole".equals(probe.ext)) break;
        }

        if (!unwrapped) {
            throw new ConversionException(ConversionStatus.UNSUPPORTED,
                    "OLE 容器中未找到可提取的原始数据（Ole10Native/CONTENTS 均不存在），原始文件已保留");
        }

        DetectedType inner = detector.detect(new ByteArrayInputStream(payload), null);
        DocumentConverter target = null;
        for (DocumentConverter c : converters) {
            if (SELF.equals(c.name())) continue;
            if (c.supports(inner.ext)) { target = c; break; }
        }
        if (target == null) {
            throw new ConversionException(ConversionStatus.UNSUPPORTED,
                    "OLE 包装内部类型 ." + inner.ext + " 暂不支持转换，原始文件已保留");
        }

        // 写入临时文件后交给对应转换器
        Path tmp = Files.createTempFile("ole-unwrap-", "." + inner.ext);
        try {
            Files.write(tmp, payload);
            ConversionResult res = target.convert(tmp, opts, ctx);
            res.warnings.add(0, "OLE 包装对象已解包：内部类型 " + inner.ext
                    + (originalName != null ? "，原始文件名: " + originalName : "，原始文件名不可用"));
            return res;
        } finally {
            Files.deleteIfExists(tmp);
        }
    }

    // ---------------- 解包 ----------------

    private static class UnwrapResult {
        final byte[] bytes;
        final String name;

        UnwrapResult(byte[] bytes, String name) {
            this.bytes = bytes;
            this.name = name;
        }
    }

    private UnwrapResult unwrapOnce(byte[] data) {
        try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(data))) {
            DirectoryEntry root = fs.getRoot();
            if (root.hasEntry("\u0001Ole10Native")) {
                Ole10Native ole = Ole10Native.createFromEmbeddedOleObject((DirectoryNode) root);
                String name = ole.getFileName();
                if (name == null || name.isBlank()) name = ole.getLabel();
                return new UnwrapResult(ole.getDataBuffer(),
                        (name != null && !name.isBlank()) ? name : null);
            }
            // 通用扫描：真实载荷可能存放在 CONTENTS / package / Embedded 等任意命名流中，
            // 按文件头魔数（ZIP/PDF/OLE/图片）识别最大的候选流。
            DocumentEntry best = null;
            long bestSize = -1;
            Deque<DirectoryEntry> stack = new ArrayDeque<>();
            stack.push(root);
            int guard = 0;
            while (!stack.isEmpty() && guard++ < 500) {
                DirectoryEntry dir = stack.pop();
                for (org.apache.poi.poifs.filesystem.Entry e : dir) {
                    if (e instanceof DirectoryEntry de) {
                        stack.push(de);
                    } else if (e instanceof DocumentEntry docE && docE.getSize() > 64) {
                        if (hasKnownMagic(docE)) {
                            if (docE.getSize() > bestSize) {
                                bestSize = docE.getSize();
                                best = docE;
                            }
                        }
                    }
                }
            }
            if (best != null) {
                try (DocumentInputStream dis = new DocumentInputStream(best)) {
                    return new UnwrapResult(dis.readAllBytes(), null);
                }
            }
        } catch (Exception ignore) {
            // 非 OLE 容器或解析失败，返回 null 由上层判断
        }
        return null;
    }

    /** 读取流前 8 字节并判断是否为已知文件魔数。 */
    private boolean hasKnownMagic(DocumentEntry entry) {
        try (DocumentInputStream dis = new DocumentInputStream(entry)) {
            byte[] head = new byte[8];
            int n = 0, total = 0;
            while (total < 8 && (n = dis.read(head, total, 8 - total)) > 0) total += n;
            if (total < 4) return false;
            // ZIP (OOXML/ODF/压缩包内容)
            if (head[0] == 'P' && head[1] == 'K' && head[2] == 3 && head[3] == 4) return true;
            // PDF
            if (head[0] == '%' && head[1] == 'P' && head[2] == 'D' && head[3] == 'F') return true;
            // OLE2（嵌套复合文档）
            if ((head[0] & 0xFF) == 0xD0 && (head[1] & 0xFF) == 0xCF
                    && (head[2] & 0xFF) == 0x11 && (head[3] & 0xFF) == 0xE0) return true;
            // 图片
            if ((head[0] & 0xFF) == 0x89 && head[1] == 'P' && head[2] == 'N' && head[3] == 'G') return true;
            if ((head[0] & 0xFF) == 0xFF && (head[1] & 0xFF) == 0xD8 && (head[2] & 0xFF) == 0xFF) return true;
            if (head[0] == 'G' && head[1] == 'I' && head[2] == 'F') return true;
            if (head[0] == 'B' && head[1] == 'M') return true;
            return false;
        } catch (Exception e) {
            return false;
        }
    }
}
