package com.example.officeconvert.extract;

import com.example.officeconvert.convert.ConversionResult;
import com.example.officeconvert.convert.EmbeddedObject;
import com.example.officeconvert.detect.DetectedType;
import com.example.officeconvert.detect.FileTypeDetector;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.poifs.filesystem.Ole10Native;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.springframework.stereotype.Component;

import java.io.ByteArrayInputStream;
import java.io.InputStream;
import java.util.regex.Pattern;

/**
 * 从 OOXML（docx/xlsx/pptx）包中提取嵌入附件（需求 7.6）。
 * 支持形态：/word|xl|ppt/embeddings/ 下的 Package 对象与 OLE 对象(.bin)，
 * 以及仅指向外部路径的链接对象（需求 7.9，不下载）。
 */
@Component
public class OoxmlEmbeddingExtractor {

    private static final Pattern EMBED_DIR = Pattern.compile("^/(word|xl|ppt)/embeddings/.*");

    private final FileTypeDetector detector;

    public OoxmlEmbeddingExtractor(FileTypeDetector detector) {
        this.detector = detector;
    }

    public void extract(OPCPackage pkg, ConversionResult res) {
        try {
            for (PackagePart pp : pkg.getParts()) {
                String name = pp.getPartName().getName();
                if (!EMBED_DIR.matcher(name).matches()) continue;
                byte[] data;
                try (InputStream is = pp.getInputStream()) {
                    data = is.readAllBytes();
                }
                String originalName = null;
                boolean nameAvailable = false;
                String suggestedExt = null;

                if (name.endsWith(".bin")) {
                    String[] oleName = parseOle10Native(data);
                    if (oleName != null) {
                        originalName = oleName[0];
                        nameAvailable = true;
                        suggestedExt = oleName[1];
                    }
                }
                if (!nameAvailable) {
                    DetectedType dt = detector.detect(new ByteArrayInputStream(data), null);
                    if (dt != null && !"bin".equals(dt.ext) && !"zip".equals(dt.ext)) {
                        suggestedExt = dt.ext;
                    } else {
                        int dot = name.lastIndexOf('.');
                        if (dot > 0) suggestedExt = name.substring(dot + 1);
                    }
                }
                res.embedded.add(new EmbeddedObject(data, originalName, nameAvailable,
                        "嵌入对象 " + name, suggestedExt));
            }
            // 外部链接对象：仅记录，不访问（需求 7.9）
            for (PackageRelationship rel : pkg.getRelationships()) {
                if (rel.getTargetMode() == TargetMode.EXTERNAL) {
                    String type = rel.getRelationshipType() == null ? "" : rel.getRelationshipType();
                    if (type.contains("oleObject") || type.contains("/package")) {
                        res.externals.add(new ConversionResult.ExternalRef(
                                "外部关联对象（未取得文件内容）", rel.getTargetURI().toString(), "OLE 外部链接"));
                    }
                }
            }
        } catch (Exception e) {
            res.warnings.add("嵌入对象扫描异常: " + e.getMessage());
        }
    }

    /** 解析 Ole10Native 流，返回 {原始文件名, 扩展名}。 */
    private String[] parseOle10Native(byte[] data) {
        try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(data))) {
            DirectoryEntry root = fs.getRoot();
            if (!root.hasEntry("\u0001Ole10Native")) return null;
            Ole10Native ole = Ole10Native.createFromEmbeddedOleObject(
                    (org.apache.poi.poifs.filesystem.DirectoryNode) root);
            String label = ole.getLabel();
            String fileName = ole.getFileName();
            String chosen = (fileName != null && !fileName.isBlank()) ? fileName : label;
            if (chosen == null || chosen.isBlank()) return null;
            String ext = "bin";
            int dot = chosen.lastIndexOf('.');
            if (dot >= 0 && dot < chosen.length() - 1) ext = chosen.substring(dot + 1);
            return new String[]{chosen, ext};
        } catch (Exception e) {
            return null;
        }
    }
}
