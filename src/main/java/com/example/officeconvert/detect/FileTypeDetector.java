package com.example.officeconvert.detect;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.Locale;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/**
 * 综合扩展名、MIME、文件头特征与容器内部结构识别真实类型（需求 7.2）。
 */
@Component
public class FileTypeDetector {

    private static final String OOXML_REL_WORD = "word/";
    private static final String OOXML_REL_XL = "xl/";
    private static final String OOXML_REL_PPT = "ppt/";

    public DetectedType detect(InputStream in, String declaredExt) {
        in.mark(1 << 20);
        byte[] head = new byte[8];
        int n;
        try {
            n = readFully(in, head);
        } catch (IOException e) {
            return new DetectedType("bin", "application/octet-stream", "读取失败: " + e.getMessage());
        }
        DetectedType t = detectByMagic(in, head, n);
        in.mark(0); // 允许后续重读（调用方负责 reset 或重新打开）
        // 扩展名一致性告警（以实际检测为准）
        if (declaredExt != null && t.ext != null && !t.ext.equals("bin")) {
            String d = declaredExt.toLowerCase(Locale.ROOT);
            if (!d.equals(t.ext) && !compatible(d, t.ext)) {
                t.warning = "扩展名 ." + d + " 与实际文件类型 " + t.ext + " 不一致，已按实际类型处理";
            }
        }
        return t;
    }

    private DetectedType detectByMagic(InputStream in, byte[] head, int n) {
        // OLE2 容器：旧版 Office 或加密 Office
        if (n >= 8 && (head[0] & 0xFF) == 0xD0 && (head[1] & 0xFF) == 0xCF) {
            return detectOle(in);
        }
        // ZIP 容器：OOXML
        if (n >= 4 && head[0] == 'P' && head[1] == 'K') {
            return detectZip(in);
        }
        if (n >= 4 && head[0] == '%' && head[1] == 'P' && head[2] == 'D' && head[3] == 'F') {
            return new DetectedType("pdf", "application/pdf", "PDF 魔数 %PDF");
        }
        if (n >= 8 && head[0] == (byte) 0x89 && head[1] == 'P' && head[2] == 'N' && head[3] == 'G') {
            return new DetectedType("png", "image/png", "PNG 魔数");
        }
        if (n >= 3 && head[0] == (byte) 0xFF && head[1] == (byte) 0xD8 && head[2] == (byte) 0xFF) {
            return new DetectedType("jpg", "image/jpeg", "JPEG 魔数");
        }
        if (n >= 6 && head[0] == 'G' && head[1] == 'I' && head[2] == 'F') {
            return new DetectedType("gif", "image/gif", "GIF 魔数");
        }
        if (n >= 2 && head[0] == 'B' && head[1] == 'M') {
            return new DetectedType("bmp", "image/bmp", "BMP 魔数");
        }
        if (looksLikeUtf8Text(head, n)) {
            return new DetectedType("txt", "text/plain", "纯文本启发式判断");
        }
        return new DetectedType("bin", "application/octet-stream", "未知二进制格式");
    }

    private DetectedType detectOle(InputStream in) {
        try {
            in.reset();
            try (POIFSFileSystem fs = new POIFSFileSystem(in)) {
                if (fs.getRoot().hasEntry("EncryptedPackage")
                        || fs.getRoot().hasEntry("EncryptionInfo")) {
                    return new DetectedType("encrypted", "application/octet-stream", "OLE 加密容器（EncryptedPackage）");
                }
                if (fs.getRoot().hasEntry("WordDocument")) {
                    return new DetectedType("doc", "application/msword", "OLE 根目录含 WordDocument");
                }
                if (fs.getRoot().hasEntry("Workbook") || fs.getRoot().hasEntry("Book")) {
                    return new DetectedType("xls", "application/vnd.ms-excel", "OLE 根目录含 Workbook");
                }
                if (fs.getRoot().hasEntry("PowerPoint Document")) {
                    return new DetectedType("ppt", "application/vnd.ms-powerpoint", "OLE 根目录含 PowerPoint Document");
                }
                return new DetectedType("ole", "application/octet-stream", "通用 OLE2 容器");
            }
        } catch (Exception e) {
            return new DetectedType("ole", "application/octet-stream", "OLE 容器解析异常: " + e.getMessage());
        }
    }

    private DetectedType detectZip(InputStream in) {
        try {
            in.reset();
            try (ZipInputStream zis = new ZipInputStream(in, StandardCharsets.UTF_8)) {
                boolean hasContentTypes = false, word = false, xl = false, ppt = false, vba = false;
                ZipEntry e;
                int scanned = 0;
                while ((e = zis.getNextEntry()) != null && scanned++ < 500) {
                    String name = e.getName();
                    if (name.equals("[Content_Types].xml")) hasContentTypes = true;
                    if (name.startsWith(OOXML_REL_WORD)) word = true;
                    if (name.startsWith(OOXML_REL_XL)) xl = true;
                    if (name.startsWith(OOXML_REL_PPT)) ppt = true;
                    if (name.contains("vbaProject.bin")) vba = true;
                }
                if (hasContentTypes) {
                    if (word) return new DetectedType("docx",
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "OOXML word 包");
                    if (xl) return new DetectedType("xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "OOXML xl 包");
                    if (ppt) return new DetectedType("pptx",
                            "application/vnd.openxmlformats-officedocument.presentationml.presentation", "OOXML ppt 包");
                    return new DetectedType("zip", "application/zip", "OOXML 结构但无法确定子类型");
                }
                return new DetectedType("zip", "application/zip", "ZIP 容器（业务确认：不解析压缩包内容）");
            }
        } catch (Exception e) {
            return new DetectedType("zip", "application/zip", "ZIP 解析异常");
        }
    }

    /** 纯文本启发式：前 8 字节均可打印或为常见换行/制表符，且无 NUL（按无符号字节比较）。 */
    private boolean looksLikeUtf8Text(byte[] b, int n) {
        if (n == 0) return false;
        for (int i = 0; i < n; i++) {
            int v = b[i] & 0xFF;
            if (v == 0 || v == 0x7F) return false;
            if (v == '\n' || v == '\r' || v == '\t') continue;
            if (v < 0x20) return false;
        }
        return true;
    }

    /** 扩展名与检测类型可互相兼容的别名（例如 .xlsm 检测为 xlsx 容器）。 */
    private boolean compatible(String declared, String detected) {
        if (declared.equals(detected)) return true;
        return switch (detected) {
            case "docx" -> declared.equals("docm");
            case "xlsx" -> declared.equals("xlsm") || declared.equals("xlsb");
            case "pptx" -> declared.equals("pptm");
            case "ole" -> declared.equals("doc") || declared.equals("xls")
                    || declared.equals("ppt") || declared.equals("docm")
                    || declared.equals("xlsm") || declared.equals("pptm");
            case "txt" -> declared.equals("csv") || declared.equals("tsv") || declared.equals("md")
                    || declared.equals("json") || declared.equals("xml") || declared.equals("html")
                    || declared.equals("log");
            default -> false;
        };
    }

    private int readFully(InputStream in, byte[] buf) throws IOException {
        int off = 0;
        while (off < buf.length) {
            int r = in.read(buf, off, buf.length - off);
            if (r < 0) break;
            off += r;
        }
        return off;
    }
}
