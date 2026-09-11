package com.example.officeconvert.convert;

import com.example.officeconvert.util.Md;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.nio.charset.Charset;
import java.nio.charset.CharsetDecoder;
import java.nio.charset.CodingErrorAction;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;

/**
 * 文本类文件（txt/md/csv/tsv/json/xml/html）转 Markdown。
 * 编码：UTF-8 优先，失败回退 GBK（需求 7.1：自动检测编码）。
 */
@Component
public class TextConverter implements DocumentConverter {

    private static final Set<String> SUPPORTED = Set.of("txt", "md", "csv", "tsv", "json", "xml", "html");

    @Override
    public boolean supports(String ext) {
        return SUPPORTED.contains(ext);
    }

    @Override
    public String name() {
        return "text";
    }

    @Override
    public ConversionResult convert(Path file, TaskOptions opts, ConvCtx ctx) throws Exception {
        ConversionResult res = new ConversionResult();
        byte[] raw;
        try (InputStream in = Files.newInputStream(file)) {
            raw = in.readAllBytes();
        }
        String text = decode(raw);
        String ext = extOf(file);
        switch (ext) {
            case "csv", "tsv" -> renderCsvTable(text, "tsv".equals(ext) ? '\t' : ',', res.markdown);
            case "md" -> res.markdown.append(text.trim()).append('\n');
            case "json", "xml", "html" -> {
                res.markdown.append("```").append(ext).append("\n")
                        .append(text.stripTrailing()).append("\n```\n");
                if ("html".equals(ext)) {
                    res.warnings.add("HTML 内容以代码块原样输出，未做 HTML 清理与渲染");
                }
            }
            default -> res.markdown.append(text.stripTrailing()).append('\n');
        }
        return res;
    }

    private String extOf(Path file) {
        String name = file.getFileName().toString();
        int dot = name.lastIndexOf('.');
        return dot >= 0 ? name.substring(dot + 1).toLowerCase() : "txt";
    }

    private String decode(byte[] raw) {
        if (raw.length >= 3 && (raw[0] & 0xFF) == 0xEF && (raw[1] & 0xFF) == 0xBB && (raw[2] & 0xFF) == 0xBF) {
            return new String(raw, 3, raw.length - 3, StandardCharsets.UTF_8);
        }
        if (raw.length >= 2 && (raw[0] & 0xFF) == 0xFF && (raw[1] & 0xFF) == 0xFE) {
            return new String(raw, 2, raw.length - 2, StandardCharsets.UTF_16LE);
        }
        if (strictDecode(raw, StandardCharsets.UTF_8)) return new String(raw, StandardCharsets.UTF_8);
        if (strictDecode(raw, Charset.forName("GBK"))) return new String(raw, Charset.forName("GBK"));
        return new String(raw, StandardCharsets.UTF_8);
    }

    private boolean strictDecode(byte[] raw, Charset cs) {
        try {
            CharsetDecoder d = cs.newDecoder()
                    .onMalformedInput(CodingErrorAction.REPORT)
                    .onUnmappableCharacter(CodingErrorAction.REPORT);
            d.decode(java.nio.ByteBuffer.wrap(raw));
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    private void renderCsvTable(String text, char sep, StringBuilder md) {
        List<List<String>> rows = parseCsv(text, sep);
        if (rows.isEmpty()) {
            md.append("（空表格文件）\n");
            return;
        }
        List<String> header = rows.get(0);
        md.append('|');
        for (String h : header) md.append(' ').append(Md.cell(h)).append(" |");
        md.append("\n|");
        for (int i = 0; i < header.size(); i++) md.append(" --- |");
        md.append('\n');
        for (int r = 1; r < rows.size(); r++) {
            md.append('|');
            List<String> row = rows.get(r);
            for (int c = 0; c < header.size(); c++) {
                md.append(' ').append(Md.cell(c < row.size() ? row.get(c) : "")).append(" |");
            }
            md.append('\n');
        }
        md.append('\n');
    }

    /** 简易 CSV 解析（支持双引号转义）。 */
    private List<List<String>> parseCsv(String text, char sep) {
        List<List<String>> rows = new ArrayList<>();
        List<String> cur = new ArrayList<>();
        StringBuilder cell = new StringBuilder();
        boolean inQuotes = false;
        for (int i = 0; i < text.length(); i++) {
            char c = text.charAt(i);
            if (inQuotes) {
                if (c == '"') {
                    if (i + 1 < text.length() && text.charAt(i + 1) == '"') {
                        cell.append('"');
                        i++;
                    } else {
                        inQuotes = false;
                    }
                } else {
                    cell.append(c);
                }
            } else if (c == '"') {
                inQuotes = true;
            } else if (c == sep) {
                cur.add(cell.toString());
                cell.setLength(0);
            } else if (c == '\n') {
                cur.add(cell.toString());
                cell.setLength(0);
                rows.add(cur);
                cur = new ArrayList<>();
            } else if (c != '\r') {
                cell.append(c);
            }
        }
        if (cell.length() > 0 || !cur.isEmpty()) {
            cur.add(cell.toString());
            rows.add(cur);
        }
        return rows;
    }
}
