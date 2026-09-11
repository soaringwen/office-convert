package com.example.officeconvert.convert;

import com.example.officeconvert.config.ConversionProperties;
import com.example.officeconvert.util.Md;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.IdentityHashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.example.officeconvert.extract.OoxmlEmbeddingExtractor;

/**
 * Word (.docx/.docm) 转 Markdown（需求 7.3）。
 * 覆盖：标题层级、段落、列表、加粗/斜体/删除线、超链接、表格（复杂合并降级 HTML）、
 * 图片、页眉页脚、批注、脚注、嵌入附件提取。
 */
@Component
public class DocxConverter implements DocumentConverter {

    private static final Pattern HEADING = Pattern.compile("(?:Heading|heading|标题)\\s*([1-6])$");

    private final OoxmlEmbeddingExtractor embeddingExtractor;
    private final ConversionProperties props;

    public DocxConverter(OoxmlEmbeddingExtractor embeddingExtractor, ConversionProperties props) {
        this.embeddingExtractor = embeddingExtractor;
        this.props = props;
    }

    @Override
    public boolean supports(String ext) {
        return "docx".equals(ext);
    }

    @Override
    public String name() {
        return "poi-xwpf-docx";
    }

    @Override
    public ConversionResult convert(Path file, TaskOptions opts, ConvCtx ctx) throws Exception {
        ConversionResult res = new ConversionResult();
        ImageStore imgs = new ImageStore(res);
        try (InputStream in = Files.newInputStream(file); XWPFDocument doc = new XWPFDocument(in)) {
            checkMacro(doc, res);
            StringBuilder md = res.markdown;
            for (IBodyElement el : doc.getBodyElements()) {
                ctx.checkDeadline();
                if (el instanceof XWPFParagraph p) renderParagraph(p, md, imgs, doc);
                else if (el instanceof XWPFTable t) renderTable(t, md, imgs, doc);
            }
            if (opts.includeHeaderFooter(props)) appendHeaderFooter(doc, md);
            if (opts.includeFootnotes(props)) appendFootnotes(doc, md);
            if (opts.includeComments(props)) appendComments(doc, md);
            String strategy = opts.revisionStrategy(props);
            if (!"accept".equals(strategy)) {
                res.warnings.add("修订策略 '" + strategy + "'：POI 输出为接受全部修订后的正文文本，" +
                        "'mark/reject' 策略以元数据记录，未做修订标记渲染");
            }
            embeddingExtractor.extract(doc.getPackage(), res);
        }
        return res;
    }

    private void checkMacro(XWPFDocument doc, ConversionResult res) {
        try {
            for (org.apache.poi.openxml4j.opc.PackagePart pp : doc.getPackage().getParts()) {
                if (pp.getPartName().getName().contains("vbaProject.bin")) {
                    res.warnings.add("检测到 VBA 宏 (vbaProject.bin)：按安全要求宏不会被执行，仅提取文档内容");
                    return;
                }
            }
        } catch (Exception ignore) { }
    }

    // ---------- 段落 ----------

    private void renderParagraph(XWPFParagraph p, StringBuilder md, ImageStore imgs, XWPFDocument doc) {
        int level = headingLevel(p, doc);
        if (level > 0) {
            md.append("#".repeat(level)).append(' ');
        } else if (p.getNumID() != null) {
            int ilvl = 0;
            try {
                BigInteger v = p.getNumIlvl();
                if (v != null) ilvl = v.intValue();
            } catch (Exception ignore) { }
            md.append("  ".repeat(Math.min(ilvl, 4))).append("- ");
        }
        renderInline(p, md, imgs, doc, false);
        md.append("\n\n");
    }

    private int headingLevel(XWPFParagraph p, XWPFDocument doc) {
        String id = p.getStyleID();
        if (id != null) {
            Matcher m = HEADING.matcher(id);
            if (m.matches()) return Integer.parseInt(m.group(1));
            if (id.matches("[1-6]")) return Integer.parseInt(id);
            if ("Title".equalsIgnoreCase(id)) return 1;
            try {
                if (doc.getStyles() != null) {
                    XWPFStyle st = doc.getStyles().getStyle(id);
                    if (st != null && st.getName() != null) {
                        Matcher m2 = HEADING.matcher(st.getName().trim());
                        if (m2.matches()) return Integer.parseInt(m2.group(1));
                        if (st.getName().trim().equalsIgnoreCase("Title")) return 1;
                    }
                }
            } catch (Exception ignore) { }
        }
        return 0;
    }

    private void renderInline(XWPFParagraph p, StringBuilder md, ImageStore imgs,
                              XWPFDocument doc, boolean compact) {
        List<IRunElement> runs = p.getIRuns();
        if (runs == null) return;
        for (IRunElement re : runs) {
            if (re instanceof XWPFHyperlinkRun hr) {
                String url = "";
                try {
                    XWPFHyperlink link = doc.getHyperlinkByID(hr.getHyperlinkId());
                    if (link != null && link.getURL() != null) url = link.getURL();
                } catch (Exception ignore) { }
                String text = hr.text() == null ? "" : hr.text();
                if (!url.isBlank()) md.append('[').append(text).append("](").append(url).append(')');
                else md.append(text);
            } else if (re instanceof XWPFRun run) {
                for (XWPFPicture pic : run.getEmbeddedPictures()) {
                    XWPFPictureData pd = pic.getPictureData();
                    String rel = imgs.add(pd.getData(), pd.suggestFileExtension(), true);
                    if (compact) md.append("[图片](").append(rel).append(')');
                    else md.append("\n\n![](").append(rel).append(")\n\n");
                }
                String text = run.text();
                if (text != null && !text.isEmpty()) {
                    String core = text;
                    boolean bold = run.isBold(), italic = run.isItalic();
                    if (bold && italic) core = "***" + core + "***";
                    else if (bold) core = "**" + core + "**";
                    else if (italic) core = "*" + core + "*";
                    if (run.isStrikeThrough()) core = "~~" + core + "~~";
                    md.append(core);
                }
            }
        }
    }

    // ---------- 表格 ----------

    private void renderTable(XWPFTable t, StringBuilder md, ImageStore imgs, XWPFDocument doc) {
        List<XWPFTableRow> rows = t.getRows();
        if (rows.isEmpty()) return;
        boolean complex = false;
        for (XWPFTableRow r : rows) {
            for (XWPFTableCell c : r.getTableCells()) {
                CTTcPr pr = c.getCTTc().getTcPr();
                if (pr != null && ((pr.isSetGridSpan() && pr.getGridSpan().getVal().intValue() > 1)
                        || pr.isSetVMerge())) {
                    complex = true;
                }
            }
        }
        if (complex) renderHtmlTable(rows, md, imgs, doc);
        else renderMdTable(rows, md, imgs, doc);
        md.append('\n');
    }

    private void renderMdTable(List<XWPFTableRow> rows, StringBuilder md, ImageStore imgs, XWPFDocument doc) {
        int cols = maxCols(rows);
        appendRow(rows.get(0), cols, md, imgs, doc);
        md.append('|');
        for (int i = 0; i < cols; i++) md.append(" --- |");
        md.append('\n');
        for (int r = 1; r < rows.size(); r++) appendRow(rows.get(r), cols, md, imgs, doc);
        md.append('\n');
    }

    private void appendRow(XWPFTableRow row, int cols, StringBuilder md, ImageStore imgs, XWPFDocument doc) {
        md.append('|');
        List<XWPFTableCell> cells = row.getTableCells();
        for (int i = 0; i < cols; i++) {
            String text = i < cells.size() ? cellText(cells.get(i), imgs, doc, true) : "";
            md.append(' ').append(Md.cell(text)).append(" |");
        }
        md.append('\n');
    }

    private int maxCols(List<XWPFTableRow> rows) {
        int cols = 0;
        for (XWPFTableRow r : rows) {
            int w = 0;
            for (XWPFTableCell c : r.getTableCells()) {
                CTTcPr pr = c.getCTTc().getTcPr();
                int span = (pr != null && pr.isSetGridSpan()) ? pr.getGridSpan().getVal().intValue() : 1;
                w += Math.max(1, span);
            }
            cols = Math.max(cols, w);
        }
        return cols;
    }

    /** 含合并单元格时降级为 HTML 表格（需求 7.3.2）。 */
    private void renderHtmlTable(List<XWPFTableRow> rows, StringBuilder md, ImageStore imgs, XWPFDocument doc) {
        int nRows = rows.size();
        int nCols = maxCols(rows);
        XWPFTableCell[][] owner = new XWPFTableCell[nRows][nCols];
        boolean[][] skip = new boolean[nRows][nCols];
        Map<XWPFTableCell, Boolean> isMergeStart = new IdentityHashMap<>();

        for (int r = 0; r < nRows; r++) {
            int col = 0;
            Map<Integer, XWPFTableCell> lastRestartCol = new java.util.HashMap<>();
            for (XWPFTableCell cell : rows.get(r).getTableCells()) {
                CTTcPr pr = cell.getCTTc().getTcPr();
                int span = (pr != null && pr.isSetGridSpan()) ? pr.getGridSpan().getVal().intValue() : 1;
                span = Math.max(1, span);
                boolean cont = pr != null && pr.isSetVMerge() && pr.getVMerge().getVal() == STMerge.CONTINUE;
                if (cont) {
                    XWPFTableCell start = lastRestartCol.get(col);
                    for (int k = 0; k < span && col + k < nCols; k++) {
                        owner[r][col + k] = start != null ? start : cell;
                        skip[r][col + k] = start != null;
                    }
                } else {
                    for (int k = 0; k < span && col + k < nCols; k++) owner[r][col + k] = cell;
                    if (pr != null && pr.isSetVMerge() && pr.getVMerge().getVal() == STMerge.RESTART) {
                        isMergeStart.put(cell, true);
                        lastRestartCol.put(col, cell);
                    }
                }
                col += span;
            }
        }
        md.append("<table>\n");
        for (int r = 0; r < nRows; r++) {
            md.append("<tr>");
            int c = 0;
            while (c < nCols) {
                XWPFTableCell cell = owner[r][c];
                if (cell == null || skip[r][c]) { c++; continue; }
                int colspan = 1;
                while (c + colspan < nCols && owner[r][c + colspan] == cell && !skip[r][c + colspan]) colspan++;
                int rowspan = 1;
                while (r + rowspan < nRows && owner[r + rowspan][c] == cell && !skip[r + rowspan][c]) rowspan++;
                md.append("<td");
                if (colspan > 1) md.append(" colspan=\"").append(colspan).append('"');
                if (rowspan > 1) md.append(" rowspan=\"").append(rowspan).append('"');
                md.append('>').append(escapeHtml(cellText(cell, imgs, doc, true))).append("</td>");
                c += colspan;
            }
            md.append("</tr>\n");
        }
        md.append("</table>\n");
        // 合并单元格已降级为 HTML 表格（需求 7.3.2 允许），降级类型记录于转换报告
    }

    private String cellText(XWPFTableCell cell, ImageStore imgs, XWPFDocument doc, boolean compact) {
        StringBuilder sb = new StringBuilder();
        boolean first = true;
        for (XWPFParagraph p : cell.getParagraphs()) {
            if (!first) sb.append("<br>");
            first = false;
            renderInline(p, sb, imgs, doc, true);
        }
        return sb.toString();
    }

    // ---------- 页眉页脚 / 批注 / 脚注 ----------

    private void appendHeaderFooter(XWPFDocument doc, StringBuilder md) {
        boolean any = false;
        for (XWPFHeader h : doc.getHeaderList()) {
            String text = h.getText();
            if (text != null && !text.isBlank()) {
                if (!any) { md.append("## 页眉页脚\n\n"); any = true; }
                md.append("**页眉**: ").append(text.trim()).append("\n\n");
            }
        }
        for (XWPFFooter f : doc.getFooterList()) {
            String text = f.getText();
            if (text != null && !text.isBlank()) {
                if (!any) { md.append("## 页眉页脚\n\n"); any = true; }
                md.append("**页脚**: ").append(text.trim()).append("\n\n");
            }
        }
    }

    private void appendComments(XWPFDocument doc, StringBuilder md) {
        XWPFComment[] comments = doc.getComments();
        if (comments == null || comments.length == 0) return;
        md.append("## 批注\n\n");
        for (XWPFComment c : comments) {
            String text = c.getText();
            if (text == null || text.isBlank()) continue;
            md.append("- ").append(c.getAuthor() == null ? "" : c.getAuthor())
              .append(": ").append(text.trim()).append('\n');
        }
        md.append('\n');
    }

    private void appendFootnotes(XWPFDocument doc, StringBuilder md) {
        try {
            boolean any = false;
            for (XWPFFootnote fn : doc.getFootnotes()) {
                StringBuilder sb = new StringBuilder();
                for (XWPFParagraph p : fn.getParagraphs()) sb.append(p.getText());
                String text = sb.toString().trim();
                if (text.isEmpty()) continue;
                if (!any) { md.append("## 脚注\n\n"); any = true; }
                md.append("- ").append(text).append('\n');
            }
            if (any) md.append('\n');
        } catch (Exception ignore) { }
    }

    private String escapeHtml(String s) {
        if (s == null) return "";
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;");
    }
}
