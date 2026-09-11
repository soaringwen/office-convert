package com.example.officeconvert.convert;

import com.example.officeconvert.extract.OleEmbeddingExtractor;
import com.example.officeconvert.util.Md;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFPictureData;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hslf.usermodel.HSLFTextRun;
import org.apache.poi.hslf.usermodel.HSLFNotes;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFHyperlink;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

/**
 * 旧版二进制 Office（.doc/.xls/.ppt）转 Markdown（业务确认：首期必须支持）。
 * 附件提取通过 OLE 容器遍历完成（\u0001Ole10Native / CONTENTS）。
 */
@Component
public class LegacyOfficeConverter implements DocumentConverter {

    private final OleEmbeddingExtractor oleExtractor;

    public LegacyOfficeConverter(OleEmbeddingExtractor oleExtractor) {
        this.oleExtractor = oleExtractor;
    }

    @Override
    public boolean supports(String ext) {
        return "doc".equals(ext) || "xls".equals(ext) || "ppt".equals(ext);
    }

    @Override
    public String name() {
        return "poi-legacy";
    }

    @Override
    public ConversionResult convert(Path file, TaskOptions opts, ConvCtx ctx) throws Exception {
        String ext = extOf(file);
        return switch (ext) {
            case "doc" -> convertDoc(file, ctx);
            case "xls" -> convertXls(file, opts, ctx);
            case "ppt" -> convertPpt(file, opts, ctx);
            default -> throw new IllegalStateException("unsupported: " + ext);
        };
    }

    // ---------------- .doc ----------------

    private ConversionResult convertDoc(Path file, ConvCtx ctx) throws Exception {
        ConversionResult res = new ConversionResult();
        ImageStore imgs = new ImageStore(res);
        try (InputStream in = Files.newInputStream(file); HWPFDocument doc = new HWPFDocument(in)) {
            Range range = doc.getRange();
            int i = 0;
            while (i < range.numParagraphs()) {
                ctx.checkDeadline();
                Paragraph p = range.getParagraph(i);
                if (p.isInTable()) {
                    try {
                        Table t = range.getTable(p);
                        renderHwpfTable(t, res);
                        i += Math.max(1, t.numParagraphs());
                        continue;
                    } catch (IllegalArgumentException ignore) {
                        // 非 table 首段，按普通段落处理
                    }
                }
                String text = cleanHwpf(p.text());
                if (!text.isBlank()) res.markdown.append(text).append("\n\n");
                i++;
            }
            // 图片
            try {
                List<Picture> pics = doc.getPicturesTable().getAllPictures();
                for (Picture pic : pics) {
                    String ext = pic.suggestFileExtension();
                    imgs.add(pic.getRawContent(), ext, true);
                }
                if (!pics.isEmpty()) {
                    res.markdown.append("## 文档图片\n\n");
                    for (ConversionResult.Asset a : res.assets) {
                        res.markdown.append("![](").append(a.relPath).append(")\n\n");
                    }
                }
            } catch (Exception e) {
                res.warnings.add("旧版 Word 图片提取异常: " + e.getMessage());
            }
            res.warnings.add("旧版 .doc 格式：输出为正文文本近似转换，复杂排版/修订/批注可能不完整");
            try (POIFSFileSystem fs = new POIFSFileSystem(Files.newInputStream(file))) {
                oleExtractor.extract(fs, res);
            }
        }
        return res;
    }

    private void renderHwpfTable(Table t, ConversionResult res) {
        int cols = 0;
        for (int r = 0; r < t.numRows(); r++) cols = Math.max(cols, t.getRow(r).numCells());
        if (cols == 0) return;
        res.markdown.append('|');
        for (int c = 0; c < cols; c++) {
            res.markdown.append(' ').append(c < t.getRow(0).numCells()
                    ? Md.cell(cleanHwpf(t.getRow(0).getCell(c).text())) : "").append(" |");
        }
        res.markdown.append("\n|");
        for (int c = 0; c < cols; c++) res.markdown.append(" --- |");
        res.markdown.append('\n');
        for (int r = 1; r < t.numRows(); r++) {
            TableRow row = t.getRow(r);
            res.markdown.append('|');
            for (int c = 0; c < cols; c++) {
                res.markdown.append(' ').append(c < row.numCells()
                        ? Md.cell(cleanHwpf(row.getCell(c).text())) : "").append(" |");
            }
            res.markdown.append('\n');
        }
        res.markdown.append('\n');
    }

    private String cleanHwpf(String s) {
        if (s == null) return "";
        return s.replace("\u0007", "").replace("\r", "\n").trim();
    }

    // ---------------- .xls ----------------

    private ConversionResult convertXls(Path file, TaskOptions opts, ConvCtx ctx) throws Exception {
        ConversionResult res = new ConversionResult();
        ImageStore store = new ImageStore(res);
        try (InputStream in = Files.newInputStream(file);
             POIFSFileSystem fs = new POIFSFileSystem(in);
             HSSFWorkbook wb = new HSSFWorkbook(fs, true)) {
            DataFormatter fmt = new DataFormatter();
            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();
            String mode = opts.formulaMode(ctx.props);
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                ctx.checkDeadline();
                HSSFSheet sh = wb.getSheetAt(i);
                boolean hidden = wb.isSheetHidden(i) || wb.isSheetVeryHidden(i);
                if (hidden && !opts.includeHiddenSheets(ctx.props)) continue;
                res.markdown.append("## ").append(wb.getSheetName(i))
                        .append(hidden ? "（隐藏工作表）" : "").append("\n\n");
                renderHssfSheet(sh, res, ev, fmt, mode);
            }
            try {
                oleExtractor.extract(fs, res);
            } catch (Exception e) {
                res.warnings.add("旧版 Excel 附件提取异常: " + e.getMessage());
            }
        }
        return res;
    }

    private void renderHssfSheet(HSSFSheet sh, ConversionResult res, FormulaEvaluator ev,
                                 DataFormatter fmt, String mode) {
        int rows = sh.getLastRowNum() + 1;
        int cols = 0;
        for (int r = 0; r < rows; r++) {
            HSSFRow row = sh.getRow(r);
            if (row != null && row.getLastCellNum() > 0) cols = Math.max(cols, row.getLastCellNum());
        }
        if (rows <= 0 || cols <= 0) {
            res.markdown.append("（空工作表）\n\n");
            return;
        }
        String[][] grid = new String[rows][cols];
        for (int r = 0; r < rows; r++) {
            HSSFRow row = sh.getRow(r);
            for (int c = 0; c < cols; c++) {
                HSSFCell cell = row == null ? null : row.getCell(c);
                grid[r][c] = hssfCellText(cell, ev, fmt, mode);
            }
        }
        for (int m = 0; m < sh.getNumMergedRegions(); m++) {
            CellRangeAddress rg = sh.getMergedRegion(m);
            if (rg.getFirstRow() >= rows || rg.getFirstColumn() >= cols) continue;
            String v = grid[rg.getFirstRow()][rg.getFirstColumn()];
            for (int r = rg.getFirstRow(); r <= rg.getLastRow() && r < rows; r++) {
                for (int c = rg.getFirstColumn(); c <= rg.getLastColumn() && c < cols; c++) {
                    grid[r][c] = (r == rg.getFirstRow() && c == rg.getFirstColumn()) ? v : "";
                }
            }
        }
        appendGrid(grid, 0, rows, res.markdown);
    }

    private String hssfCellText(HSSFCell cell, FormulaEvaluator ev, DataFormatter fmt, String mode) {
        if (cell == null) return "";
        String base;
        switch (cell.getCellType()) {
            case FORMULA -> {
                String f = cell.getCellFormula();
                String val;
                try {
                    val = fmt.formatCellValue(cell, ev);
                } catch (Exception e) {
                    val = "#EVAL#";
                }
                base = switch (mode == null ? "both" : mode) {
                    case "formula" -> "`=" + f + "`";
                    case "value" -> val;
                    default -> "`=" + f + "` → " + val;
                };
            }
            case STRING -> base = cell.getStringCellValue();
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    base = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(cell.getDateCellValue());
                } else {
                    base = fmt.formatCellValue(cell);
                }
            }
            case BOOLEAN -> base = String.valueOf(cell.getBooleanCellValue());
            case ERROR -> base = "错误值";
            default -> {
                return "";
            }
        }
        HSSFHyperlink link = cell.getHyperlink();
        if (link != null && link.getAddress() != null && !link.getAddress().isBlank()) {
            base = "[" + base + "](" + link.getAddress() + ")";
        }
        return base;
    }

    private void appendGrid(String[][] grid, int fromRow, int toRow, StringBuilder md) {
        md.append('|');
        for (int c = 0; c < grid[0].length; c++) md.append(' ').append(Md.cell(grid[fromRow][c])).append(" |");
        md.append("\n|");
        for (int c = 0; c < grid[0].length; c++) md.append(" --- |");
        md.append('\n');
        for (int r = fromRow + 1; r < toRow; r++) {
            md.append('|');
            for (int c = 0; c < grid[0].length; c++) md.append(' ').append(Md.cell(grid[r][c])).append(" |");
            md.append('\n');
        }
        md.append('\n');
    }

    // ---------------- .ppt ----------------

    private ConversionResult convertPpt(Path file, TaskOptions opts, ConvCtx ctx) throws Exception {
        ConversionResult res = new ConversionResult();
        ImageStore imgs = new ImageStore(res);
        try (InputStream in = Files.newInputStream(file); HSLFSlideShow ppt = new HSLFSlideShow(in)) {
            int index = 1;
            for (HSLFSlide slide : ppt.getSlides()) {
                ctx.checkDeadline();
                res.markdown.append("## 幻灯片 ").append(index++).append("\n\n");
                boolean wrote = false;
                for (List<HSLFTextParagraph> paras : slide.getTextParagraphs()) {
                    for (HSLFTextParagraph para : paras) {
                        StringBuilder sb = new StringBuilder();
                        for (HSLFTextRun run : para.getTextRuns()) {
                            if (run.getRawText() != null) sb.append(run.getRawText());
                        }
                        if (sb.length() > 0) {
                            res.markdown.append("- ").append(sb.toString().trim()).append('\n');
                            wrote = true;
                        }
                    }
                }
                if (wrote) res.markdown.append('\n');
                try {
                    HSLFNotes notes = slide.getNotes();
                    if (notes != null && "include".equals(opts.notesStrategy(ctx.props))) {
                        StringBuilder sb = new StringBuilder();
                        for (List<HSLFTextParagraph> paras : notes.getTextParagraphs()) {
                            for (HSLFTextParagraph para : paras) {
                                for (HSLFTextRun run : para.getTextRuns()) {
                                    if (run.getRawText() != null) sb.append(run.getRawText());
                                }
                            }
                        }
                        String text = sb.toString().trim();
                        if (!text.isEmpty()) {
                            res.markdown.append("> 演讲者备注：").append(text.replace("\n", " / "))
                                    .append("\n\n");
                        }
                    }
                } catch (Exception ignore) { }
                try {
                    for (HSLFPictureData pd : ppt.getPictureData()) {
                        String ext = pd.getType().extension;
                        imgs.add(pd.getData(), ext, true);
                    }
                    if (!res.assets.isEmpty() && index == 2) {
                        res.markdown.append("## 文档图片\n\n");
                        for (ConversionResult.Asset a : res.assets) {
                            res.markdown.append("![](").append(a.relPath).append(")\n\n");
                        }
                    }
                } catch (Exception ignore) { }
            }
            res.warnings.add("旧版 .ppt 格式：输出为文本近似转换，表格/图形对象可能不完整");
            try (POIFSFileSystem fs = new POIFSFileSystem(Files.newInputStream(file))) {
                oleExtractor.extract(fs, res);
            }
        }
        return res;
    }

    private String extOf(Path file) {
        String name = file.getFileName().toString();
        int dot = name.lastIndexOf('.');
        return dot >= 0 ? name.substring(dot + 1).toLowerCase() : "";
    }
}
