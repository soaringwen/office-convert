package com.example.officeconvert.convert;

import com.example.officeconvert.config.ConversionProperties;
import com.example.officeconvert.extract.OoxmlEmbeddingExtractor;
import com.example.officeconvert.util.Md;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

/**
 * Excel (.xlsx) 转 Markdown（需求 7.4）。
 * 覆盖：工作表顺序与命名、隐藏工作表、有效数据区域、公式与计算值（三种输出模式）、
 * 合并单元格、多行文本、批注、超链接、错误值、冻结窗格/命名区域元数据、
 * 大型工作表分块 + CSV 补充产物（不静默截断）。
 */
@Component
public class XlsxConverter implements DocumentConverter {

    private final OoxmlEmbeddingExtractor embeddingExtractor;
    private final ConversionProperties props;

    public XlsxConverter(OoxmlEmbeddingExtractor embeddingExtractor, ConversionProperties props) {
        this.embeddingExtractor = embeddingExtractor;
        this.props = props;
    }

    @Override
    public boolean supports(String ext) {
        return "xlsx".equals(ext);
    }

    @Override
    public String name() {
        return "poi-xssf-xlsx";
    }

    @Override
    public ConversionResult convert(Path file, TaskOptions opts, ConvCtx ctx) throws Exception {
        ConversionResult res = new ConversionResult();
        ImageStore store = new ImageStore(res);
        try (InputStream in = Files.newInputStream(file); XSSFWorkbook wb = new XSSFWorkbook(in)) {
            DataFormatter fmt = new DataFormatter();
            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();
            String mode = opts.formulaMode(props);

            appendNamedRangeInfo(wb, res);

            boolean anySheet = false;
            for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                ctx.checkDeadline();
                XSSFSheet sh = wb.getSheetAt(i);
                String name = wb.getSheetName(i);
                boolean hidden = wb.isSheetHidden(i) || wb.isSheetVeryHidden(i);
                if (hidden && !opts.includeHiddenSheets(props)) {
                    res.warnings.add("已按配置跳过隐藏工作表: " + name);
                    continue;
                }
                anySheet = true;
                res.markdown.append("## ").append(name);
                if (hidden) res.markdown.append("（隐藏工作表）");
                res.markdown.append("\n\n");

                PaneInformation pane = sh.getPaneInformation();
                if (pane != null && pane.isFreezePane()) {
                    res.markdown.append("> 冻结窗格：前 ").append(pane.getHorizontalSplitPosition())
                            .append(" 行 / 前 ").append(pane.getVerticalSplitPosition()).append(" 列\n\n");
                }
                renderSheet(sh, res, store, ev, fmt, mode, opts);
            }
            if (!anySheet) res.markdown.append("（工作簿不包含可输出的工作表）\n");

            if (hasExternalLinks(wb)) {
                res.warnings.add("工作簿引用了外部链接工作簿，按策略未访问外部文件（需求 7.9）");
            }
            embeddingExtractor.extract(wb.getPackage(), res);
        }
        return res;
    }

    private void renderSheet(XSSFSheet sh, ConversionResult res, ImageStore store,
                             FormulaEvaluator ev, DataFormatter fmt, String mode, TaskOptions opts) {
        int rows = sh.getLastRowNum() + 1;
        if (rows <= 0) {
            res.markdown.append("（空工作表）\n\n");
            return;
        }
        int cols = 0;
        for (int r = 0; r < rows; r++) {
            XSSFRow row = sh.getRow(r);
            if (row != null && row.getLastCellNum() > 0) cols = Math.max(cols, row.getLastCellNum());
        }
        if (cols <= 0) {
            res.markdown.append("（空工作表）\n\n");
            return;
        }

        // 构建单元格文本网格
        String[][] grid = new String[rows][cols];
        for (int r = 0; r < rows; r++) {
            XSSFRow row = sh.getRow(r);
            for (int c = 0; c < cols; c++) {
                XSSFCell cell = row == null ? null : row.getCell(c);
                grid[r][c] = cellText(cell, ev, fmt, mode, opts);
            }
        }
        // 合并单元格：值保留在左上角，其余置空并记录告警
        List<CellRangeAddress> merged = sh.getMergedRegions();
        if (merged != null && !merged.isEmpty()) {
            for (CellRangeAddress rg : merged) {
                if (rg.getFirstRow() >= rows || rg.getFirstColumn() >= cols) continue;
                String v = grid[rg.getFirstRow()][rg.getFirstColumn()];
                for (int r = rg.getFirstRow(); r <= rg.getLastRow() && r < rows; r++) {
                    for (int c = rg.getFirstColumn(); c <= rg.getLastColumn() && c < cols; c++) {
                        grid[r][c] = (r == rg.getFirstRow() && c == rg.getFirstColumn()) ? v : "";
                    }
                }
            }
            res.warnings.add("工作表 '" + sh.getSheetName() + "' 含合并单元格：Markdown 无法直接表达，"
                    + "已将值保留在区域左上角，其余单元格置空（可参照原始文件核对）");
        }

        int largeRows = props.getLargeSheetRows();
        if (rows > largeRows) {
            // 大型工作表：CSV 补充产物 + 分块输出（需求 7.4.3，不静默截断）
            String csv = toCsv(grid);
            String rel = store.add("sheet", csv.getBytes(StandardCharsets.UTF_8), "csv", false);
            res.markdown.append("> 本工作表共 ").append(rows).append(" 行，超过阈值 ").append(largeRows)
                    .append(" 行，已分块输出；完整数据见 [CSV 补充产物](").append(rel).append(")。\n\n");
            int chunk = Math.max(100, props.getSheetChunkRows());
            res.markdown.append("### 分块索引\n\n");
            for (int start = 0; start < rows; start += chunk) {
                int end = Math.min(start + chunk, rows);
                res.markdown.append("- [行 ").append(start + 1).append('-').append(end).append("](#行-")
                        .append(start + 1).append('-').append(end).append(")\n");
            }
            res.markdown.append('\n');
            for (int start = 0; start < rows; start += chunk) {
                int end = Math.min(start + chunk, rows);
                res.markdown.append("### 行 ").append(start + 1).append('-').append(end).append("\n\n");
                appendGrid(grid, start, end, res.markdown);
            }
        } else {
            appendGrid(grid, 0, rows, res.markdown);
        }
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

    private String cellText(XSSFCell cell, FormulaEvaluator ev, DataFormatter fmt, String mode, TaskOptions opts) {
        if (cell == null) return "";
        String base;
        CellType type = cell.getCellType();
        switch (type) {
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
                    default -> "`=" + f + "` → " + val;   // 默认同时保留公式与计算值（需求 18）
                };
            }
            case STRING -> base = cell.getStringCellValue();
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date d = cell.getDateCellValue();
                    base = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(d);
                } else {
                    base = fmt.formatCellValue(cell);
                }
            }
            case BOOLEAN -> base = String.valueOf(cell.getBooleanCellValue());
            case ERROR -> {
                try {
                    base = "错误值: " + FormulaError.forInt(cell.getErrorCellValue()).getString();
                } catch (Exception e) {
                    base = "错误值";
                }
            }
            default -> {
                return "";
            }
        }
        XSSFHyperlink link = cell.getHyperlink();
        if (link != null && link.getAddress() != null && !link.getAddress().isBlank()) {
            base = "[" + base + "](" + link.getAddress() + ")";
        }
        try {
            if (cell.getCellComment() != null) {
                String ct = cell.getCellComment().getString().getString();
                if (ct != null && !ct.isBlank()) base = base + "（批注: " + ct.trim() + "）";
            }
        } catch (Exception ignore) { }
        return base;
    }

    private String toCsv(String[][] grid) {
        StringBuilder sb = new StringBuilder();
        for (String[] row : grid) {
            for (int c = 0; c < row.length; c++) {
                if (c > 0) sb.append(',');
                sb.append(csvQuote(row[c]));
            }
            sb.append('\n');
        }
        return sb.toString();
    }

    private String csvQuote(String v) {
        if (v == null) return "";
        if (v.contains(",") || v.contains("\"") || v.contains("\n")) {
            return '"' + v.replace("\"", "\"\"") + '"';
        }
        return v;
    }

    /** 命名区域写入元数据说明（需求 7.4.1）。 */
    private void appendNamedRangeInfo(XSSFWorkbook wb, ConversionResult res) {
        try {
            List<org.apache.poi.xssf.usermodel.XSSFName> names = wb.getAllNames();
            if (names == null || names.isEmpty()) return;
            StringBuilder sb = new StringBuilder();
            for (org.apache.poi.xssf.usermodel.XSSFName nm : names) {
                if (sb.length() > 0) sb.append(", ");
                sb.append(nm.getNameName());
            }
            res.markdown.append("> 命名区域：").append(sb).append("\n\n");
        } catch (Exception ignore) { }
    }

    private boolean hasExternalLinks(XSSFWorkbook wb) {
        try {
            for (org.apache.poi.openxml4j.opc.PackagePart pp : wb.getPackage().getParts()) {
                if (pp.getPartName().getName().startsWith("/xl/externalLinks")) return true;
            }
        } catch (Exception ignore) { }
        return false;
    }
}
