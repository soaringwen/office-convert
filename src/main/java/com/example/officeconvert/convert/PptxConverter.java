package com.example.officeconvert.convert;

import com.example.officeconvert.config.ConversionProperties;
import com.example.officeconvert.extract.OoxmlEmbeddingExtractor;
import com.example.officeconvert.util.Md;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFGraphicFrame;
import org.apache.poi.xslf.usermodel.XSLFHyperlink;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.springframework.stereotype.Component;

import java.awt.geom.Rectangle2D;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

/**
 * PowerPoint (.pptx) 转 Markdown（需求 7.5）。
 * 覆盖：按幻灯片顺序、标题、正文/列表（按位置排序恢复阅读顺序）、表格、图片、
 * 超链接、演讲者备注（策略可配）、隐藏幻灯片、图表降级说明、嵌入附件提取。
 */
@Component
public class PptxConverter implements DocumentConverter {

    private final OoxmlEmbeddingExtractor embeddingExtractor;
    private final ConversionProperties props;

    public PptxConverter(OoxmlEmbeddingExtractor embeddingExtractor, ConversionProperties props) {
        this.embeddingExtractor = embeddingExtractor;
        this.props = props;
    }

    @Override
    public boolean supports(String ext) {
        return "pptx".equals(ext);
    }

    @Override
    public String name() {
        return "poi-xslf-pptx";
    }

    @Override
    public ConversionResult convert(Path file, TaskOptions opts, ConvCtx ctx) throws Exception {
        ConversionResult res = new ConversionResult();
        ImageStore imgs = new ImageStore(res);
        try (InputStream in = Files.newInputStream(file); XMLSlideShow ppt = new XMLSlideShow(in)) {
            checkMacro(ppt, res);
            int index = 1;
            for (XSLFSlide slide : ppt.getSlides()) {
                ctx.checkDeadline();
                boolean hidden = slide.isHidden();
                if (hidden && !opts.includeHiddenSlides(props)) {
                    res.warnings.add("已按配置跳过隐藏幻灯片: 第 " + index + " 页");
                    index++;
                    continue;
                }
                res.markdown.append("## 幻灯片 ").append(index);
                if (hidden) res.markdown.append("（隐藏）");
                String title = slide.getTitle();
                if (title != null && !title.isBlank()) res.markdown.append("：").append(title.trim());
                res.markdown.append("\n\n");

                renderShapes(slide, res, imgs);
                renderNotes(slide, res, opts);
                index++;
            }
            embeddingExtractor.extract(ppt.getPackage(), res);
        }
        return res;
    }

    private void checkMacro(XMLSlideShow ppt, ConversionResult res) {
        try {
            for (org.apache.poi.openxml4j.opc.PackagePart pp : ppt.getPackage().getParts()) {
                if (pp.getPartName().getName().contains("vbaProject.bin")) {
                    res.warnings.add("检测到 VBA 宏 (vbaProject.bin)：按安全要求宏不会被执行，仅提取内容");
                    return;
                }
            }
        } catch (Exception ignore) { }
    }

    /** 依据对象位置恢复阅读顺序（需求 7.5）。 */
    private void renderShapes(XSLFSlide slide, ConversionResult res, ImageStore imgs) {
        List<XSLFShape> shapes = new ArrayList<>(slide.getShapes());
        shapes.sort(Comparator
                .comparingDouble((XSLFShape s) -> s.getAnchor().getMinY())
                .thenComparingDouble(s -> s.getAnchor().getMinX()));
        for (XSLFShape shape : shapes) {
            try {
                if (shape instanceof XSLFTable table) {
                    renderTable(table, res);
                } else if (shape instanceof XSLFPictureShape pic) {
                    XSLFPictureData pd = pic.getPictureData();
                    String rel = imgs.add(pd.getData(), pd.suggestFileExtension(), true);
                    res.markdown.append("![](").append(rel).append(")\n\n");
                } else if (shape instanceof XSLFGraphicFrame frame) {
                    renderGraphicFrame(frame, res);
                } else if (shape instanceof XSLFTextShape ts) {
                    renderTextShape(ts, res);
                }
            } catch (Exception e) {
                res.warnings.add("幻灯片对象处理异常（已跳过该对象）: " + e.getMessage());
            }
        }
    }

    private void renderTextShape(XSLFTextShape ts, ConversionResult res) {
        boolean wrote = false;
        for (XSLFTextParagraph para : ts.getTextParagraphs()) {
            List<XSLFTextRun> runs = para.getTextRuns();
            if (runs == null) continue;
            StringBuilder sb = new StringBuilder();
            for (XSLFTextRun run : runs) {
                String t = run.getRawText();
                if (t == null || t.isEmpty()) continue;
                String core = t;
                if (run.isBold() && run.isItalic()) core = "***" + core + "***";
                else if (run.isBold()) core = "**" + core + "**";
                else if (run.isItalic()) core = "*" + core + "*";
                try {
                    XSLFHyperlink link = run.getHyperlink();
                    if (link != null && link.getAddress() != null && !link.getAddress().isBlank()) {
                        core = "[" + t + "](" + link.getAddress() + ")";
                    }
                } catch (Exception ignore) { }
                sb.append(core);
            }
            if (sb.length() == 0) continue;
            int lvl = Math.max(0, para.getIndentLevel());
            res.markdown.append("  ".repeat(Math.min(lvl, 4))).append("- ").append(sb).append('\n');
            wrote = true;
        }
        if (wrote) res.markdown.append('\n');
    }

    private void renderTable(XSLFTable table, ConversionResult res) {
        List<XSLFTableRow> rows = table.getRows();
        if (rows.isEmpty()) return;
        int cols = rows.get(0).getCells().size();
        res.markdown.append('|');
        for (int c = 0; c < cols; c++) {
            res.markdown.append(' ').append(Md.cell(rows.get(0).getCells().get(c).getText())).append(" |");
        }
        res.markdown.append("\n|");
        for (int c = 0; c < cols; c++) res.markdown.append(" --- |");
        res.markdown.append('\n');
        for (int r = 1; r < rows.size(); r++) {
            res.markdown.append('|');
            List<XSLFTableCell> cells = rows.get(r).getCells();
            for (int c = 0; c < cols; c++) {
                String t = c < cells.size() ? cells.get(c).getText() : "";
                res.markdown.append(' ').append(Md.cell(t)).append(" |");
            }
            res.markdown.append('\n');
        }
        res.markdown.append('\n');
    }

    private void renderGraphicFrame(XSLFGraphicFrame frame, ConversionResult res) {
        // 图表优先转图片在纯 Java 侧不可行（POI 不做渲染），降级记录元数据（需求 7.4.4/7.5）
        XSLFChart chart = null;
        try {
            chart = frame.getChart();
        } catch (Exception ignore) { }
        if (chart != null) {
            res.markdown.append("> [图表对象]（未能渲染为图片，原始图表保留在源文件中；类型: ")
                    .append(frame.getShapeName()).append("）\n\n");
            res.degradations.add("图表未渲染为图片，已降级为占位说明并记录元数据");
        } else {
            res.markdown.append("> [复杂图形对象]（SmartArt/流程图等，未能结构化转换，已降级为占位说明）\n\n");
            res.degradations.add("SmartArt/复杂图形降级为占位说明");
        }
    }

    private void renderNotes(XSLFSlide slide, ConversionResult res, TaskOptions opts) {
        if (!"include".equals(opts.notesStrategy(props))) return;
        XSLFNotes notes = slide.getNotes();
        if (notes == null) return;
        StringBuilder sb = new StringBuilder();
        try {
            for (List<XSLFTextParagraph> paras : notes.getTextParagraphs()) {
                for (XSLFTextParagraph para : paras) {
                    for (XSLFTextRun run : para.getTextRuns()) {
                        if (run.getRawText() != null) sb.append(run.getRawText());
                    }
                }
            }
        } catch (Exception ignore) { }
        String text = sb.toString().trim();
        if (text.isEmpty()) return;
        res.markdown.append("> 演讲者备注：").append(text.replace("\n", " / ")).append("\n\n");
    }
}
