package com.example.officeconvert.convert;

import com.example.officeconvert.config.ConversionProperties;

/**
 * 任务级转换选项（需求 12.1 可选参数）。字段为包装类型，null 表示沿用全局默认。
 */
public class TaskOptions {

    public String markdownDialect;
    public Boolean parseAttachments;
    public Integer maxDepth;
    public Boolean enableOcr;              // 业务确认：不需要 OCR，接受参数但记录“已忽略”
    public Boolean includeHeaderFooter;
    public Boolean includeComments;
    public Boolean includeFootnotes;
    public String revisionStrategy;        // accept | reject | mark
    public String formulaMode;             // value | formula | both
    public Boolean includeHiddenSheets;
    public Boolean includeHiddenSlides;
    public String notesStrategy;           // include | omit

    public String dialect(ConversionProperties p) {
        return markdownDialect != null ? markdownDialect : p.getMarkdownDialect();
    }

    public boolean parseAttachments(ConversionProperties p) {
        return parseAttachments == null || parseAttachments;
    }

    public int maxDepth(ConversionProperties p) {
        if (maxDepth == null) return p.getMaxDepth();
        return Math.max(0, Math.min(maxDepth, p.getMaxDepth()));
    }

    public boolean includeHeaderFooter(ConversionProperties p) { return p.isIncludeHeaderFooter(); }
    public boolean includeComments(ConversionProperties p) { return p.isIncludeComments(); }
    public boolean includeFootnotes(ConversionProperties p) { return p.isIncludeFootnotes(); }

    public String revisionStrategy(ConversionProperties p) {
        return revisionStrategy != null ? revisionStrategy : p.getRevisionStrategy();
    }

    public String formulaMode(ConversionProperties p) {
        return formulaMode != null ? formulaMode : p.getFormulaMode();
    }

    public boolean includeHiddenSheets(ConversionProperties p) { return p.isIncludeHiddenSheets(); }
    public boolean includeHiddenSlides(ConversionProperties p) { return p.isIncludeHiddenSlides(); }

    public String notesStrategy(ConversionProperties p) {
        return notesStrategy != null ? notesStrategy : p.getNotesStrategy();
    }

    public java.util.Map<String, Object> summary(ConversionProperties p) {
        java.util.Map<String, Object> m = new java.util.LinkedHashMap<>();
        m.put("markdownDialect", dialect(p));
        m.put("parseAttachments", parseAttachments(p));
        m.put("maxDepth", maxDepth(p));
        m.put("enableOcr", Boolean.TRUE.equals(enableOcr) ? Boolean.TRUE + "（已忽略：业务确认不需要 OCR）" : false);
        m.put("includeHeaderFooter", includeHeaderFooter(p));
        m.put("includeComments", includeComments(p));
        m.put("includeFootnotes", includeFootnotes(p));
        m.put("revisionStrategy", revisionStrategy(p));
        m.put("formulaMode", formulaMode(p));
        m.put("includeHiddenSheets", includeHiddenSheets(p));
        m.put("includeHiddenSlides", includeHiddenSlides(p));
        m.put("notesStrategy", notesStrategy(p));
        return m;
    }
}
