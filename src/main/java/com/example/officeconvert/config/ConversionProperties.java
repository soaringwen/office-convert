package com.example.officeconvert.config;

import org.springframework.boot.context.properties.ConfigurationProperties;

/**
 * 全局转换配置（需求 13：管理配置项）。
 */
@ConfigurationProperties(prefix = "office-convert")
public class ConversionProperties {

    private String workDir = "./data";
    private long maxFileBytes = 20L * 1024 * 1024;
    private int maxFilesPerTask = 20;
    private int maxAttachments = 10;
    private int maxDepth = 2;
    private int concurrency = 2;
    private int taskTimeoutSeconds = 600;
    private int largeSheetRows = 5000;
    private int sheetChunkRows = 1000;
    private boolean includeHeaderFooter = true;
    private boolean includeComments = true;
    private boolean includeFootnotes = true;
    private String revisionStrategy = "accept";
    private String formulaMode = "both";
    private boolean includeHiddenSheets = true;
    private boolean includeHiddenSlides = true;
    private String notesStrategy = "include";
    private String markdownDialect = "gfm";
    private Sidecar sidecar = new Sidecar();

    public static class Sidecar {
        private boolean enabled = false;
        private String baseUrl = "http://127.0.0.1:8900";

        public boolean isEnabled() { return enabled; }
        public void setEnabled(boolean enabled) { this.enabled = enabled; }
        public String getBaseUrl() { return baseUrl; }
        public void setBaseUrl(String baseUrl) { this.baseUrl = baseUrl; }
    }

    public String getWorkDir() { return workDir; }
    public void setWorkDir(String workDir) { this.workDir = workDir; }
    public long getMaxFileBytes() { return maxFileBytes; }
    public void setMaxFileBytes(long v) { this.maxFileBytes = v; }
    public int getMaxFilesPerTask() { return maxFilesPerTask; }
    public void setMaxFilesPerTask(int v) { this.maxFilesPerTask = v; }
    public int getMaxAttachments() { return maxAttachments; }
    public void setMaxAttachments(int v) { this.maxAttachments = v; }
    public int getMaxDepth() { return maxDepth; }
    public void setMaxDepth(int v) { this.maxDepth = v; }
    public int getConcurrency() { return concurrency; }
    public void setConcurrency(int v) { this.concurrency = v; }
    public int getTaskTimeoutSeconds() { return taskTimeoutSeconds; }
    public void setTaskTimeoutSeconds(int v) { this.taskTimeoutSeconds = v; }
    public int getLargeSheetRows() { return largeSheetRows; }
    public void setLargeSheetRows(int v) { this.largeSheetRows = v; }
    public int getSheetChunkRows() { return sheetChunkRows; }
    public void setSheetChunkRows(int v) { this.sheetChunkRows = v; }
    public boolean isIncludeHeaderFooter() { return includeHeaderFooter; }
    public void setIncludeHeaderFooter(boolean v) { this.includeHeaderFooter = v; }
    public boolean isIncludeComments() { return includeComments; }
    public void setIncludeComments(boolean v) { this.includeComments = v; }
    public boolean isIncludeFootnotes() { return includeFootnotes; }
    public void setIncludeFootnotes(boolean v) { this.includeFootnotes = v; }
    public String getRevisionStrategy() { return revisionStrategy; }
    public void setRevisionStrategy(String v) { this.revisionStrategy = v; }
    public String getFormulaMode() { return formulaMode; }
    public void setFormulaMode(String v) { this.formulaMode = v; }
    public boolean isIncludeHiddenSheets() { return includeHiddenSheets; }
    public void setIncludeHiddenSheets(boolean v) { this.includeHiddenSheets = v; }
    public boolean isIncludeHiddenSlides() { return includeHiddenSlides; }
    public void setIncludeHiddenSlides(boolean v) { this.includeHiddenSlides = v; }
    public String getNotesStrategy() { return notesStrategy; }
    public void setNotesStrategy(String v) { this.notesStrategy = v; }
    public String getMarkdownDialect() { return markdownDialect; }
    public void setMarkdownDialect(String v) { this.markdownDialect = v; }
    public Sidecar getSidecar() { return sidecar; }
    public void setSidecar(Sidecar sidecar) { this.sidecar = sidecar; }
}
