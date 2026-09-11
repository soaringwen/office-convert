package com.example.officeconvert.model;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 转换报告（需求 7.13）。
 */
public class ConversionReport {

    public String status;
    public boolean bodySuccess;
    public int attachmentTotal;
    public int extractedSuccess;
    public int convertedSuccess;
    public int failedCount;
    public int skippedCount;
    public int unsupportedCount;
    public int imageCount;
    public int resourceCount;
    public List<String> warnings = new ArrayList<>();
    public List<Map<String, String>> errors = new ArrayList<>();
    public boolean truncated;
    public String limitReached;
    public long durationMs;
    public Map<String, Long> stageTimings = new LinkedHashMap<>();
    public List<String> converters = new ArrayList<>();
    public Map<String, Object> configSummary = new LinkedHashMap<>();
}
