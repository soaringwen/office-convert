package com.example.officeconvert.model;

import java.util.ArrayList;
import java.util.List;

/**
 * 清单文件（需求 7.12）。
 */
public class Manifest {

    public String taskId;
    public String documentName;
    public SourceFile sourceFile = new SourceFile();
    public Conversion conversion = new Conversion();
    public List<AttachmentInfo> attachments = new ArrayList<>();
    public List<String> warnings = new ArrayList<>();
    public List<String> errors = new ArrayList<>();

    public static class SourceFile {
        public String name;
        public String type;       // detected type (canonical ext)
        public String mime;
        public long size;
        public String sha256;
    }

    public static class Conversion {
        public String status;                 // success | partial_success | failed
        public String startedAt;
        public String finishedAt;
        public String markdownDialect;
    }
}
