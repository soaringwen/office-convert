package com.example.officeconvert.model;

import java.util.ArrayList;
import java.util.List;

/**
 * 附件信息（需求 7.6），同时用于 manifest.json 的 attachments 节点。
 * 使用公共字段以简化 Jackson 序列化。
 */
public class AttachmentInfo {

    public String id;                       // att-001
    public String parentId;                 // 父附件编号，主文档为 null
    public int depth;                       // 递归层级，主文档直接附件为 1
    public String name;                     // 原始文件名（可能不可用）
    public String safeName;                 // 安全处理后的文件名
    public boolean originalNameAvailable = true;
    public String detectedType;             // 实际检测到的格式
    public String mime;
    public long size;
    public String sha256;
    public SourceLocation sourceLocation = new SourceLocation();
    public String extractionStatus;         // ExtractionStatus
    public String conversionStatus;         // ConversionStatus
    public String failureReason;
    public String originalPath;             // 相对于所在文档输出目录
    public String markdownPath;
    public String duplicateOf;              // 重复附件指向首次出现的附件 id
    public boolean externalRef;             // 仅外部链接、未实际嵌入
    public String externalUri;
    public List<AttachmentInfo> children = new ArrayList<>();

    public static class SourceLocation {
        public String type;                 // embedded_object | ole_stream | pdf_attachment | external_link
        public String detail;

        public SourceLocation() {}
        public SourceLocation(String type, String detail) {
            this.type = type; this.detail = detail;
        }
    }
}
