package com.example.officeconvert.detect;

/** 文件类型识别结果（需求 7.2）。 */
public class DetectedType {
    public String ext;          // 规范化扩展名，如 docx / xlsx / pdf / txt
    public String mime;
    public String warning;      // 扩展名与实际格式不一致时记录告警
    public String detail;       // 检测依据说明

    public DetectedType(String ext, String mime, String detail) {
        this.ext = ext; this.mime = mime; this.detail = detail;
    }
}
