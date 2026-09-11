package com.example.officeconvert.model;

/** 附件转换状态（需求 8.2）。 */
public enum ConversionStatus {
    SUCCESS, FAILED, UNSUPPORTED, ENCRYPTED, CORRUPTED, LIMIT_EXCEEDED,
    SECURITY_BLOCKED, NOT_EMBEDDED
}
