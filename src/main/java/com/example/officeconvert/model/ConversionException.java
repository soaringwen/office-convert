package com.example.officeconvert.model;

/** 单次转换中抛出的可识别错误类型，用于映射附件转换状态。 */
public class ConversionException extends RuntimeException {
    public final ConversionStatus status;

    public ConversionException(ConversionStatus status, String message) {
        super(message);
        this.status = status;
    }

    public ConversionException(ConversionStatus status, String message, Throwable cause) {
        super(message, cause);
        this.status = status;
    }
}
