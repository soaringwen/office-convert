package com.example.officeconvert.convert;

import java.nio.file.Path;

/**
 * 可插拔转换器（需求 11.3：格式路由 -> 转换器 -> 统一结果 -> Markdown）。
 */
public interface DocumentConverter {

    /** 是否支持该（规范化后的）扩展名。 */
    boolean supports(String ext);

    String name();

    ConversionResult convert(Path file, TaskOptions opts, ConvCtx ctx) throws Exception;
}
