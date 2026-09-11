package com.example.officeconvert.config;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.springframework.context.annotation.Configuration;

import jakarta.annotation.PostConstruct;

/**
 * POI 安全参数调优：
 * 默认最小解压比 0.01 会对包含高度压缩小部件（如 vbaProject、嵌入对象）的正常文档
 * 产生 "Zip bomb detected" 误判（需求 10.1 要求防压缩炸弹，但不能误伤正常文件）。
 * 调整为 0.0005（仍可拦截 2000:1 以上的压缩炸弹）。
 */
@Configuration
public class PoiSecurityConfig {

    @PostConstruct
    public void init() {
        ZipSecureFile.setMinInflateRatio(0.0005);
        ZipSecureFile.setMaxTextSize(20L * 1024 * 1024);
    }
}
