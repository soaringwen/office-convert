package com.example.officeconvert.convert;

import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Set;

/**
 * 图片文件：不转换内容，作为资源文件提取并在 Markdown 中引用（需求 5.1 常见附件）。
 */
@Component
public class ImageConverter implements DocumentConverter {

    private static final Set<String> SUPPORTED = Set.of("png", "jpg", "jpeg", "gif", "bmp", "webp");

    @Override
    public boolean supports(String ext) {
        return SUPPORTED.contains(ext);
    }

    @Override
    public String name() {
        return "image-passthrough";
    }

    @Override
    public ConversionResult convert(Path file, TaskOptions opts, ConvCtx ctx) throws Exception {
        ConversionResult res = new ConversionResult();
        byte[] data;
        try (InputStream in = Files.newInputStream(file)) {
            data = in.readAllBytes();
        }
        String name = file.getFileName().toString();
        String rel = new ImageStore(res).add("image", data, extOf(file), true);
        res.markdown.append("# ").append(name).append("\n\n![](").append(rel).append(")\n");
        return res;
    }

    private String extOf(Path file) {
        String name = file.getFileName().toString();
        int dot = name.lastIndexOf('.');
        return dot >= 0 ? name.substring(dot + 1).toLowerCase() : "bin";
    }
}
