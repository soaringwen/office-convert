package com.example.officeconvert.convert;

import com.example.officeconvert.model.ConversionException;
import com.example.officeconvert.model.ConversionStatus;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.springframework.stereotype.Component;

import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.Set;

/**
 * Python sidecar 兜底转换器（MarkItDown + oletools，需求 11.3 可插拔机制）。
 * 用于 Java 原生转换器未覆盖的格式（如 .rtf），sidecar 未启用时明确返回"不受支持"。
 */
@Component
public class SidecarConverter implements DocumentConverter {

    /** Java 原生转换器之外的兜底扩展名。 */
    private static final Set<String> FALLBACK_EXTS = Set.of("rtf", "xlsb", "doc", "xls", "ppt", "odt", "ods", "odp");

    private final com.example.officeconvert.config.ConversionProperties props;
    private final ObjectMapper mapper = new ObjectMapper();
    private final HttpClient http = HttpClient.newBuilder()
            .connectTimeout(Duration.ofSeconds(5))
            .build();

    public SidecarConverter(com.example.officeconvert.config.ConversionProperties props) {
        this.props = props;
    }

    @Override
    public boolean supports(String ext) {
        return FALLBACK_EXTS.contains(ext);
    }

    @Override
    public String name() {
        return "markitdown-sidecar";
    }

    @Override
    public ConversionResult convert(Path file, TaskOptions opts, ConvCtx ctx) throws Exception {
        if (!props.getSidecar().isEnabled()) {
            throw new ConversionException(ConversionStatus.UNSUPPORTED,
                    "当前部署未启用 MarkItDown sidecar，无法转换该格式（可将 office-convert.sidecar.enabled 置为 true 并启动 python-sidecar）");
        }
        String boundary = "----officeconvert" + System.nanoTime();
        String fileName = file.getFileName().toString();
        byte[] fileBytes = Files.readAllBytes(file);
        var byteArrays = new java.io.ByteArrayOutputStream();
        byteArrays.writeBytes(("--" + boundary + "\r\n"
                + "Content-Disposition: form-data; name=\"file\"; filename=\"" + fileName + "\"\r\n"
                + "Content-Type: application/octet-stream\r\n\r\n").getBytes(java.nio.charset.StandardCharsets.UTF_8));
        byteArrays.writeBytes(fileBytes);
        byteArrays.writeBytes(("\r\n--" + boundary + "--\r\n").getBytes(java.nio.charset.StandardCharsets.UTF_8));

        HttpRequest request = HttpRequest.newBuilder()
                .uri(URI.create(props.getSidecar().getBaseUrl() + "/convert"))
                .timeout(Duration.ofSeconds(120))
                .header("Content-Type", "multipart/form-data; boundary=" + boundary)
                .POST(HttpRequest.BodyPublishers.ofByteArray(byteArrays.toByteArray()))
                .build();
        HttpResponse<String> response = http.send(request, HttpResponse.BodyHandlers.ofString());
        if (response.statusCode() != 200) {
            throw new ConversionException(ConversionStatus.FAILED,
                    "sidecar 转换失败，HTTP " + response.statusCode());
        }
        JsonNode node = mapper.readTree(response.body());
        ConversionResult res = new ConversionResult();
        String md = node.path("markdown").asText("");
        if (md.isBlank()) {
            throw new ConversionException(ConversionStatus.FAILED, "sidecar 返回空内容");
        }
        res.markdown.append(md.trim()).append('\n');
        return res;
    }
}
