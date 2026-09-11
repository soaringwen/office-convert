package com.example.officeconvert.pipeline;

import com.example.officeconvert.config.ConversionProperties;
import com.example.officeconvert.convert.ConvCtx;
import com.example.officeconvert.convert.ConversionResult;
import com.example.officeconvert.convert.DocumentConverter;
import com.example.officeconvert.convert.EmbeddedObject;
import com.example.officeconvert.convert.TaskOptions;
import com.example.officeconvert.detect.DetectedType;
import com.example.officeconvert.detect.FileTypeDetector;
import com.example.officeconvert.model.AttachmentInfo;
import com.example.officeconvert.model.ConversionException;
import com.example.officeconvert.model.ConversionReport;
import com.example.officeconvert.model.ConversionStatus;
import com.example.officeconvert.model.ExtractionStatus;
import com.example.officeconvert.model.Manifest;
import com.example.officeconvert.util.Hashes;
import com.example.officeconvert.util.SafeNames;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.NotOLE2FileException;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

/**
 * 递归转换编排（需求 6、7.7）：
 * 识别 -> 路由转换器 -> 生成 Markdown/资源 -> 提取嵌入附件 -> 递归（深度/数量/大小受限）
 * -> 建立 manifest 与报告。单个附件失败不影响主文档（需求 6 部分成功机制）。
 */
@Service
public class ConversionService {

    private static final List<String> CONVERTER_ORDER = List.of(
            "poi-xwpf-docx", "poi-xssf-xlsx", "poi-xslf-pptx", "pdfbox-pdf",
            "poi-legacy", "ole-unwrap", "text", "image-passthrough", "markitdown-sidecar");

    private final ConversionProperties props;
    private final FileTypeDetector detector;
    private final List<DocumentConverter> converters;
    private final ObjectMapper mapper = new ObjectMapper()
            .enable(SerializationFeature.INDENT_OUTPUT);

    public ConversionService(ConversionProperties props, FileTypeDetector detector,
                             List<DocumentConverter> converters) {
        this.props = props;
        this.detector = detector;
        this.converters = orderConverters(converters);
    }

    private List<DocumentConverter> orderConverters(List<DocumentConverter> list) {
        List<DocumentConverter> sorted = new ArrayList<>();
        for (String name : CONVERTER_ORDER) {
            list.stream().filter(c -> name.equals(c.name())).findFirst().ifPresent(sorted::add);
        }
        for (DocumentConverter c : list) {
            if (!sorted.contains(c)) sorted.add(c);
        }
        return sorted;
    }

    // ---------------- 输出结构 ----------------

    public static class DocOutcome {
        public List<AttachmentInfo> attachments = new ArrayList<>();
        public Manifest manifest = new Manifest();
        public ConversionReport report = new ConversionReport();
    }

    // ---------------- 主入口 ----------------

    /**
     * 转换单个文档（或附件）到 outDir，返回附件子树与清单。
     *
     * @param parentId    父附件 id，主文档为 null
     * @param depth       递归层级，主文档为 0，直接附件为 1
     * @param sourceType  来源位置类型描述
     */
    public DocOutcome convert(Path sourceFile, Path outDir, String taskId,
                              String parentId, int depth, String sourceType, ConvCtx ctx) throws Exception {
        long t0 = System.currentTimeMillis();
        Files.createDirectories(outDir);
        DocOutcome outcome = new DocOutcome();
        Manifest manifest = outcome.manifest;
        manifest.taskId = taskId;
        ConversionReport report = outcome.report;
        report.configSummary = ctx.opts.summary(props);

        ctx.checkDeadline();

        // 1. 读取与基础校验（大小限制，需求 7.1 / 7.7）
        byte[] data;
        try (InputStream in = Files.newInputStream(sourceFile)) {
            data = in.readAllBytes();
        }
        if (data.length == 0) {
            throw new ConversionException(ConversionStatus.CORRUPTED, "文件为空");
        }
        if (data.length > props.getMaxFileBytes()) {
            throw new ConversionException(ConversionStatus.LIMIT_EXCEEDED,
                    "文件大小 " + data.length + " 字节超过限制 " + props.getMaxFileBytes() + " 字节");
        }

        // 2. 类型识别（不依赖扩展名，需求 7.2）
        DetectedType type = detector.detect(new java.io.ByteArrayInputStream(data), extOf(sourceFile));
        if (type.warning != null) manifest.warnings.add(type.warning);

        String sha256 = Hashes.sha256(data);
        manifest.sourceFile.name = sourceFile.getFileName().toString();
        manifest.sourceFile.type = type.ext;
        manifest.sourceFile.mime = type.mime;
        manifest.sourceFile.size = data.length;
        manifest.sourceFile.sha256 = sha256;
        manifest.conversion.markdownDialect = ctx.opts.dialect(props);
        manifest.conversion.startedAt = Instant.now().toString();

        // 3. 恶意可执行文件检查（需求 10.1：不执行，标记安全拦截）
        if (data.length >= 2 && ((data[0] == 'M' && data[1] == 'Z')
                || (data.length >= 4 && data[0] == 0x7F && data[1] == 'E' && data[2] == 'L' && data[3] == 'F'))) {
            manifest.conversion.status = ConversionStatus.SECURITY_BLOCKED.name().toLowerCase();
            manifest.errors.add("检测到可执行文件，按安全策略拦截转换（原始文件已保留）");
            writeOutputs(outDir, data, null, List.of(), manifest, report, t0, "可执行文件");
            return outcome;
        }

        // 4. 重复/循环附件检测（需求 7.7）
        if (depth > 0 && ctx.markSeen(sha256)) {
            String firstId = ctx.seenHashes.get(sha256);
            AttachmentInfo dup = baseAttachment(sha256, data.length, type, parentId, depth, sourceType);
            dup.conversionStatus = firstId != null ? ConversionStatus.SUCCESS.name().toLowerCase()
                    : ConversionStatus.LIMIT_EXCEEDED.name().toLowerCase();
            dup.extractionStatus = ExtractionStatus.SKIPPED.name().toLowerCase();
            dup.failureReason = firstId != null
                    ? "与附件 " + firstId + " 内容相同（重复附件，未生成副本）"
                    : "检测到循环引用，停止展开";
            if (firstId != null) dup.duplicateOf = firstId;
            outcome.attachments.add(dup);
            manifest.conversion.status = "success";
            finishManifest(manifest, t0);
            return outcome;
        }

        // 5. 路由转换器（需求 11.3）
        DocumentConverter converter = null;
        for (DocumentConverter c : converters) {
            if (c.supports(type.ext)) { converter = c; break; }
        }

        ConversionResult result = null;
        ConversionStatus failureStatus = null;
        String failureReason = null;
        long tBody = System.currentTimeMillis();
        if (converter == null) {
            failureStatus = ConversionStatus.UNSUPPORTED;
            failureReason = "格式 ." + type.ext + " 暂不支持转换（原始文件已保留）";
            manifest.warnings.add(failureReason);
        } else {
            try {
                result = converter.convert(sourceFile, ctx.opts, ctx);
                report.converters.add(converter.name());
            } catch (ConversionException ce) {
                failureStatus = ce.status;
                failureReason = ce.getMessage();
            } catch (EncryptedDocumentException e) {
                failureStatus = ConversionStatus.ENCRYPTED;
                failureReason = "文件已加密，系统不尝试破解";
            } catch (InvalidFormatException | NotOLE2FileException e) {
                failureStatus = ConversionStatus.CORRUPTED;
                failureReason = "文件结构损坏或格式无效";
            } catch (Exception e) {
                String msg = String.valueOf(e.getMessage());
                if (msg.toLowerCase().contains("password") || msg.toLowerCase().contains("encrypt")) {
                    failureStatus = ConversionStatus.ENCRYPTED;
                    failureReason = "文件已加密，系统不尝试破解";
                } else if (msg.toLowerCase().contains("corrupt") || msg.toLowerCase().contains("invalid header")
                        || msg.toLowerCase().contains("truncated")) {
                    failureStatus = ConversionStatus.CORRUPTED;
                    failureReason = "文件损坏，无法解析";
                } else {
                    failureStatus = ConversionStatus.FAILED;
                    failureReason = "转换器异常: " + rootMessage(e);
                }
            }
        }
        report.stageTimings.put("body_conversion_ms", System.currentTimeMillis() - tBody);
        report.bodySuccess = result != null;

        // 6. 递归处理嵌入附件（需求 7.7），失败不影响主文档
        List<AttachmentInfo> children = new ArrayList<>();
        if (result != null) {
            int maxDepth = ctx.opts.maxDepth(props);
            for (EmbeddedObject emb : result.embedded) {
                ctx.checkDeadline();
                children.add(processEmbedded(emb, taskId, parentId, depth, maxDepth, outDir, ctx, report));
            }
            for (ConversionResult.ExternalRef ext : result.externals) {
                AttachmentInfo info = new AttachmentInfo();
                info.id = nextAttId(ctx);
                info.parentId = parentId;
                info.depth = depth + 1;
                info.name = ext.label;
                info.safeName = SafeNames.sanitize(ext.label, "external-ref");
                info.externalRef = true;
                info.externalUri = ext.uri;
                info.sourceLocation = new AttachmentInfo.SourceLocation("external_link", ext.location);
                info.extractionStatus = ExtractionStatus.SKIPPED.name().toLowerCase();
                info.conversionStatus = ConversionStatus.NOT_EMBEDDED.name().toLowerCase();
                info.failureReason = "外部关联文件：仅保留链接，未取得文件内容，不自动下载（需求 7.9）";
                children.add(info);
            }
        }
        outcome.attachments.addAll(children);
        manifest.attachments.addAll(children);

        // 7. 汇总状态后生成 content.md（正文 + 附件章节，需求 7.10.2）
        report.truncated = result != null && result.truncated;
        if (result != null) {
            report.warnings.addAll(result.warnings);
            report.warnings.addAll(result.degradations.stream().map(d -> "降级: " + d).toList());
            report.imageCount = (int) result.assets.stream().filter(a -> a.image).count();
            report.resourceCount = result.assets.size();
        }
        manifest.conversion.status = resolveStatus(result != null, children, failureStatus);
        if (failureReason != null && failureStatus != null && result == null) {
            manifest.errors.add(failureReason);
            report.errors.add(java.util.Map.of(
                    "code", failureStatus.name().toLowerCase(),
                    "message", failureReason));
        }
        manifest.documentName = manifest.sourceFile.name;
        finishManifest(manifest, t0);
        report.status = manifest.conversion.status;
        String bodyMarkdown = result != null ? result.markdown.toString() : placeholder(failureStatus, failureReason);
        bodyMarkdown = bodyMarkdown + buildAttachmentSection(children, result != null);
        writeOutputs(outDir, data, bodyMarkdown,
                result != null ? result.assets : List.of(), manifest, report, t0,
                result != null ? sourceType : null);
        return outcome;
    }

    // ---------------- 附件处理 ----------------

    private AttachmentInfo processEmbedded(EmbeddedObject emb, String taskId, String parentId,
                                           int depth, int maxDepth, Path parentOutDir,
                                           ConvCtx ctx, ConversionReport report) {
        report.attachmentTotal++;
        AttachmentInfo info = new AttachmentInfo();
        info.id = nextAttId(ctx);
        info.parentId = parentId;
        info.depth = depth + 1;
        info.name = emb.originalName;
        info.originalNameAvailable = emb.nameAvailable;
        info.sha256 = Hashes.sha256(emb.data);
        info.size = emb.data.length;
        info.sourceLocation = new AttachmentInfo.SourceLocation("embedded_object", emb.sourceLocation);

        String displayName = emb.nameAvailable
                ? SafeNames.sanitize(emb.originalName, "attachment-" + info.id)
                : "attachment-" + info.id + "." + (emb.suggestedExt != null ? emb.suggestedExt : "bin");
        info.safeName = displayName;
        if (!emb.nameAvailable) {
            info.failureReason = "原始文件名不可用，已生成稳定文件名（需求 7.6）";
        }

        try {
            // 大小限制
            if (emb.data.length > props.getMaxFileBytes()) {
                info.extractionStatus = ExtractionStatus.SKIPPED.name().toLowerCase();
                info.conversionStatus = ConversionStatus.LIMIT_EXCEEDED.name().toLowerCase();
                info.failureReason = (info.failureReason == null ? "" : info.failureReason + "；")
                        + "附件大小超过限制 " + props.getMaxFileBytes() + " 字节";
                report.failedCount++;
                report.limitReached = "max_attachment_size";
                return info;
            }
            // 数量限制
            if (!ctx.tryAcquireAttachment()) {
                ctx.releaseAttachment();
                info.extractionStatus = ExtractionStatus.SKIPPED.name().toLowerCase();
                info.conversionStatus = ConversionStatus.LIMIT_EXCEEDED.name().toLowerCase();
                info.failureReason = (info.failureReason == null ? "" : info.failureReason + "；")
                        + "附件数量达到上限 " + props.getMaxAttachments() + "，原件已保留但未解析";
                Path overflowDir = parentOutDir.resolve("attachments").resolve("overflow");
                Files.createDirectories(overflowDir);
                Path saved = overflowDir.resolve(SafeNames.sanitize(displayName, info.id + ".bin"));
                Files.write(saved, emb.data);
                info.originalPath = rel(parentOutDir, saved);
                report.limitReached = "max_attachments";
                report.skippedCount++;
                return info;
            }
            try {
                Path attDir = parentOutDir.resolve("attachments").resolve(info.id);
                Files.createDirectories(attDir);
                Path original = attDir.resolve(displayName);
                Files.write(original, emb.data);
                info.originalPath = rel(parentOutDir, original);
                info.extractionStatus = ExtractionStatus.SUCCESS.name().toLowerCase();
                report.extractedSuccess++;

                // 深度限制：保留原件，不再展开（需求 7.7）
                if (depth + 1 > maxDepth) {
                    info.conversionStatus = ConversionStatus.LIMIT_EXCEEDED.name().toLowerCase();
                    info.failureReason = (info.failureReason == null ? "" : info.failureReason + "；")
                            + "达到最大递归层级 " + maxDepth + "，停止展开";
                    report.limitReached = "max_depth";
                    ctx.seenHashes.put(info.sha256, info.id);
                    return info;
                }

                // 递归转换
                DocOutcome child = convert(original, attDir, taskId, info.id, depth + 1,
                        "附件 " + info.safeName, ctx);
                info.detectedType = child.manifest.sourceFile.type;
                info.mime = child.manifest.sourceFile.mime;
                info.conversionStatus = child.manifest.conversion.status;
                info.markdownPath = rel(parentOutDir, attDir.resolve("content.md"));
                info.children.addAll(child.attachments);
                ctx.seenHashes.put(info.sha256, info.id);
                if (ConversionStatus.SUCCESS.name().equalsIgnoreCase(child.manifest.conversion.status)) {
                    report.convertedSuccess++;
                } else {
                    report.failedCount++;
                }
                return info;
            } finally {
                ctx.releaseAttachment();
            }
        } catch (Exception e) {
            info.extractionStatus = ExtractionStatus.FAILED.name().toLowerCase();
            info.conversionStatus = ConversionStatus.FAILED.name().toLowerCase();
            info.failureReason = "附件处理异常: " + rootMessage(e);
            report.failedCount++;
            return info;
        }
    }

    // ---------------- Markdown 附件章节 ----------------

    private String buildAttachmentSection(List<AttachmentInfo> attachments, boolean hasBody) {
        if (attachments.isEmpty()) return "";
        StringBuilder md = new StringBuilder("\n\n## 嵌入附件\n\n");
        for (AttachmentInfo a : attachments) {
            md.append("### 附件：").append(a.safeName).append("\n\n");
            md.append("- 来源位置：").append(a.sourceLocation.detail == null ? "未知" : a.sourceLocation.detail).append('\n');
            md.append("- 文件类型：").append(a.detectedType != null ? a.detectedType : "未识别").append('\n');
            if (a.duplicateOf != null) md.append("- 重复附件：与 ").append(a.duplicateOf).append(" 内容相同\n");
            if (a.externalRef) {
                md.append("- 外部关联文件：[").append(a.externalUri).append("](").append(a.externalUri)
                        .append(")（未取得文件内容）\n");
                md.append("\n> 此对象为外部关联文件，未实际嵌入文档，系统未自动访问。\n\n");
                continue;
            }
            md.append("- 转换状态：").append(statusLabel(a.conversionStatus)).append('\n');
            if (a.originalPath != null) {
                md.append("- 原始文件：[下载原始附件](./").append(a.originalPath).append(")\n");
            }
            if (a.markdownPath != null) {
                md.append("- 转换内容：[查看 Markdown](./").append(a.markdownPath).append(")\n");
            }
            if (a.failureReason != null) {
                boolean ok = "success".equalsIgnoreCase(a.conversionStatus) || a.duplicateOf != null;
                if (ok) {
                    md.append("- 备注：").append(a.failureReason).append('\n');
                } else {
                    md.append("\n> 此附件未能完整转换。原因：").append(a.failureReason).append('\n');
                    if (a.originalPath != null) {
                        md.append("> 原始附件已保留：./").append(a.originalPath).append('\n');
                    }
                }
            }
            md.append('\n');
        }
        return md.toString();
    }

    // ---------------- 输出写盘 ----------------

    private void writeOutputs(Path outDir, byte[] sourceData, String markdown,
                              List<ConversionResult.Asset> assets,
                              Manifest manifest, ConversionReport report, long t0,
                              String sourceType) throws IOException {
        Files.createDirectories(outDir);
        // original/
        Path originalDir = outDir.resolve("original");
        Files.createDirectories(originalDir);
        Files.write(originalDir.resolve(manifest.sourceFile.name), sourceData);
        // assets/
        if (assets != null) {
            Path assetsDir = outDir.resolve("assets");
            Files.createDirectories(assetsDir);
            for (ConversionResult.Asset a : assets) {
                Path target = assetsDir.resolve(a.fileName).normalize();
                if (!target.startsWith(assetsDir)) continue; // 路径穿越防护
                Files.write(target, a.data);
            }
        }
        // content.md
        if (markdown != null) {
            Files.write(outDir.resolve("content.md"),
                    markdown.getBytes(StandardCharsets.UTF_8));
            manifest.conversion.finishedAt = Instant.now().toString();
        }
        // manifest.json / conversion-report.json
        mapper.writeValue(outDir.resolve("manifest.json").toFile(), manifest);
        report.durationMs = System.currentTimeMillis() - t0;
        report.status = manifest.conversion.status;
        mapper.writeValue(outDir.resolve("conversion-report.json").toFile(), report);
    }

    // ---------------- 辅助 ----------------

    private AttachmentInfo baseAttachment(String sha256, long size, DetectedType type,
                                          String parentId, int depth, String sourceType) {
        AttachmentInfo info = new AttachmentInfo();
        info.parentId = parentId;
        info.depth = depth + 1;
        info.sha256 = sha256;
        info.size = size;
        info.detectedType = type.ext;
        info.mime = type.mime;
        info.sourceLocation = new AttachmentInfo.SourceLocation(sourceType, null);
        return info;
    }

    private String nextAttId(ConvCtx ctx) {
        return String.format("att-%03d", ctx.idSeq.incrementAndGet());
    }

    private String resolveStatus(boolean bodyOk, List<AttachmentInfo> children, ConversionStatus failure) {
        if (!bodyOk) {
            // 正文失败时保留更精确的失败类别（需求 8.2 状态语义）
            return switch (failure == null ? ConversionStatus.FAILED : failure) {
                case UNSUPPORTED -> "unsupported";
                case ENCRYPTED -> "encrypted";
                case CORRUPTED -> "corrupted";
                case SECURITY_BLOCKED -> "security_blocked";
                case LIMIT_EXCEEDED -> "limit_exceeded";
                default -> "failed";
            };
        }
        boolean allOk = children.stream().allMatch(a ->
                "success".equalsIgnoreCase(a.conversionStatus)
                        || "not_embedded".equalsIgnoreCase(a.conversionStatus)
                        || (a.duplicateOf != null));
        return allOk ? "success" : "partial_success";
    }

    private String placeholder(ConversionStatus status, String reason) {
        String s = status != null ? status.name().toLowerCase() : "failed";
        return "> 此文档未能转换（状态: " + s + "）。\n> 原因: " + (reason == null ? "未知" : reason) + "\n";
    }

    private String statusLabel(String status) {
        if (status == null) return "未知";
        return switch (status) {
            case "success" -> "成功";
            case "failed" -> "失败";
            case "unsupported" -> "格式不受支持";
            case "encrypted" -> "文件已加密";
            case "corrupted" -> "文件损坏";
            case "limit_exceeded" -> "超过限制";
            case "security_blocked" -> "安全拦截";
            case "not_embedded" -> "外部关联文件";
            case "partial_success" -> "部分成功";
            default -> status;
        };
    }

    private String rel(Path base, Path target) {
        return base.relativize(target).toString().replace('\\', '/');
    }

    private void finishManifest(Manifest manifest, long t0) {
        if (manifest.conversion.startedAt == null) {
            manifest.conversion.startedAt = Instant.now().toString();
        }
        manifest.conversion.finishedAt = Instant.now().toString();
    }

    private String extOf(Path file) {
        String name = file.getFileName().toString();
        int dot = name.lastIndexOf('.');
        return dot >= 0 ? name.substring(dot + 1) : null;
    }

    private String rootMessage(Throwable e) {
        Throwable t = e;
        while (t.getCause() != null && t.getCause() != t) t = t.getCause();
        String msg = String.valueOf(t.getMessage());
        // 不暴露服务器路径/调用栈（需求 9）
        return msg.length() > 200 ? msg.substring(0, 200) : msg;
    }

    /** 供任务层读取截止时间。 */
    public static Instant deadlineOf(ConversionProperties props) {
        return Instant.now().plus(Duration.ofSeconds(props.getTaskTimeoutSeconds()));
    }
}
