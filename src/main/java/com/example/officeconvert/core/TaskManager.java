package com.example.officeconvert.core;

import com.example.officeconvert.config.ConversionProperties;
import com.example.officeconvert.convert.ConvCtx;
import com.example.officeconvert.convert.TaskOptions;
import com.example.officeconvert.model.ConversionException;
import com.example.officeconvert.model.TaskStatus;
import com.example.officeconvert.pipeline.ConversionService;
import com.example.officeconvert.util.Hashes;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import jakarta.annotation.PostConstruct;
import jakarta.annotation.PreDestroy;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.util.ArrayList;
import java.util.HexFormat;
import java.util.List;
import java.util.Map;
import java.util.UUID;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;

/**
 * 任务管理：创建/查询/取消/重试、并发控制、超时、幂等与重启恢复（需求 7.1、11.1、11.2）。
 */
@Service
public class TaskManager {

    private final ConversionProperties props;
    private final ConversionService conversionService;
    private final ObjectMapper mapper = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);
    private final Map<String, ConversionTask> tasks = new ConcurrentHashMap<>();
    private final Map<String, Future<?>> futures = new ConcurrentHashMap<>();
    private final Map<String, String> idempotency = new ConcurrentHashMap<>();
    private ExecutorService pool;

    public TaskManager(ConversionProperties props, ConversionService conversionService) {
        this.props = props;
        this.conversionService = conversionService;
    }

    public Path taskRoot() {
        return Path.of(props.getWorkDir(), "tasks");
    }

    @PostConstruct
    public void init() throws IOException {
        pool = Executors.newFixedThreadPool(Math.max(1, props.getConcurrency()));
        // 重启恢复：PROCESSING/PENDING 任务标记失败可重试（需求 11.2）
        Path root = taskRoot();
        if (!Files.exists(root)) return;
        try (var stream = Files.list(root)) {
            for (Path dir : (Iterable<Path>) stream::iterator) {
                Path tf = dir.resolve("task.json");
                if (!Files.exists(tf)) continue;
                try {
                    ConversionTask t = mapper.readValue(tf.toFile(), ConversionTask.class);
                    if (t.status == TaskStatus.PENDING || t.status == TaskStatus.PROCESSING) {
                        t.status = TaskStatus.FAILED;
                        t.failureSummary = "服务重启导致任务中断，可调用 retry 重新处理";
                        t.errors.add(t.failureSummary);
                    }
                    tasks.put(t.taskId, t);
                    if (t.idempotencyKey != null) idempotency.put(t.idempotencyKey, t.taskId);
                } catch (Exception ignore) { }
            }
        }
    }

    @PreDestroy
    public void shutdown() {
        if (pool != null) pool.shutdownNow();
    }

    // ---------------- 创建 ----------------

    public synchronized ConversionTask create(List<MultipartFile> files, TaskOptions options) throws IOException {
        if (files == null || files.isEmpty()) {
            throw new IllegalArgumentException("未提交任何文件");
        }
        if (files.size() > props.getMaxFilesPerTask()) {
            throw new IllegalArgumentException("单任务文件数超过限制 " + props.getMaxFilesPerTask());
        }
        for (MultipartFile f : files) {
            if (f == null || f.isEmpty()) throw new IllegalArgumentException("存在空文件");
            if (f.getSize() > props.getMaxFileBytes()) {
                throw new IllegalArgumentException("文件 '" + f.getOriginalFilename() + "' 大小超过限制 "
                        + props.getMaxFileBytes() + " 字节");
            }
            String name = f.getOriginalFilename();
            if (name == null || name.isBlank() || name.contains("..") || name.contains("/") || name.contains("\\")) {
                throw new IllegalArgumentException("存在非法文件名");
            }
        }

        // 基于哈希的幂等控制（需求 7.1）
        StringBuilder keySrc = new StringBuilder();
        for (MultipartFile f : files) keySrc.append(hashOf(f)).append('|');
        keySrc.append(optionsKey(options));
        String key = Hashes.sha256Utf8(keySrc.toString());
        String existingId = idempotency.get(key);
        if (existingId != null && tasks.containsKey(existingId)) {
            return tasks.get(existingId);
        }

        String taskId = "task-" + java.time.LocalDate.now().toString().replace("-", "")
                + "-" + UUID.randomUUID().toString().substring(0, 8);
        Path taskDir = taskRoot().resolve(taskId);
        Path inputDir = taskDir.resolve("input");
        Files.createDirectories(inputDir);

        ConversionTask task = new ConversionTask();
        task.taskId = taskId;
        task.options = options != null ? options : new TaskOptions();
        int i = 0;
        for (MultipartFile f : files) {
            i++;
            String safe = com.example.officeconvert.util.SafeNames.sanitize(f.getOriginalFilename(), "file-" + i);
            Path stored = inputDir.resolve(i + "-" + safe);
            f.transferTo(stored.toAbsolutePath());
            task.docs.add(new ConversionTask.DocInput(taskDir.relativize(stored).toString(), safe));
        }
        task.totalDocs = task.docs.size();
        task.idempotencyKey = key;
        tasks.put(taskId, task);
        idempotency.put(key, taskId);
        persist(task);
        submit(task);
        return task;
    }

    private void submit(ConversionTask task) {
        task.status = TaskStatus.PENDING;
        persist(task);
        futures.put(task.taskId, pool.submit(() -> run(task)));
    }

    // ---------------- 执行 ----------------

    private void run(ConversionTask task) {
        task.status = TaskStatus.PROCESSING;
        task.startedAt = java.time.Instant.now();
        task.stage = "转换中";
        persist(task);
        Path taskDir = taskRoot().resolve(task.taskId);

        ConvCtx ctx = new ConvCtx(props, task.options, ConversionService.deadlineOf(props));
        int success = 0, partial = 0, failed = 0;
        boolean cancelled = false;
        for (int i = 0; i < task.docs.size(); i++) {
            ConversionTask.DocInput doc = task.docs.get(i);
            if (ctx.cancelled.get()) { cancelled = true; break; }
            task.stage = "转换 " + (i + 1) + "/" + task.totalDocs + ": " + doc.displayName;
            Path source = taskDir.resolve(doc.storedPath).normalize();
            Path outDir = taskDir.resolve("output")
                    .resolve(String.format("%02d", i + 1) + "-"
                            + com.example.officeconvert.util.SafeNames.sanitize(doc.displayName, "doc-" + i));
            try {
                ConversionService.DocOutcome outcome = conversionService.convert(
                        source, outDir, task.taskId, null, 0, "主文档", ctx);
                task.resultDirs.add(taskDir.relativize(outDir).toString());
                task.processedAttachments += countAttachments(outcome.attachments);
                switch (outcome.manifest.conversion.status) {
                    case "success" -> success++;
                    case "partial_success" -> partial++;
                    default -> failed++;
                }
                task.warnings.addAll(outcome.manifest.warnings);
                task.errors.addAll(outcome.manifest.errors);
            } catch (ConversionException ce) {
                if (ctx.cancelled.get()) { cancelled = true; break; }
                failed++;
                task.errors.add("文档 '" + doc.displayName + "' 转换失败: " + ce.getMessage());
            } catch (Exception e) {
                if (ctx.cancelled.get()) { cancelled = true; break; }
                failed++;
                task.errors.add("文档 '" + doc.displayName + "' 转换异常: " + e.getMessage());
            } finally {
                task.processedDocs++;
                persist(task);
            }
        }
        task.finishedAt = java.time.Instant.now();
        if (cancelled) {
            task.status = TaskStatus.CANCELLED;
        } else if (failed == 0 && partial == 0) {
            task.status = TaskStatus.SUCCESS;
        } else if (success + partial > 0) {
            task.status = TaskStatus.PARTIAL_SUCCESS;
        } else {
            task.status = TaskStatus.FAILED;
        }
        task.stage = "已完成";
        task.failureSummary = task.errors.isEmpty() ? null : String.join("; ", task.errors);
        persist(task);
    }

    private int countAttachments(List<com.example.officeconvert.model.AttachmentInfo> list) {
        int n = list.size();
        for (com.example.officeconvert.model.AttachmentInfo a : list) n += countAttachments(a.children);
        return n;
    }

    // ---------------- 查询 / 取消 / 重试 ----------------

    public ConversionTask get(String taskId) {
        return tasks.get(taskId);
    }

    public List<ConversionTask> list() {
        return tasks.values().stream()
                .sorted(java.util.Comparator.comparing(t -> t.createdAt)).toList();
    }

    public boolean cancel(String taskId) {
        ConversionTask t = tasks.get(taskId);
        if (t == null) return false;
        if (t.status == TaskStatus.PENDING || t.status == TaskStatus.PROCESSING) {
            Future<?> f = futures.get(taskId);
            if (f != null) f.cancel(true);
            t.status = TaskStatus.CANCELLED;
            t.finishedAt = java.time.Instant.now();
            t.stage = "已取消";
            persist(t);
            return true;
        }
        return false;
    }

    public synchronized ConversionTask retry(String taskId, TaskOptions newOptions) throws IOException {
        ConversionTask t = tasks.get(taskId);
        if (t == null) throw new IllegalArgumentException("任务不存在: " + taskId);
        if (t.status == TaskStatus.PENDING || t.status == TaskStatus.PROCESSING) {
            throw new IllegalArgumentException("任务正在处理中，无法重试");
        }
        Path taskDir = taskRoot().resolve(taskId);
        // 清理旧产物，避免重试产生重复附件（需求 11.2）
        Path output = taskDir.resolve("output");
        if (Files.exists(output)) deleteRecursively(output);
        t.resultDirs.clear();
        t.warnings.clear();
        t.errors.clear();
        t.failureSummary = null;
        t.processedDocs = 0;
        t.processedAttachments = 0;
        t.createdAt = java.time.Instant.now();
        t.startedAt = null;
        t.finishedAt = null;
        if (newOptions != null) t.options = newOptions;
        submit(t);
        return t;
    }

    // ---------------- 持久化 ----------------

    private void persist(ConversionTask task) {
        try {
            Path dir = taskRoot().resolve(task.taskId);
            Files.createDirectories(dir);
            mapper.writeValue(dir.resolve("task.json").toFile(), task);
        } catch (Exception ignore) { }
    }

    private String hashOf(MultipartFile f) {
        try {
            MessageDigest md = MessageDigest.getInstance("SHA-256");
            return HexFormat.of().formatHex(md.digest(f.getBytes()));
        } catch (Exception e) {
            return String.valueOf(f.getSize());
        }
    }

    private String optionsKey(TaskOptions o) {
        if (o == null) return "default";
        return String.valueOf(com.example.officeconvert.util.Hashes.sha256Utf8(
                String.valueOf(o.summary(props))));
    }

    private void deleteRecursively(Path path) throws IOException {
        if (!Files.exists(path)) return;
        try (var walk = Files.walk(path)) {
            walk.sorted(java.util.Comparator.reverseOrder()).forEach(p -> {
                try { Files.delete(p); } catch (IOException ignore) { }
            });
        }
    }
}
