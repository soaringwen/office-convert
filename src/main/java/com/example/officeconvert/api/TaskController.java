package com.example.officeconvert.api;

import com.example.officeconvert.convert.TaskOptions;
import com.example.officeconvert.core.ConversionTask;
import com.example.officeconvert.core.TaskManager;
import com.example.officeconvert.model.TaskStatus;
import org.springframework.core.io.FileSystemResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * 转换任务接口（需求 12）。
 */
@RestController
@RequestMapping("/api/v1/conversion-tasks")
public class TaskController {

    private final TaskManager taskManager;

    public TaskController(TaskManager taskManager) {
        this.taskManager = taskManager;
    }

    /** 12.1 创建转换任务（单文件或多文件）。 */
    @PostMapping
    public ResponseEntity<?> create(@RequestParam("files") List<MultipartFile> files,
                                    TaskOptions options) throws IOException {
        try {
            ConversionTask task = taskManager.create(files, options);
            return ResponseEntity.status(HttpStatus.ACCEPTED).body(view(task));
        } catch (IllegalArgumentException e) {
            return ResponseEntity.badRequest().body(Map.of("code", "INVALID_REQUEST", "message", e.getMessage()));
        }
    }

    /** 任务列表。 */
    @GetMapping
    public List<Map<String, Object>> list() {
        List<Map<String, Object>> out = new ArrayList<>();
        for (ConversionTask t : taskManager.list()) out.add(view(t));
        return out;
    }

    /** 12.2 查询任务状态。 */
    @GetMapping("/{taskId}")
    public ResponseEntity<?> status(@PathVariable String taskId) {
        ConversionTask t = taskManager.get(taskId);
        if (t == null) return notFound();
        return ResponseEntity.ok(view(t));
    }

    /** 12.3 获取转换结果（清单 + 报告 + 产物路径）。 */
    @GetMapping("/{taskId}/result")
    public ResponseEntity<?> result(@PathVariable String taskId) throws IOException {
        ConversionTask t = taskManager.get(taskId);
        if (t == null) return notFound();
        if (t.status == TaskStatus.PENDING || t.status == TaskStatus.PROCESSING) {
            return ResponseEntity.status(HttpStatus.CONFLICT)
                    .body(Map.of("code", "NOT_FINISHED", "message", "任务尚未完成"));
        }
        Path taskDir = taskManager.taskRoot().resolve(taskId);
        List<Map<String, Object>> docs = new ArrayList<>();
        for (String rel : t.resultDirs) {
            Path docDir = taskDir.resolve(rel);
            Map<String, Object> d = new LinkedHashMap<>();
            d.put("name", rel);
            d.put("contentMd", Files.exists(docDir.resolve("content.md")) ? rel + "/content.md" : null);
            d.put("manifest", Files.exists(docDir.resolve("manifest.json")) ? rel + "/manifest.json" : null);
            d.put("report", Files.exists(docDir.resolve("conversion-report.json"))
                    ? rel + "/conversion-report.json" : null);
            docs.add(d);
        }
        Map<String, Object> body = new LinkedHashMap<>();
        body.put("taskId", taskId);
        body.put("status", t.status.name().toLowerCase());
        body.put("documents", docs);
        body.put("warnings", t.warnings);
        body.put("errors", t.errors);
        body.put("zip", "/api/v1/conversion-tasks/" + taskId + "/result/zip");
        return ResponseEntity.ok(body);
    }

    /** 主文档 Markdown 内容。 */
    @GetMapping("/{taskId}/result/markdown")
    public ResponseEntity<?> markdown(@PathVariable String taskId) throws IOException {
        ConversionTask t = taskManager.get(taskId);
        if (t == null) return ResponseEntity.notFound().build();
        if (t.resultDirs.isEmpty()) return ResponseEntity.notFound().build();
        Path md = taskManager.taskRoot().resolve(taskId).resolve(t.resultDirs.get(0)).resolve("content.md");
        if (!Files.exists(md)) return ResponseEntity.notFound().build();
        return ResponseEntity.ok().contentType(MediaType.parseMediaType("text/markdown; charset=UTF-8"))
                .body(Files.readString(md, StandardCharsets.UTF_8));
    }

    /** 结果 ZIP 包下载（相对路径结构，可整体迁移，需求 17.6）。 */
    @GetMapping("/{taskId}/result/zip")
    public ResponseEntity<StreamingResponseBody> zip(@PathVariable String taskId) throws IOException {
        ConversionTask t = taskManager.get(taskId);
        if (t == null) return ResponseEntity.notFound().build();
        Path taskDir = taskManager.taskRoot().resolve(taskId).resolve("output");
        if (!Files.exists(taskDir)) return ResponseEntity.notFound().build();
        String fileName = "attachment; filename*=UTF-8''"
                + URLEncoder.encode(taskId + "-result.zip", StandardCharsets.UTF_8);
        StreamingResponseBody body = out -> {
            try (ZipOutputStream zos = new ZipOutputStream(out)) {
                try (var walk = Files.walk(taskDir)) {
                    for (Path p : (Iterable<Path>) walk::iterator) {
                        if (Files.isDirectory(p)) continue;
                        String entry = taskDir.relativize(p).toString().replace('\\', '/');
                        zos.putNextEntry(new ZipEntry(entry));
                        Files.copy(p, zos);
                        zos.closeEntry();
                    }
                }
            }
        };
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, fileName)
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(body);
    }

    /** 产物文件下载（含路径穿越防护，需求 10.1）。 */
    @GetMapping("/{taskId}/files/**")
    public ResponseEntity<FileSystemResource> file(@PathVariable String taskId,
                                                   jakarta.servlet.http.HttpServletRequest request) {
        ConversionTask t = taskManager.get(taskId);
        if (t == null) return ResponseEntity.notFound().build();
        String prefix = "/api/v1/conversion-tasks/" + taskId + "/files/";
        String uri = request.getRequestURI();
        int idx = uri.indexOf(prefix);
        if (idx < 0) return ResponseEntity.notFound().build();
        String rel = uri.substring(idx + prefix.length());
        Path taskDir = taskManager.taskRoot().resolve(taskId).toAbsolutePath().normalize();
        Path target = taskDir.resolve(rel).normalize();
        if (!target.startsWith(taskDir) || !Files.exists(target) || Files.isDirectory(target)) {
            return ResponseEntity.notFound().build();
        }
        return ResponseEntity.ok(new FileSystemResource(target));
    }

    /** 12.4 取消任务。 */
    @PostMapping("/{taskId}/cancel")
    public ResponseEntity<?> cancel(@PathVariable String taskId) {
        boolean ok = taskManager.cancel(taskId);
        if (!ok) {
            return ResponseEntity.badRequest()
                    .body(Map.of("code", "NOT_CANCELLABLE", "message", "任务不存在或已结束，无法取消"));
        }
        return ResponseEntity.ok(view(taskManager.get(taskId)));
    }

    /** 12.4 重试任务（可携带新配置重新转换，需求 12.4）。 */
    @PostMapping("/{taskId}/retry")
    public ResponseEntity<?> retry(@PathVariable String taskId, TaskOptions options) throws IOException {
        try {
            ConversionTask task = taskManager.retry(taskId, options);
            return ResponseEntity.accepted().body(view(task));
        } catch (IllegalArgumentException e) {
            return ResponseEntity.badRequest().body(Map.of("code", "INVALID_REQUEST", "message", e.getMessage()));
        }
    }

    // ---------------- 视图 ----------------

    private Map<String, Object> view(ConversionTask t) {
        Map<String, Object> m = new LinkedHashMap<>();
        m.put("taskId", t.taskId);
        m.put("status", t.status.name().toLowerCase());
        m.put("stage", t.stage);
        m.put("progress", t.totalDocs == 0 ? 0 : Math.round(100.0 * t.processedDocs / t.totalDocs));
        m.put("totalDocs", t.totalDocs);
        m.put("processedDocs", t.processedDocs);
        m.put("processedAttachments", t.processedAttachments);
        m.put("warningCount", t.warnings.size());
        m.put("errorCount", t.errors.size());
        m.put("createdAt", t.createdAt == null ? null : t.createdAt.toString());
        m.put("startedAt", t.startedAt == null ? null : t.startedAt.toString());
        m.put("finishedAt", t.finishedAt == null ? null : t.finishedAt.toString());
        return m;
    }

    private ResponseEntity<Object> notFound() {
        return ResponseEntity.status(HttpStatus.NOT_FOUND)
                .body(Map.of("code", "NOT_FOUND", "message", "任务不存在"));
    }
}
