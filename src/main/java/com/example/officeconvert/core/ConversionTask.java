package com.example.officeconvert.core;

import com.example.officeconvert.model.TaskStatus;

import java.time.Instant;
import java.util.ArrayList;
import java.util.List;

/** 转换任务（可序列化为 task.json 用于重启恢复，需求 11.2）。 */
public class ConversionTask {

    public String taskId;
    public TaskStatus status = TaskStatus.PENDING;
    public com.example.officeconvert.convert.TaskOptions options = new com.example.officeconvert.convert.TaskOptions();

    public List<DocInput> docs = new ArrayList<>();
    public int totalDocs;
    public int processedDocs;
    public int processedAttachments;
    public String stage = "等待处理";

    public List<String> warnings = new ArrayList<>();
    public List<String> errors = new ArrayList<>();
    public List<String> resultDirs = new ArrayList<>();
    public String failureSummary;

    public Instant createdAt = Instant.now();
    public Instant startedAt;
    public Instant finishedAt;

    /** 幂等键：首文件哈希 + 配置哈希（需求 7.1）。 */
    public String idempotencyKey;

    public static class DocInput {
        public String storedPath;      // 相对 taskDir 的输入副本路径
        public String displayName;

        public DocInput() {}
        public DocInput(String storedPath, String displayName) {
            this.storedPath = storedPath;
            this.displayName = displayName;
        }
    }
}
