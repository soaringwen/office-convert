package com.example.officeconvert.convert;

import com.example.officeconvert.config.ConversionProperties;
import com.example.officeconvert.model.ConversionException;
import com.example.officeconvert.model.ConversionStatus;

import java.time.Instant;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * 单任务转换上下文：携带截止时间、资源限制计数器与重复附件检测（需求 7.7）。
 */
public class ConvCtx {

    public final ConversionProperties props;
    public final TaskOptions opts;
    public final Instant deadline;
    /** 任务级取消标记。 */
    public final java.util.concurrent.atomic.AtomicBoolean cancelled =
            new java.util.concurrent.atomic.AtomicBoolean(false);
    /** 附件全局编号序列（att-001, att-002 ...）。 */
    public final AtomicInteger idSeq = new AtomicInteger();
    public final AtomicInteger attachmentsUsed = new AtomicInteger();
    /** sha256 -> 已存在附件 id，用于重复附件检测。 */
    public final Map<String, String> seenHashes = new java.util.concurrent.ConcurrentHashMap<>();
    /** sha256 -> 首次出现的原始文件相对路径。 */
    public final Map<String, String> seenPaths = new java.util.concurrent.ConcurrentHashMap<>();
    public final Set<String> inProgressHashes = java.util.concurrent.ConcurrentHashMap.newKeySet();

    public ConvCtx(ConversionProperties props, TaskOptions opts, Instant deadline) {
        this.props = props;
        this.opts = opts;
        this.deadline = deadline;
    }

    /** 超时 / 取消检查点。 */
    public void checkDeadline() {
        if (cancelled.get()) {
            throw new ConversionException(ConversionStatus.FAILED, "任务已取消");
        }
        if (deadline != null && Instant.now().isAfter(deadline)) {
            throw new ConversionException(ConversionStatus.LIMIT_EXCEEDED, "任务执行超时，已终止当前内容处理");
        }
    }

    /** 附件计数限制（需求 7.7：最大附件总数）。 */
    public boolean tryAcquireAttachment() {
        int used = attachmentsUsed.incrementAndGet();
        return used <= props.getMaxAttachments();
    }

    public void releaseAttachment() {
        attachmentsUsed.decrementAndGet();
    }

    /** 循环引用 / 重复检测：返回 true 表示此前已处理过同哈希内容。 */
    public boolean markSeen(String sha256) {
        return !inProgressHashes.add(sha256) || seenHashes.containsKey(sha256);
    }

    public void completeSeen(String sha256) {
        inProgressHashes.remove(sha256);
    }
}
