package com.example.officeconvert.util;

import java.util.HashMap;
import java.util.HashSet;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

/**
 * 文件名安全处理（需求 7.11：避免非法字符、重名覆盖和路径穿越）。
 */
public final class SafeNames {
    private SafeNames() {}

    private static final Set<String> RESERVED = Set.of(
            "CON", "PRN", "AUX", "NUL", "COM1", "COM2", "COM3", "COM4",
            "LPT1", "LPT2", "LPT3");

    /** 清理非法字符，限制长度，避免路径穿越与 Windows 保留名。 */
    public static String sanitize(String name, String fallback) {
        if (name == null) name = fallback;
        name = name.replace("\\", "/");
        int slash = name.lastIndexOf('/');
        if (slash >= 0) name = name.substring(slash + 1);
        // 去掉控制字符与非法字符
        StringBuilder sb = new StringBuilder();
        for (char c : name.toCharArray()) {
            if (c < 0x20 || "\\/:*?\"<>|".indexOf(c) >= 0) sb.append('_');
            else sb.append(c);
        }
        name = sb.toString().trim();
        if (name.isEmpty() || name.equals(".") || name.equals("..")) {
            name = fallback != null ? fallback : "file";
        }
        // 去掉结尾的点（Windows 兼容）
        while (name.endsWith(".")) name = name.substring(0, name.length() - 1);
        String base = name;
        String ext = "";
        int dot = name.lastIndexOf('.');
        if (dot > 0) { base = name.substring(0, dot); ext = name.substring(dot); }
        if (base.length() > 80) base = base.substring(0, 80);
        if (RESERVED.contains(base.toUpperCase(Locale.ROOT))) base = "_" + base;
        return base + ext;
    }

    /** 对一批名字做去重：冲突时追加 -2、-3 …，保证不覆盖（需求 14.2.5）。 */
    public static class Deduper {
        private final Map<String, Integer> used = new HashMap<>();
        private final Set<String> canonical = new HashSet<>();

        public synchronized String dedupe(String name) {
            Integer n = used.get(name.toLowerCase(Locale.ROOT));
            String candidate = name;
            if (n != null) {
                String base = name, ext = "";
                int dot = name.lastIndexOf('.');
                if (dot > 0) { base = name.substring(0, dot); ext = name.substring(dot); }
                do {
                    n++;
                    candidate = base + "-" + n + ext;
                } while (canonical.contains(candidate.toLowerCase(Locale.ROOT)));
            }
            used.put(candidate.toLowerCase(Locale.ROOT), countOf(candidate, used));
            canonical.add(candidate.toLowerCase(Locale.ROOT));
            return candidate;
        }

        private Integer countOf(String candidate, Map<String, Integer> used) {
            Integer existing = used.get(candidate.toLowerCase(Locale.ROOT));
            return existing == null ? 1 : existing;
        }
    }
}
