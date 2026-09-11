package com.example.officeconvert.util;

/** Markdown 生成辅助工具。 */
public final class Md {
    private Md() {}

    /** 表格单元格内容转义竖线，并将换行转换为 <br>。 */
    public static String cell(String text) {
        if (text == null) return "";
        return text.replace("|", "\\|").replace("\r\n", "<br>")
                .replace("\n", "<br>").replace("\r", "<br>").trim();
    }

    public static String heading(int level, String text) {
        return "#".repeat(Math.max(1, Math.min(6, level))) + " " + (text == null ? "" : text.trim());
    }

    public static boolean isTrue(Boolean v, boolean def) {
        return v == null ? def : v;
    }
}
