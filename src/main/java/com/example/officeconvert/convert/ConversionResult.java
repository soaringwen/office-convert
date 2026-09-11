package com.example.officeconvert.convert;

import java.util.ArrayList;
import java.util.List;

/** 单个文档（或附件）的一次转换结果。 */
public class ConversionResult {

    public final StringBuilder markdown = new StringBuilder();
    /** 资源文件（图片等），相对文档输出目录。 */
    public final List<Asset> assets = new ArrayList<>();
    /** 嵌入附件对象。 */
    public final List<EmbeddedObject> embedded = new ArrayList<>();
    /** 外部关联引用（未实际嵌入，需求 7.9）。 */
    public final List<ExternalRef> externals = new ArrayList<>();
    public final List<String> warnings = new ArrayList<>();
    /** 是否发生内容截断（需求 7.4.3：不得静默截断）。 */
    public boolean truncated;
    public String limitNote;
    /** SmartArt / 图表等降级说明。 */
    public final List<String> degradations = new ArrayList<>();

    public static class Asset {
        public String fileName;   // image-001.png
        public String relPath;    // assets/image-001.png
        public byte[] data;
        public boolean image;

        public Asset(String fileName, String relPath, byte[] data, boolean image) {
            this.fileName = fileName; this.relPath = relPath; this.data = data; this.image = image;
        }
    }

    public static class ExternalRef {
        public String label;
        public String uri;
        public String location;

        public ExternalRef(String label, String uri, String location) {
            this.label = label; this.uri = uri; this.location = location;
        }
    }
}
