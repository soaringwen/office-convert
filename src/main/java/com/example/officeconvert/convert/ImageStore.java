package com.example.officeconvert.convert;

import com.example.officeconvert.util.Hashes;
import com.example.officeconvert.util.SafeNames;

import java.util.HashMap;
import java.util.Map;

/**
 * 文档内图片资源写入器：按内容去重，命名 image-001.png 等，输出相对路径引用。
 */
public class ImageStore {

    private final ConversionResult result;
    private final Map<String, String> byHash = new HashMap<>();
    private final SafeNames.Deduper deduper = new SafeNames.Deduper();
    private int seq = 0;

    public ImageStore(ConversionResult result) {
        this.result = result;
    }

    /** 添加图片，返回相对路径（如 assets/image-001.png）。 */
    public synchronized String add(byte[] data, String ext, boolean image) {
        return add("image", data, ext, image);
    }

    /** 添加通用资源文件（如 Excel 补充 CSV），返回相对路径。 */
    public synchronized String add(String prefix, byte[] data, String ext, boolean image) {
        String hash = Hashes.sha256(data);
        String existing = byHash.get(hash);
        if (existing != null) return existing;
        seq++;
        String fileName = String.format("%s-%03d.%s", prefix, seq, ext == null ? "bin" : ext);
        fileName = deduper.dedupe(fileName);
        String relPath = "assets/" + fileName;
        result.assets.add(new ConversionResult.Asset(fileName, relPath, data, image));
        byHash.put(hash, relPath);
        return relPath;
    }
}
