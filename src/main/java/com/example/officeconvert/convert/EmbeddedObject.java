package com.example.officeconvert.convert;

/** 从文档中提取出的嵌入附件原始对象（需求 7.6）。 */
public class EmbeddedObject {

    public byte[] data;
    public String originalName;             // 可能为 null（原始名称不可用）
    public boolean nameAvailable;
    public String sourceLocation;           // 例如 "对象 #1"、"OLE 流 /ObjectPool/..."、"PDF 附件"
    public String suggestedExt;             // 无法确定名称时的建议扩展名

    public EmbeddedObject(byte[] data, String originalName, boolean nameAvailable,
                          String sourceLocation, String suggestedExt) {
        this.data = data;
        this.originalName = originalName;
        this.nameAvailable = nameAvailable;
        this.sourceLocation = sourceLocation;
        this.suggestedExt = suggestedExt;
    }
}
