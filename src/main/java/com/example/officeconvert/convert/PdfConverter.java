package com.example.officeconvert.convert;

import com.example.officeconvert.model.ConversionException;
import com.example.officeconvert.model.ConversionStatus;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentCatalog;
import org.apache.pdfbox.pdmodel.PDDocumentNameDictionary;
import org.apache.pdfbox.pdmodel.PDEmbeddedFilesNameTreeNode;
import org.apache.pdfbox.pdmodel.common.COSObjectable;
import org.apache.pdfbox.pdmodel.common.filespecification.PDComplexFileSpecification;
import org.apache.pdfbox.pdmodel.common.filespecification.PDEmbeddedFile;
import org.apache.pdfbox.text.PDFTextStripper;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Map;

/**
 * PDF 转 Markdown（每页一个二级标题），并提取 PDF 内嵌附件（需求 7.6）。
 */
@Component
public class PdfConverter implements DocumentConverter {

    @Override
    public boolean supports(String ext) {
        return "pdf".equals(ext);
    }

    @Override
    public String name() {
        return "pdfbox-pdf";
    }

    @Override
    public ConversionResult convert(Path file, TaskOptions opts, ConvCtx ctx) throws Exception {
        ConversionResult res = new ConversionResult();
        try (InputStream in = Files.newInputStream(file); PDDocument doc = PDDocument.load(in)) {
            if (doc.isEncrypted()) {
                throw new ConversionException(ConversionStatus.ENCRYPTED, "PDF 已加密，系统不尝试破解");
            }
            int pages = doc.getNumberOfPages();
            PDFTextStripper stripper = new PDFTextStripper();
            for (int p = 1; p <= pages; p++) {
                ctx.checkDeadline();
                stripper.setStartPage(p);
                stripper.setEndPage(p);
                String text = stripper.getText(doc);
                res.markdown.append("## 第 ").append(p).append(" 页\n\n");
                if (text != null && !text.isBlank()) {
                    res.markdown.append(text.trim()).append("\n\n");
                } else {
                    res.markdown.append("（本页无可提取文本，可能为扫描图片页，OCR 未启用）\n\n");
                    res.warnings.add("第 " + p + " 页未提取到文本（可能为图片页；OCR 按业务确认不启用）");
                }
            }
            extractPdfAttachments(doc, res);
        }
        return res;
    }

    private void extractPdfAttachments(PDDocument doc, ConversionResult res) {
        try {
            PDDocumentCatalog catalog = doc.getDocumentCatalog();
            PDDocumentNameDictionary names = catalog.getNames();
            if (names == null) return;
            PDEmbeddedFilesNameTreeNode tree = names.getEmbeddedFiles();
            if (tree == null) return;
            Map<String, PDComplexFileSpecification> map = tree.getNames();
            if (map == null) return;
            int n = 0;
            for (Map.Entry<String, PDComplexFileSpecification> e : map.entrySet()) {
                PDComplexFileSpecification fs = e.getValue();
                PDEmbeddedFile ef = fs.getEmbeddedFile();
                if (ef == null) continue;
                byte[] data;
                try (InputStream is = ef.createInputStream()) {
                    data = is.readAllBytes();
                }
                String name = fs.getFile() != null ? fs.getFile() : e.getKey();
                res.embedded.add(new EmbeddedObject(data, name, true, "PDF 内嵌附件 #" + (++n), null));
            }
        } catch (Exception ex) {
            res.warnings.add("PDF 内嵌附件提取异常: " + ex.getMessage());
        }
    }
}
