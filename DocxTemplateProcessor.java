package com.cn.mdhw.utils;

import cn.hutool.core.util.ReUtil;
import cn.hutool.core.util.StrUtil;
import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.util.Map;

public class DocxTemplateProcessor {
    public static void replaceVariablesInDocx(String templatePath, String outputPath, Map<String, String> data) {
        try (FileInputStream fis = new FileInputStream(templatePath);
             XWPFDocument doc = new XWPFDocument(fis)) {
            // 替换段落中的占位符
            replaceInParagraph(doc, data);
            // 替换表格中的占位符
            for (XWPFTable table : doc.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph p : cell.getParagraphs()) {
                            replaceCoreCode(p, data);
                        }
                    }
                }
            }
            // 保存替换后的文档
            try (FileOutputStream out = new FileOutputStream(outputPath)) {
                doc.write(out);
            }
        } catch (IOException e) {
            throw new RuntimeException("替换文档变量失败", e);
        }
    }

    /**
     * 替换段落中的变量
     * @param p 段落
     * @param data key - value  key：是占位符(默认以${...}的形式)，value：是替换值
     */
    private static void replaceCoreCode(XWPFParagraph p, Map<String, String> data) {
        String paragraphText = p.getParagraphText();
        //regex ${.*?}
        String regex = "\\$\\{.*?\\}";
        boolean contains = ReUtil.contains(regex, paragraphText);
        if(contains){
            StringBuilder sb = new StringBuilder();
            boolean pushFlag = false;
            for (XWPFRun run : p.getRuns()) {
                String text = run.getText(0);
                if (StrUtil.isEmpty(text)) {
                    continue;
                }
                if (text.contains("$") || pushFlag) {
                    pushFlag = true;
                    sb.append(text);
                    run.setText("", 0);
                }
                if (text.contains("}")) {
                    String rowStr = sb.toString();
                    for (Map.Entry<String, String> entry : data.entrySet()) {
                        rowStr = rowStr.replace(entry.getKey(), entry.getValue());
                    }
                    run.setText(rowStr, 0);
                    pushFlag = false;
                }
            }
        }
    }
    private static void replaceInParagraph(XWPFDocument doc , Map<String, String> data) {
        for (XWPFParagraph p : doc.getParagraphs()) {
            replaceCoreCode(p, data);
        }
    }

    /**
     * 将docx转为pdf
     * @param docxPath 要转换的docx路径
     * @param pdfPath 转换后的pdf路径
     * @return
     */
    public static File toPDF(String docxPath, String pdfPath) {
        try {
            //新建一个pdf文档
            File file = new File(pdfPath);
            FileOutputStream os = new FileOutputStream(file);

            Document doc = new Document(docxPath);
            //全面支持DOC, DOCX, OOXML, RTF HTML, OpenDocument, PDF, EPUB, XPS, SWF 相互转换
            doc.save(os, SaveFormat.PDF);
            os.close();
            return file;
        } catch (Exception e) {
            throw new RuntimeException("Word 转 Pdf 失败...", e);
        }
    }


}