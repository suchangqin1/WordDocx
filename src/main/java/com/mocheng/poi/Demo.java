package com.mocheng.poi;

import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;

/**
 * <TODO:一句话概述>
 *
 * @author by suchangqin
 * @date 2023/8/25 15:54
 */
public class Demo {

    public static void main(String[] args) throws Exception {
        Map<String, String> commentMap = new HashMap<>(4);
        commentMap.put("基金", "不合法词汇");
        DocxDocument docxDocument = new DocxDocument("C:\\Users\\mocheng\\Desktop\\test\\基金销售系统.docx", commentMap);
        docxDocument.setAuthor("mocheng");
        XWPFDocument document = docxDocument.execute();

        document.write(Files.newOutputStream(Paths.get("C:\\Users\\mocheng\\Desktop\\test\\基金销售系统-11.docx")));
    }
}
