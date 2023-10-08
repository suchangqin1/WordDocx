# 介绍

本项目主要是实现了docx文档的批注标记文本内容功能。

若想对doc文档进行批注, 可利用jacob在window环境下进行调用office或WPS进行doc转docx再进行批注，后续将提供次方案工程。



# 案例

可参考Demo.java 文件的案例实现



```Java
public static void main(String[] args) throws Exception {
        Map<String, String> commentMap = new HashMap<>(4);
        commentMap.put("基金", "不合法词汇");
        DocxDocument docxDocument = new DocxDocument("C:\\Users\\mocheng\\Desktop\\test\\test.docx", commentMap);
        // 会从源文档中先删除此作者的所有批注, 然后再添加本次匹配的批注
        docxDocument.setAuthor("mocheng");
        XWPFDocument document = docxDocument.execute();

        document.write(Files.newOutputStream(Paths.get("C:\\Users\\mocheng\\Desktop\\test\\test-11.docx")));
}
```

