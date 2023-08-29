package com.mocheng.poi;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTComment;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTComments;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CommentsDocument;

import javax.xml.namespace.QName;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigInteger;
import java.util.*;

import static org.apache.poi.ooxml.POIXMLTypeLoader.DEFAULT_XML_OPTIONS;

/**
 * 自定义docx文档的批注对象, 对/word/comments.xml进行封装
 *
 * @author by suchangqin
 * @date 2023/7/18 17:36
 */
@Slf4j
public class DocxComments extends POIXMLDocumentPart {
    private final List<BigInteger> CLEAR_COMMENT_ID_LIST = new ArrayList<>();

    private String author = "robot";
    private CTComments comments;
    private BigInteger maxCommentId;


    public DocxComments(PackagePart part) {
        super(part);
        try {
            comments = CommentsDocument.Factory.parse(part.getInputStream(), DEFAULT_XML_OPTIONS).getComments();
        } catch (Exception ex) {
            log.warn("从源文件没有获取到comments, 将生成comments");
        }
        if (comments == null) {
            comments = CommentsDocument.Factory.newInstance().addNewComments();
        }

        maxCommentId = getMaxCommentId();
    }

    public CTComments getComments() {
        return comments;
    }

    /**
     * 获取已删除的批注内容的ID集合
     */
    public List<BigInteger> getClearCommentIdList() {
        return CLEAR_COMMENT_ID_LIST;
    }

    /**
     * 删除指定作者的批注内容
     *
     */
    public void clearComment() {
        for (int i = 0; i < comments.getCommentList().size(); i++) {
            CTComment comment = comments.getCommentArray(i);
            if (Objects.equals(comment.getAuthor(), author)) {
                CLEAR_COMMENT_ID_LIST.add(comment.getId());
                comments.removeComment(i--);
                continue;
            }
            if (author == null) {
                CLEAR_COMMENT_ID_LIST.add(comment.getId());
                comments.removeComment(i--);
            }
        }
    }

    /**
     * 新增批注内容
     *
     * @param text 批注的文本
     * @return 新增的批注ID
     */
    public BigInteger createComment(String text) {
        // 维护最多批注ID , 避免重复遍历获取
        maxCommentId = maxCommentId.add(BigInteger.ONE);
        CTComment ctComment = comments.addNewComment();
        ctComment.setAuthor(author);
        ctComment.setInitials("");
        ctComment.setDate(new GregorianCalendar(Locale.CHINA));
        ctComment.addNewP().addNewR().addNewT().setStringValue(text);
        ctComment.setId(maxCommentId);
        return maxCommentId;
    }

    /**
     * 获取当前最多的批注ID
     */
    public BigInteger getMaxCommentId() {
        BigInteger cId = BigInteger.ZERO;
        for (CTComment ctComment : comments.getCommentList()) {
            if (ctComment.getId().compareTo(cId) > 0) {
                cId = ctComment.getId();
            }
        }
        return cId;
    }

    @Override
    protected void commit() throws IOException {
        XmlOptions xmlOptions = new XmlOptions(DEFAULT_XML_OPTIONS);
        xmlOptions.setSaveSyntheticDocumentElement(new QName(CTComments.type.getName().getNamespaceURI(), "comments"));
        PackagePart part = getPackagePart();
        OutputStream out = part.getOutputStream();
        comments.save(out, xmlOptions);
        out.close();
    }

    public void setAuthor(String author) {
        this.author = author;
    }
}
