package com.scq.poi.utils;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.w3c.dom.Node;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;

/**
 * POI处理docx文件的工具类
 *
 * @author by suchangqin
 */
@Slf4j
public class DocxHelper {

    private static final String ALL_CHILD_NODE = "./*";
    private static final String COMMENT_START_PART = "commentRangeStart";
    private static final String COMMENT_END_PART = "commentRangeEnd";
    private static final String COMMENT_END_XML = "w:commentRangeEnd";
    private static final String COMMENT_START_XML = "w:commentRangeStart";
    private static final String ID_XML = "w:id";

    /**
     * 获取docx文件某一个段落中的所有图片的id
     *
     * @param paragraph 当前段落
     * @return 段落中的所有图片id , 通过xwpfDocument.getPictureDataByID(pictureId) 获取到图片
     */
    public static List<String> readImageInParagraph(XWPFParagraph paragraph) {
        //图片索引List
        List<String> imageBundleList = new ArrayList<>();
        //段落中所有XWPFRun
        List<XWPFRun> runList = paragraph.getRuns();
        for (XWPFRun run : runList) {
            imageBundleList.addAll(getImageInRun(run));
        }
        return imageBundleList;
    }

    /**
     * 获取docx文件某一个run中的所有图片的id
     *
     * @param run 当前XWPFRun对象
     * @return run中的所有图片id , 通过xwpfDocument.getPictureDataByID(pictureId) 获取到图片
     */
    public static List<String> getImageInRun(XWPFRun run) {
        //图片索引List
        List<String> imageIdList = new ArrayList<>();
        //XWPFRun是POI对xml元素解析后生成的自己的属性，无法通过xml解析，需要先转化成CTR
        CTR ctr = run.getCTR();

        //获取光标: 对子元素进行遍历
        XmlCursor c = ctr.newCursor();
        //这个就是拿到所有的子元素：
        c.selectPath(ALL_CHILD_NODE);
        while (c.toNextSelection()) {
            XmlObject o = c.getObject();
            //如果子元素是<w:drawing>这样的形式，使用CTDrawing保存图片
            if (o instanceof CTDrawing) {
                CTDrawing drawing = (CTDrawing) o;
                for (CTAnchor anchor : drawing.getAnchorList()) {
                    XmlCursor cursor = anchor.getGraphic().getGraphicData().newCursor();
                    getImageId(imageIdList, cursor);
                }

                for (CTInline ctInline : drawing.getInlineList()) {
                    XmlCursor cursor = ctInline.getGraphic().getGraphicData().newCursor();
                    getImageId(imageIdList, cursor);
                }
            }
        }

        return imageIdList;
    }

    private static void getImageId(List<String> imageIdList, XmlCursor cursor) {
        cursor.selectPath(ALL_CHILD_NODE);
        while (cursor.toNextSelection()) {
            // 如果子元素是<pic:pic>这样的形式
            XmlObject xmlObject = cursor.getObject();
            if (xmlObject instanceof CTPicture) {
                CTPicture picture = (CTPicture) xmlObject;
                //拿到元素的属性
                imageIdList.add(picture.getBlipFill().getBlip().getEmbed());
            }
        }
    }


    public static void insertCommentRangeToRun(XWPFRun run, boolean start, BigInteger commentId) {
        String uri = CTMarkupRange.type.getName().getNamespaceURI();
        String localPart;
        XmlCursor cursor = run.getCTR().newCursor();
        if (start) {
            // 批注的开始标签名, org.openxmlformats.schemas.wordprocessingml.x2006.main.impl.CTRImpl.PROPERTY_QNAME
            localPart = COMMENT_START_PART;
        } else {
            if (!cursor.toNextSibling()) {
                // 如果没有下一个兄弟节点 , 则跳到父节点的下一个兄弟节点添加结束标签
                cursor.toParent();
                cursor.toNextSibling();
            }
            // 批注的结束标签名
            localPart = COMMENT_END_PART;
        }
        cursor.beginElement(localPart, uri);

        cursor.toParent();

        CTMarkupRange markup = (CTMarkupRange) cursor.getObject();
        cursor.dispose();
        markup.setId(commentId);
    }

    /**
     * 递归: 删除指定批注的范围标签, 删除当前run的父级下的所有匹配的,
     * 使用默认方法:
     * paragraph.getCTP().getCommentRangeStartList().removeIf(p);
     * paragraph.getCTP().getCommentRangeEndList().removeIf(p);
     * 无法在WPS编辑内容的批注标签 , 集合上没有
     *
     * @param run                当前run
     * @param clearCommentIdList 指定的批注的id集
     */
    public static void clearRunCommentStartAndEndXml(XWPFRun run, List<BigInteger> clearCommentIdList) {
        XmlCursor oldCursor = run.getCTR().newCursor();
        oldCursor.toParent();
        oldCursor.toFirstChild();
        removeCommentXml(clearCommentIdList, oldCursor);
    }

    private static void removeCommentXml(List<BigInteger> clearCommentIdList, XmlCursor oldCursor) {
        Node node = oldCursor.getDomNode();
        boolean delete = false;
        if (COMMENT_START_XML.equals(node.getNodeName()) || COMMENT_END_XML.equals(node.getNodeName())) {
            String id = getAttribute(node, ID_XML);
            if (clearCommentIdList.contains(new BigInteger(id))) {
                oldCursor.removeXml();
                delete = true;
            }

            if (delete || oldCursor.toNextSibling()) {
                // 删除后光标会到下一个节点
                removeCommentXml(clearCommentIdList, oldCursor);
            }

        } else {
            if (oldCursor.toNextSibling()) {
                removeCommentXml(clearCommentIdList, oldCursor);
            }
        }
    }


    /***
     * 获取当前标签, 指定属性的值
     *
     * @param node 当前的Node
     * @param attName 要获取的属性名
     * @return 属性值, 没有该属性时返回null
     */
    public static String getAttribute(Node node, String attName) {
        return (node.hasAttributes() && node.getAttributes().getNamedItem(
                attName) != null) ? node.getAttributes().getNamedItem(attName)
                .getNodeValue() : null;
    }

    public static void copyStyle(XWPFRun fromRun, XWPFRun toRun) {
        CTR source = fromRun.getCTR();
        CTRPr rPrSource = source.getRPr();
        if (rPrSource != null) {
            CTRPr rPrDest = (CTRPr) rPrSource.copy();
            CTR target = toRun.getCTR();
            target.setRPr(rPrDest);
        }
    }

}