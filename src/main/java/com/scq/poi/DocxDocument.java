package com.scq.poi;

import com.scq.poi.algorithm.AcMatchUtils;
import com.scq.poi.algorithm.AcTree;
import com.scq.poi.utils.DocxHelper;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTMarkup;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.springframework.util.CollectionUtils;
import org.springframework.util.StringUtils;

import java.io.IOException;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.function.Predicate;

/**
 * 自定义word docx文档处理类
 *
 * @author by suchangqin
 * @date 2023/7/18 17:45
 */
public class DocxDocument {

    private static final String COMMENTS_XML_PATH = "/word/comments.xml";
    private static final String WORD_COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";

    private final XWPFDocument document;
    private final Map<String, String> addCommentMap;
    private AcMatchUtils acMatchUtils;
    private DocxComments docxComments;

    /**
     * 创建 自定义word docx文档处理对象
     *
     * @param filePath      docx文档文件路径 , 必须是docx文档且后缀名是docx
     * @param addCommentMap 新增批注的内容, key: 批注所标记的文本, value: 批注的内容。
     *                      如：
     *                      文档的文本：“这是一个自定义word docx文档处理类” ,
     *                      需要再此文本内容的 "自定义" 词语标记批注内容 "标记此内容, 原因是违规",
     *                      则 key就是 "自定义", value就是 "标记此内容, 原因是违规"
     * @throws IOException 异常
     */
    public DocxDocument(String filePath, Map<String, String> addCommentMap) throws Exception {
        this.document = new XWPFDocument(Files.newInputStream(Paths.get(filePath)));
        this.addCommentMap = addCommentMap;
        createAcTree();
        createDocxComment();
    }


    /**
     * 基于AC自动机算法, 构建算法树
     */
    private void createAcTree() {
        ArrayList<String> list = new ArrayList<>(addCommentMap.keySet());
        this.acMatchUtils = new AcMatchUtils(new AcTree(list));
    }

    /**
     * 获取已有批注内容
     */
    private void createDocxComment() throws InvalidFormatException {
        DocxComments docxComments = null;

        // 先尝试去读取原文件中的comments 批注内容
        for (POIXMLDocumentPart.RelationPart relationPart : document.getRelationParts()) {
            String relation = relationPart.getRelationship().getRelationshipType();
            if (relation.equals(XWPFRelation.COMMENT.getRelation())) {
                POIXMLDocumentPart part = relationPart.getDocumentPart();
                docxComments = new DocxComments(part.getPackagePart());
                String rId = document.getRelationId(part);
                document.addRelation(rId, XWPFRelation.COMMENT, docxComments);
            }
        }

        // 如果没有批注内容 , 则新建一个
        if (docxComments == null) {
            OPCPackage opcPackage = document.getPackage();
            PackagePartName partName = PackagingURIHelper.createPartName(COMMENTS_XML_PATH);
            PackagePart part = opcPackage.createPart(partName, WORD_COMMENTS_CONTENT_TYPE);
            docxComments = new DocxComments(part);
            document.addRelation(null, XWPFRelation.COMMENT, docxComments);
        }

        this.docxComments = docxComments;
    }

    /**
     * 设置本次处理批注作者的内容,
     * 用于删除文档批注内容的作者为该值的批注 , 新生成的批注的作者名也为该值
     *
     * @param author 批注作者名
     */
    public void setAuthor(String author) {
        docxComments.setAuthor(author);
    }

    /**
     * 对文档增加指定批注批注
     *
     * @return 处理完后的文档
     */
    public XWPFDocument execute() {
        // 检查当前作者有没有批注过 , 有则删除当前作者的全部批注
        docxComments.clearComment();
        // 先处理表格
        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    for (XWPFParagraph paragraph : cell.getParagraphs()) {
                        dealDocxParagraph(paragraph);
                    }
                }
            }
        }
        // 再处理普通段落
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            dealDocxParagraph(paragraph);
        }

        return document;
    }

    /**
     * 对当前段落进行批注
     * 前提: 正文内容不会发生增或减
     * 原理: 1.先记录段落每个字符的索引所对应的run, 记录每个run所包含的字符索引, 记录每个run的索引
     * 2.然后获取不合法词汇的首尾字符在段落中的索引, 在此索引位置分割run , 刷新上一步的每个集合, 并保存run的标签前所设置的批注的开始和结束范围标签
     * 3.最后统一处理在对应的run中设置范围标签和批注引用
     *
     * @param paragraph 当前段落
     */
    private void dealDocxParagraph(XWPFParagraph paragraph) {
        // 当前段落每个字符索引索对应的run集合
        Map<Integer, XWPFRun> charRunMap = new HashMap<>((int) (paragraph.getText().length() / 0.75));
        // 当前段落每个run所包含的文字的全部索引, 索引集正序
        Map<XWPFRun, List<Integer>> runCharMap = new HashMap<>((int) (paragraph.getRuns().size() / 0.75));
        // 当前段落每个run所属段落的索引
        Map<XWPFRun, Integer> runMap = new HashMap<>((int) (paragraph.getRuns().size() / 0.75));

        // 批注标签的范围标签集合
        Map<XWPFRun, List<BigInteger>> commentRangeStartMap = new HashMap<>(16);
        Map<XWPFRun, List<BigInteger>> commentRangeEndMap = new HashMap<>(16);

        // 处理原始段落所有run的数据
        String paragraphText = dealAllSourceRunData(paragraph, charRunMap, runCharMap, runMap);
        // 智检
        Map<String, List<Integer>> matchList = acMatchUtils.match(paragraphText);

        for (Map.Entry<String, List<Integer>> entry : matchList.entrySet()) {
            String word = entry.getKey();
            for (Integer startIndex : entry.getValue()) {
                // 创建当前不合法词的批注
                BigInteger commentId = docxComments.createComment(addCommentMap.get(word));

                // -------------处理批注范围的开始标签-------------
                // 当前索引所属字符所属的run
                XWPFRun run = charRunMap.get(startIndex);
                // 当前run的文本字符原始索引集
                List<Integer> indexList = runCharMap.get(run);
                if (Objects.equals(startIndex, indexList.get(0))) {
                    // 新增当前敏感词批注范围的开始标签
                    addCommentIdToMap(commentRangeStartMap, commentId, run);
                } else {
                    XWPFRun newRun = splitRunOnIndex(paragraph, charRunMap, runCharMap, runMap, startIndex, run, indexList);
                    // 新增当前敏感词批注范围的开始标签
                    addCommentIdToMap(commentRangeStartMap, commentId, newRun);
                }

                // -------------处理批注范围的结束标签-------------
                // 当前批注文字的结束字符所在段落文本的索引
                int endIndex = startIndex + word.length() - 1;
                if (endIndex == paragraphText.length() - 1) {
                    paragraph.createRun();
                }
                XWPFRun endRun = charRunMap.get(endIndex);
                // 当前run的文本字符原始索引集
                indexList = runCharMap.get(endRun);
                if (!Objects.equals(endIndex, indexList.get(indexList.size() - 1))) {
                    // 在结束字符的下一个字符索引位置切割
                    splitRunOnIndex(paragraph, charRunMap, runCharMap, runMap, endIndex + 1, endRun, indexList);
                } else {
                    // 当前位置是旧的run的文本的结束位置, 如果旧的run有结束标签, 则调整结束标签的位置
                    if (commentRangeEndMap.containsKey(run)) {
                        for (int i = 0; i < commentRangeEndMap.get(run).size(); i++) {
                            addCommentIdToMap(commentRangeEndMap, commentRangeEndMap.get(run).get(i), endRun);
                            commentRangeEndMap.get(run).remove(i);
                        }
                    }
                }
                // 新增当前敏感词批注范围的结束标签
                addCommentIdToMap(commentRangeEndMap, commentId, endRun);
            }
        }

        // 开始统一处理批注的范围标签, 若在新增批注的遍历中同时新增范围标签, 可能会因为拆分run并在指定位置插入新run的时候导致范围标签位置错误
        addCommentLabel(commentRangeStartMap, true);
        addCommentLabel(commentRangeEndMap, false);
    }


    /**
     * 对段落中的原始run进行处理 , 并设置图片的批注
     *
     * @return 当前段落的文本
     */
    private String dealAllSourceRunData(XWPFParagraph paragraph, Map<Integer, XWPFRun> charRunMap,
                                        Map<XWPFRun, List<Integer>> runCharMap,
                                        Map<XWPFRun, Integer> runMap) {

        // 获取删除的批注内容
        List<BigInteger> clearCommentIdList = docxComments.getClearCommentIdList();
        // 过滤器, 用于删除对应批注的标签
        Predicate<CTMarkup> p = ctMarkup -> clearCommentIdList.contains(ctMarkup.getId());
        // 段落文本
        StringBuilder paragraphText = new StringBuilder(64);

        // 当前段落文本长度
        int length = 0;
        for (int i = 0; i < paragraph.getRuns().size(); i++) {
            XWPFRun run = paragraph.getRuns().get(i);
            if (!CollectionUtils.isEmpty(run.getCTR().getDelTextList())) {
                // 如果启用了审阅(修订)并且这是一次已删除的run，则不包括此run
                continue;
            }
            // WPS在线编辑生成的docx文档, 因在线WPS在线编辑插入的内容和其它内容不在同一级 , 每个run都遍历并清一遍父节点下的所有批注范围标签
            DocxHelper.clearRunCommentStartAndEndXml(run, clearCommentIdList);
            runMap.put(run, i);
            // 删除需要删除的批注的引用
            run.getCTR().getCommentReferenceList().removeIf(p);
            String text = run.text();
            if (!StringUtils.isEmpty(text)) {
                paragraphText.append(text);
                // 处理段落每个字符索引所对应的run
                for (int ch = 0; ch < text.length(); ch++) {
                    int index = ch + length;
                    charRunMap.put(index, run);
                    if (runCharMap.containsKey(run)) {
                        List<Integer> indexList = runCharMap.get(run);
                        indexList.add(index);
                    } else {
                        List<Integer> indexList = new ArrayList<>();
                        indexList.add(index);
                        runCharMap.put(run, indexList);
                    }
                }

                length += text.length();
            }

            //TODO 图片内容提取并检测批注
            // 获取run中的图片id
            /*List<String> imageInRunList = DocxUtils.getImageInRun(run);
            for (String blipId : imageInRunList) {
                XWPFPictureData pictureData = paragraph.getDocument().getPictureDataByID(blipId);
                byte[] data = pictureData.getData();
                File imageFile = new File(fileDirectory + File.separator + UUID.randomUUID() + ".jpg");
            }*/
        }

        return paragraphText.toString();
    }

    /**
     * 保存指定run的批注id
     */
    public void addCommentIdToMap(Map<XWPFRun, List<BigInteger>> commentRangeMap, BigInteger commentId, XWPFRun run) {
        if (commentRangeMap.containsKey(run)) {
            commentRangeMap.get(run).add(commentId);
        } else {
            List<BigInteger> commentRangeList = new ArrayList<>();
            commentRangeList.add(commentId);
            commentRangeMap.put(run, commentRangeList);
        }
    }

    /**
     * 在指定索引位置切割run
     *
     * @param index 所切割字符所在当前段落文本的索引位置, 在字符的左边切割
     */
    public XWPFRun splitRunOnIndex(XWPFParagraph paragraph, Map<Integer, XWPFRun> charRunMap, Map<XWPFRun, List<Integer>> runCharMap,
                                   Map<XWPFRun, Integer> runMap, Integer index, XWPFRun run, List<Integer> indexList) {
        // 当前需要处理的段落字符所属的run的字符的所在run文本的索引
        int runTextIndex = 0;
        // 新的原始run的文本字符索引集
        List<Integer> newSourceIndexList = new ArrayList<>();
        // 新的原始run的文本字符索引集
        List<Integer> newRunIndexList = new ArrayList<>();
        // 获取当前不合法词汇字符索引所在run文本的位置
        for (int i = 0; i < indexList.size(); i++) {
            Integer textCharIndex = indexList.get(i);
            if (textCharIndex.compareTo(index) < 0) {
                newSourceIndexList.add(textCharIndex);
                continue;
            }
            if (Objects.equals(textCharIndex, index)) {
                runTextIndex = i;
            }
            newRunIndexList.add(textCharIndex);
        }

        String runText = run.text();
        String sourceRunText = runText.substring(0, runTextIndex);
        // pos 为run中的非修订删除的文本标签集合的索引 , 此时无法找到<w:delText>(修订的文本)标签的内容
        run.setText(sourceRunText, 0);

        String newRunText = runText.substring(runTextIndex);

        int newRunIndex = runMap.get(run) + 1;
        // 在旧run的索引之后新增一个run, 保存切割后的新文本
        XWPFRun newRun;
        if (newRunIndex == runMap.size()) {
            // 要指定新增的索引位置已超出段落以后的run的索引( 旧的run已是当前段落的最后一个run )
            newRun = paragraph.createRun();
        } else {
            // 指定位置插入新的run
            int rSize = Optional.ofNullable(paragraph)
                    .map(XWPFParagraph::getCTP)
                    .map(CTP::getRList)
                    .map(List::size)
                    .orElse(0);
            if (rSize <= newRunIndex) {
                for (int i = 0; i <= newRunIndex - rSize; i++) {
                    paragraph.getCTP().addNewR();
                }
            }
            newRun = paragraph.insertNewRun(newRunIndex);
        }
        newRun.setText(newRunText);
        DocxHelper.copyStyle(run, newRun);
        // 更新字符索引集
        for (int i = sourceRunText.length(); i > 0; i--) {
            charRunMap.put(index - i, run);
        }
        for (int i = 0; i < newRunText.length(); i++) {
            charRunMap.put(index + i, newRun);
        }
        // 更新run的索引集
        runCharMap.put(run, newSourceIndexList);
        runCharMap.put(newRun, newRunIndexList);
        // 更新段落的run的索引
        for (Map.Entry<XWPFRun, Integer> entry : runMap.entrySet()) {
            Integer value = entry.getValue();
            XWPFRun key = entry.getKey();
            if (value.compareTo(runMap.get(run)) > 0) {
                runMap.put(key, value + 1);
            }
        }
        runMap.put(newRun, runMap.get(run) + 1);
        return newRun;
    }

    /**
     * 在指定的run中, 添加docx批注的范围标签
     */
    public void addCommentLabel(Map<XWPFRun, List<BigInteger>> commentRangeStartMap, boolean start) {
        for (Map.Entry<XWPFRun, List<BigInteger>> entry : commentRangeStartMap.entrySet()) {
            for (BigInteger commentId : entry.getValue()) {
                XWPFRun run = entry.getKey();
                DocxHelper.insertCommentRangeToRun(run, start, commentId);
                if (!start) {
                    // 结束标签设置批注引用
                    CTMarkup ctMarkup = run.getCTR().addNewCommentReference();
                    ctMarkup.setId(commentId);
                }
            }
        }
    }
}
