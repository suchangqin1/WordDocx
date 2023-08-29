package com.mocheng.poi.algorithm;

import lombok.Getter;
import org.springframework.util.CollectionUtils;

import java.util.*;

/**
 * AC自动机算法树
 * 参考文档:
 *  <a href="https://blog.csdn.net/Mr_SCX/article/details/104065446">...</a>
 * <a href="https://www.iteye.com/blog/buddie-2391496">...</a>
 *
 * @author by suchangqin
 * @date 2023/7/25 14:11
 */
@Getter
public class AcTree {

    private final AcNode rootNode;


    public AcTree(List<String> wordList) {
        rootNode = new AcNode();
        initTree(wordList);
        buildFailNode();
    }

    private void initTree(List<String> wordList) {
        for (String word : wordList) {
            if (word.isEmpty()) {
                continue;
            }
            char[] charArray = word.toLowerCase().toCharArray();
            buildTreeByWord(charArray);
        }
    }

    private void buildTreeByWord(char[] charArray) {
        AcNode curNode = rootNode;
        List<Integer> wordsLengthList = new ArrayList<>(4);
        wordsLengthList.add(charArray.length);
        for (char c : charArray) {
            // 已存在则只移动指针
            if (curNode.containChildren(c)) {
                curNode = curNode.getChildren(c);
                continue;
            }
            // 没有则新增
            AcNode childrenNode = new AcNode();
            curNode.addChildren(c, childrenNode);
            curNode = childrenNode;
        }
        // 最后的字符节点如果没有设置过层级
        if (Objects.isNull(curNode.getLevel())) {
            curNode.setLevel(charArray.length);
            curNode.setEnd(true);
        }
    }

    private void buildFailNode() {
        buildFirstFailNode();
        buildOtherFailNode();

    }

    /**
     * 建立第一层级的节点的失败指针节点
     */
    private void buildFirstFailNode() {
        rootNode.getChildren().forEach((character, acNode) -> acNode.setFailNode(rootNode));
    }

    /**
     * 建立非第一层级的节点的失败指针节点
     */
    private void buildOtherFailNode() {
        Queue<AcNode> queue = new LinkedList<>(rootNode.getChildren().values());
        AcNode node;
        while (!queue.isEmpty()) {
            node = queue.remove();
            buildNodeFailLink(node, queue);
        }

    }

    /**
     * 根据父节点, 建立子节点的失败指针节点
     *
     * @param parent 父节点
     * @param queue  建立失败指针节点任务队列
     */
    private void buildNodeFailLink(AcNode parent, Queue<AcNode> queue) {
        if (CollectionUtils.isEmpty(parent.getChildren())) {
            return;
        }
        System.out.println("1111");
        queue.addAll(parent.getChildren().values());
        AcNode parentFailNode = parent.getFailNode();
        Set<Character> childrenKeys = parent.getChildren().keySet();
        AcNode failNode;
        for (Character key : childrenKeys) {
            failNode = parentFailNode;
            while (failNode != rootNode && !failNode.containChildren(key)) {
                failNode = failNode.getFailNode();
            }
            failNode = failNode.getChildren(key);
            if (failNode == null) {
                failNode = rootNode;
            }
            parent.getChildren(key).setFailNode(failNode);
        }
    }

}
