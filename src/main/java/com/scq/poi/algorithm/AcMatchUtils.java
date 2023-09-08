package com.scq.poi.algorithm;


import lombok.AllArgsConstructor;
import org.springframework.util.StringUtils;

import java.util.*;

/**
 * AC多模式串匹配敏感词算法执行器
 *
 * @author by suchangqin
 * @date 2023/07/26
 */
@AllArgsConstructor
public class AcMatchUtils {
    /**
     * 敏感词集构建的树
     */
    private final AcTree tree;

    /**
     * 使用AC自动机算法, 将给定的文本过滤掉敏感词, 敏感词将被 "*" 替换
     * 参考实现: <a href="http://www.javashuo.com/article/p-fwqgabqb-sw.html">...</a>
     *
     * @param word 文本
     * @return 过滤后的文本
     */
    public String filter(String word) {
        if (StringUtils.isEmpty(word)) {
            return "";
        }
        //文本集
        char[] words = word.toLowerCase().toCharArray();
        //结果集
        char[] result = null;
        //文本匹配的敏感词集
        AcNode curNode = tree.getRootNode();
        AcNode childNode;
        Character c;
        int fromPos = 0;
        for (int i = 0; i < words.length; i++) {
            c = words[i];
            childNode = curNode.getChildren(c);

            while (childNode == null && curNode != tree.getRootNode()) {
                curNode = curNode.getFailNode();
                childNode = curNode.getChildren(c);
            }
            if (childNode != null) {
                curNode = childNode;
            }
            if (curNode.isEnd()) {
                int pos = i - curNode.getLevel() + 1;
                if (pos < fromPos) {
                    pos = fromPos;
                }
                if (result == null) {
                    result = word.toLowerCase().toCharArray();
                }

                for (; pos <= i; pos++) {
                    result[pos] = '*';
                }
                fromPos = i + 1;
            }
        }
        if (result == null) {
            return word;
        }
        return String.valueOf(result);
    }


    /**
     * 使用AC自动机算法, 匹配敏感词, 并获取命中的敏感词集,
     * 参考实现: <a href="https://blog.csdn.net/Mr_SCX/article/details/104065446">...</a>
     */
    public Map<String, List<Integer>> match(String word) {
        Map<String, List<Integer>> matchWordMap = new HashMap<>(word.length());
        char[] text = word.toCharArray();
        int textLength = text.length;
        AcNode rootNode = tree.getRootNode();
        AcNode p = rootNode;
        String matchWord;
        for (int i = 0; i < textLength; ++i) {
            char c = text[i];
            // 判断子节点中是否存在当前字符, 有则继续, 没有则触发失败指针
            while (p.getChildren(c) == null && p != rootNode) {
                // 失败指针发挥作用的地方
                p = p.getFailNode();
            }
            // 获取当前字符节点
            p = p.getChildren(c);
            // 如果没有匹配的，从root开始重新匹配
            if (p == null) {
                p = rootNode;
            }
            AcNode tmp = p;
            // 处理命中的敏感词
            while (tmp != rootNode) {
                if (tmp.isEnd()) {
                    int pos = i - tmp.getLevel() + 1;
                    // 截取命中的敏感词
                    if (i == textLength - 1) {
                        matchWord = word.substring(pos);
                    } else {
                        matchWord = word.substring(pos, pos + tmp.getLevel());
                    }

                    // 命中的敏感词
                    if (matchWordMap.containsKey(matchWord)) {
                        matchWordMap.get(matchWord).add(pos);
                    } else {
                        List<Integer> list = new ArrayList<>();
                        list.add(pos);
                        matchWordMap.put(matchWord, list);
                    }

                }
                tmp = tmp.getFailNode();
            }
        }

        return matchWordMap;
    }
}