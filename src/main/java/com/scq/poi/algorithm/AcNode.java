package com.scq.poi.algorithm;

import lombok.Getter;
import lombok.Setter;
import org.springframework.util.CollectionUtils;

import java.util.HashMap;
import java.util.Map;

/**
 * AC算法节点类
 *
 * @author by suchangqin
 * @date 2023/7/25 14:45
 */
@Getter
@Setter
public class AcNode {

    private Integer level;
    private Map<Character, AcNode> children;
    private AcNode failNode;
    private boolean end = false;

    /**
     * 重写, 因为子节点会有引用父节点的情况 , toString会循环调用, 所以不打印子节点的
     */
    @Override
    public String toString() {
        return "AcNode{" +
                "level=" + level +
                ", failNode=" + failNode +
                ", end=" + end +
                '}';
    }

    /**
     * 当前结点是否已包含指定Key值的子结点
     */
    public boolean containChildren(Character c) {
        return !CollectionUtils.isEmpty(children) && children.containsKey(c);
    }

    /**
     * 获取指定Key值的子结点
     */
    public AcNode getChildren(Character c) {
        return CollectionUtils.isEmpty(children) ? null : children.get(c);
    }

    /**
     * 添加子结点
     */
    public void addChildren(Character c, AcNode node) {
        if (children == null) {
            children = new HashMap<>(8);
        }
        children.put(c, node);
    }


}
