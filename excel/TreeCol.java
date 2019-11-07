package com.lcfc.budget.excel;

import com.google.common.collect.Lists;
import lombok.Data;

import java.lang.ref.WeakReference;
import java.util.List;

/**
 * 合并的列-树列
 *
 * @author lin.jiale
 * @since 2019-11-1 22:29
 */
@Data
public class TreeCol extends ExcelCol {
    private List<ExcelCol> subCols;

    public TreeCol(String title) {
        super(title);
        subCols = Lists.newArrayList();
    }

    /**
     * 添加子节点
     *
     * @param col
     * @return
     */
    public TreeCol child(ExcelCol col) {
        col._parentRef = new WeakReference<>(this);
        subCols.add(col);
        return this;
    }

    /**
     * 总使用列数，遍历所有子节点计算得出
     *
     * @return
     */
    @Override
    public int takeCols() {
        int size = 1;
        if (subCols.size() > 0) {
            for (ExcelCol subCol : subCols) {
                size += subCol.takeCols();
            }
            size--;
        }
        return size;
    }

    /**
     * 使用的总行数，遍历所有子节点计算得出
     *
     * @return
     */
    @Override
    public int takeRows() {
        int trows = this.rows;
        if (subCols.size() > 0) {
            int max_t_rows = 0;
            for (ExcelCol subCol : subCols) {
                max_t_rows = Math.max(max_t_rows, subCol.takeRows());
            }
            trows += max_t_rows;
        }
        return trows;
    }

    /**
     * 树列使用的列数量由计算得出 ，不能直接赋值
     *
     * @param cols
     * @return
     */
    @Override
    public ExcelCol cols(int cols) {
        return this;
    }
}
