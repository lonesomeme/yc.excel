package com.lcfc.budget.excel;

import net.bytebuddy.implementation.bind.annotation.Super;

import java.lang.ref.WeakReference;

/**
 * excel单元格
 *
 * @author lin.jiale
 * @since 2019-8-10 22:54
 */
public class ExcelCol {
    /**
     * 父弱引用
     */
    WeakReference<ExcelCol> _parentRef;
    /**
     * 标题
     */
    String                  title;
    /**
     * 样式
     */
    ExcelColStyle           style;
    /**
     * 单元格宽度， 0位自动
     */
    int                     width;
    /**
     * 使用的行数量
     */
    int                     rows = 1;
    /**
     * 使用的列数量
     */
    int                     cols = 1;

    public ExcelCol(String title) {
        this.title = title;
    }

    public ExcelCol(String title, int rows, int cols) {
        this.title = title;
        this.rows = rows;
        this.cols = cols;
    }

    public ExcelCol style(ExcelColStyle style) {
        this.style = style;
        return this;
    }

    public ExcelCol rows(int rows) {
        this.rows = rows;
        return this;
    }

    public ExcelCol cols(int cols) {
        this.cols = cols;
        return this;
    }

    /**
     * 总使用的行数量
     *
     * @return
     */
    public int takeRows() {
        return rows;
    }

    /**
     * 总使用的列数量
     *
     * @return
     */
    public int takeCols() {
        return cols;
    }

    public ExcelCol width(int width) {
        this.width = width;
        return this;
    }
}
