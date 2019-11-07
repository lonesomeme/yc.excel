package com.lcfc.budget.excel;

import java.util.function.Function;

/**
 * 数据列，用于渲染列表数据
 *
 * @param <T> 元数据类型
 */
public class DataCol<T> extends ExcelCol {
    /**
     * 数据动态适配
     */
    Function value;
    T        data;

    public DataCol(String title) {
        super(title);
    }

    /**
     * 计算结果
     *
     * @param value
     * @param <TR>
     * @return
     */
    public <TR> DataCol value(Function<T, TR> value) {
        this.value = value;
        return this;
    }

    /**
     * 关联数据
     *
     * @param data
     * @return
     */
    public DataCol<T> data(T data) {
        this.data = data;
        return this;
    }
}
