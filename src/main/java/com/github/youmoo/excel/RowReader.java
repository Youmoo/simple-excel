package com.github.youmoo.excel;

import java.util.List;

/**
 * Excel表格行数据提取接口
 *
 * @author youmoo
 * @since 2013年12月1日
 */
public interface RowReader<T> {
    /**
     * 读取行数据
     *
     * @param t 行数据的提供者
     * @return 一行中所有单元格数据的集合
     */
    public List<Object> read(T t);
}
