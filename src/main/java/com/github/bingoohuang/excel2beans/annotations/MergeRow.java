package com.github.bingoohuang.excel2beans.annotations;

import java.lang.annotation.*;

/**
 * 合并多行。
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface MergeRow {
    /**
     * 从哪个单元格开始。
     *
     * @return 开始单元格。
     */
    String fromRef();

    /**
     * 合并方式。
     *
     * @return 合并方式。
     */
    MergeType type() default MergeType.SameValue;


    /**
     * 修正单元格的值，去除指定字符的前缀。
     *
     * @return 前缀分隔符。
     */
    String prefixSeperate() default "";

    /**
     * 合并行的时候，同时合并的列数。
     *
     * @return 多合并的列数。
     */
    int moreCols() default 0;
}
