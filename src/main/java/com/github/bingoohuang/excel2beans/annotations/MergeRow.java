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
     * @return 开始单元格
     */
    String fromRef();

    /**
     * 合并方式。
     *
     * @return
     */
    MergeType type() default MergeType.Direct;


    /**
     * 修正单元格的值，去除指定字符的前缀。
     * @return
     */
    String removePrefixBefore() default "";
}
