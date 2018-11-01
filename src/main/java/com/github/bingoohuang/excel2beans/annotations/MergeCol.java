package com.github.bingoohuang.excel2beans.annotations;

import java.lang.annotation.*;

/**
 * 在列上合并单元格。
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface MergeCol {
    /**
     * EXCEL中以字母标识的开始列索引，例如A,B。
     *
     * @return 开始列索引。
     */
    String fromColRef();

    /**
     * EXCEL中以字母标识的开始列索引，例如E,F。
     *
     * @return 结束列索引。
     */
    String toColRef();
}
