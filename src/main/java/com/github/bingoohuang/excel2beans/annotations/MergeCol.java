package com.github.bingoohuang.excel2beans.annotations;

import java.lang.annotation.*;

/**
 * 在列上合并单元格。
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface MergeCol {
    String fromColRef();

    String toColRef();
}
