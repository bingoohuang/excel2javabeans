package com.github.bingoohuang.excel2beans.annotations;

import java.lang.annotation.*;

/**
 * Ignore the excel mapping for the annotated field.
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColIgnore {
}
