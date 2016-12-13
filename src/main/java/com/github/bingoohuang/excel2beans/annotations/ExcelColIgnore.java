package com.github.bingoohuang.excel2beans.annotations;

import java.lang.annotation.*;

/**
 * Ignore the excel mapping for the annotated field.
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColIgnore {
}
