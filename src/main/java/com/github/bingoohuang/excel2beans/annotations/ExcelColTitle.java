package com.github.bingoohuang.excel2beans.annotations;

import java.lang.annotation.*;

/**
 * Set title substring for excel column value mapping.
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColTitle {
    String value();
}
