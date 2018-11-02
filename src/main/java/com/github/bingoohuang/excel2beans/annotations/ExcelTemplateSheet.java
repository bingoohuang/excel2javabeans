package com.github.bingoohuang.excel2beans.annotations;

import java.lang.annotation.*;

@Documented
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelTemplateSheet {
    int titleRowRef() default 1;

    int templateRowRef() default 2;
}
