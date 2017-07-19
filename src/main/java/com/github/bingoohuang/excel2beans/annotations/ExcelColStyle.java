package com.github.bingoohuang.excel2beans.annotations;

import java.lang.annotation.*;

@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColStyle {
    /**
     * 对齐方式。
     *
     * @return 对齐方式
     */
    ExcelColAlign align() default ExcelColAlign.NONE;
}
