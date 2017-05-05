package com.github.bingoohuang.excel2beans.annotations;

import java.lang.annotation.*;

/**
 * Created by bingoohuang on 2017/5/5.
 */
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
