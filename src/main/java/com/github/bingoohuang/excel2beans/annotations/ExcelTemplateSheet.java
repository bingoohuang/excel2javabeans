package com.github.bingoohuang.excel2beans.annotations;

import java.lang.annotation.*;

@Documented
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelTemplateSheet {
    /**
     * 标题所在行的索引（1-based）。默认值1。
     *
     * @return 标题行索引
     */
    int titleRowRef() default 1;

    /**
     * 值模板行索引（1-based)。默认值2。
     *
     * @return 值模板行索引
     */
    int templateRowRef() default 2;

    /**
     * 模板所在的Sheet名称。当模板excel中包含多个页时，需要指定模板页名称，如果不指定，则默认为第一页。
     *
     * @return 模板sheet名称。
     */
    String templateSheetName() default "";
}
