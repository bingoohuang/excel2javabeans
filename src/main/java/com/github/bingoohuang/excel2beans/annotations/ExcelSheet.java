package com.github.bingoohuang.excel2beans.annotations;

import java.lang.annotation.*;

/**
 * Created by bingoohuang on 2017/3/20.
 */
@Documented
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelSheet {
    String name();

    /**
     * 生成excel单页的抬头行（第一行，合并单元格）的键名。
     *
     * @return 抬头信息的主键名称
     */
    String headKey() default "";
}
