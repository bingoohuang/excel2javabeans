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
}
