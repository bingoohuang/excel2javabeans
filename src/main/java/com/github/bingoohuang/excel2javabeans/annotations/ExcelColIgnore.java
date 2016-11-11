package com.github.bingoohuang.excel2javabeans.annotations;

import java.lang.annotation.*;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColIgnore {
}
