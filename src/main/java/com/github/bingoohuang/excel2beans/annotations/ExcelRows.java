package com.github.bingoohuang.excel2beans.annotations;

import java.lang.annotation.*;

/**
 * 指示写入Excel中的多个行。
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelRows {
    /**
     * 从哪个单元格开始写入，例如B7。当只指定了列索引时（例如B），需要结合searchKey使用。
     *
     * @return 起始单元格索引。
     */
    String fromRef() default "";

    /**
     * 单元格中内容所包含的关键字，与fromColRef配合使用。
     *
     * @return 单元格中内容所包含的关键字
     */
    String searchKey() default "";


    MergeRow[] mergeRows() default {};

    MergeCol[] mergeCols() default {};


}
