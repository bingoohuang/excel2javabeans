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
     * 从哪个单元格开始写入
     *
     * @return 起始单元格索引。
     */
    String fromRef() default "";


    /**
     * 从哪个列的哪个关键字所在的单元格开始，与fromKey配合使用。
     *
     * @return 列索引，例如A,B
     */
    String fromColRef() default "";

    /**
     * 单元格中内容所包含的关键字，与fromColRef配合使用。
     *
     * @return 单元格中内容所包含的关键字
     */
    String fromKey() default "";


    MergeRow[] mergeRows() default {};

    MergeCol[] mergeCols() default {};


}
