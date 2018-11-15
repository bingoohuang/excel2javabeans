package com.github.bingoohuang.excel2beans.annotations;

import java.lang.annotation.*;

/**
 * 关联Excel中的单元格。
 */
@Documented
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCell {
    /**
     * 单元格索引。例如A5,B12等。
     *
     * @return 单元格索引
     */
    String value() default "";


    /**
     * 替换单元格中取值的指定内容。例如把XX替换成真实取值。
     * 默认值（不设置时）是替换整个单元格取值。
     *
     * @return 替换内容
     */
    String replace() default "";

    /**
     * 是否是处理表单名称。
     *
     * @return 是否是处理表单名称
     */
    boolean sheetName() default false;


    /**
     * 单元格一行最大字符数。
     *
     * @return 最大字符数
     */
    int maxLineLen() default 0;


    /**
     * 模板单元格定义。
     * 注意：模板单元请在最上面的独立行中进行定义。EXCEL生成完毕后，会删除模板单元格所在行。
     * 例如 {"PASS:F10", "FAIL:F11"}
     * private String score;        // 得分
     * private String scoreTmpl;    // 得分套用模板名称，传PASS/FAIL
     *
     * @return 模板单元格定义。
     */
    String[] templateCells() default {};
}
