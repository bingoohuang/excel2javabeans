package com.github.bingoohuang.beans2excel;

import com.github.bingoohuang.excel2beans.annotations.*;
import lombok.Builder;
import lombok.Data;
import lombok.Singular;

import java.util.List;

@Data @Builder
public class CepingResult {
    @ExcelCell(sheetName = true)
    private String sheetName;       // 表单名称

    @ExcelCell(value = "A2", replace = "XX")
    private String interviewCode;   // 面试编号

    @ExcelCell("B3")
    private String name;           // 身份证姓名
    @ExcelCell("E3")
    private String gender;         // 性别
    @ExcelCell("G3")
    private String age;            // 年龄

    @ExcelCell("B4")
    private String position;       // 应聘职位
    @ExcelCell("E4")
    private String level;          // 推荐职级
    @ExcelCell("G4")
    private String annualSalary;   // 期望年薪

    @ExcelCell("C5")
    private double matchScore;     // 岗位匹配度
    @ExcelCell(value = "C6", maxLineLen = 40)
    private String matchComment;   // 岗位匹配度评语

    @ExcelRows(fromRef = "C", searchKey = "心理健康",     // 起点在A列的包含有"四种工具"关键字的单元格
            mergeRows = {
                    @MergeRow(fromRef = "A5", type = MergeType.Direct), // 从A5单元格开始向下直接合并
                    @MergeRow(fromRef = "B", type = MergeType.Direct),  // 从B列单元格开始向下直接合并
                    @MergeRow(fromRef = "C")},                          // 从C列单元格开始向下按值合并
            mergeCols = {
                    @MergeCol(fromColRef = "D", toColRef = "G")})
    @Singular private List<ItemComment> itemComments; // 四种工具测评结论

    @Data @Builder
    public static class ItemComment {
        private String item;
        @ExcelCell(maxLineLen = 36)
        private String comment;
    }

    @ExcelRows(fromRef = "A", searchKey = "核心素质", // 起点在A列的包含有"核心素质"关键字的单元格
            mergeRows = {
                    @MergeRow(fromRef = "A"),                        // 从A列单元格开始向下按值合并
                    @MergeRow(fromRef = "B", moreCols = 1),          // 从B列单元格开始向下按值合并，并且多合并1列
                    @MergeRow(fromRef = "F", prefixSeperate = "^"),  // 从F列单元格开始向下按值合并，并且合并后去除^分割的前缀
                    @MergeRow(fromRef = "G")},                       // 从G列单元格开始向下按值合并
            mergeCols = {
                    @MergeCol(fromColRef = "D", toColRef = "E")})    // 每一行，合并单元格：从D行到E行
    @Singular private List<Item> items;    // 附录

    @Data @Builder
    public static class Item {
        private String category;     // 类别
        private String quality;      // 素质项
        private String _1;           // 留空，方便合并
        @ExcelCell(maxLineLen = 15)
        private String dimension;    // 测评维度
        private String _2;           // 留空，方便合并
        @ExcelCell(templateCells = {"FAIL:F12", "DEFAULT:F9"})
        private String score;        // 得分
        private String scoreTmpl;    // 得分套用模板名称，传PASS/FAIL

        private String remark;       // 备注
    }
}
