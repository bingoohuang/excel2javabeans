package com.github.bingoohuang.beans2excel;

import com.github.bingoohuang.excel2beans.annotations.*;
import lombok.Builder;
import lombok.Data;
import lombok.Singular;

import java.util.List;

@Data @Builder
public class HanergyCepingResult {
    @ExcelCell(value = "A2", replace = "XX")
    private String interviewCode; // 面试编号

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
    private double matchScore;    // 岗位匹配度
    @ExcelCell("C6")
    private String matchComment;  // 岗位匹配度评语

    @ExcelRows(fromRef = "C7",
            mergeRows = {
                    @MergeRow(fromRef = "A5"),
                    @MergeRow(fromRef = "B7"),
                    @MergeRow(fromRef = "C7", type = MergeType.SameValue)},
            mergeCols = {
                    @MergeCol(fromColRef = "D", toColRef = "G")})
    @Singular private List<ItemComment> itemComments; // 四种工具测评结论

    @Data @Builder
    public static class ItemComment {
        private String item;
        private String comment;
    }

    @ExcelRows(fromColRef = "A", fromKey = "核心素质",
            mergeRows = {
                    @MergeRow(fromRef = "A"),
                    @MergeRow(fromRef = "B"),
                    @MergeRow(fromRef = "F", type = MergeType.SameValue, removePrefixBefore = "^")},
            mergeCols = {
                    @MergeCol(fromColRef = "B", toColRef = "C")})
    @Singular private List<Item> items;    // 附录

    @Data
    public static class Item {
        private String category;     // 类别
        private String blank;        // 留空，方便合并
        private String quality;      // 素质项
        private String dimension;    // 测评维度
        private String score;        // 得分
        private String remark;       // 备注
    }
}
