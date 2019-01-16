package com.github.bingoohuang.beans2excel;

import com.alibaba.fastjson.JSON;
import com.github.bingoohuang.excel2beans.BeansToExcelOnTemplate;
import com.github.bingoohuang.excel2beans.PoiUtil;
import com.github.bingoohuang.excel2beans.annotations.*;
import com.github.bingoohuang.utils.lang.Classpath;
import com.github.bingoohuang.utils.lang.Collects;
import lombok.*;
import org.apache.commons.lang3.StringUtils;
import org.junit.Test;

import java.util.List;

public class CepingTest {
    @Test
    @SneakyThrows
    public void test() {
        String json = Classpath.loadResAsString("cepingtable.json");
        ExportCepingResultTable table = JSON.parseObject(json, ExportCepingResultTable.class);

        @Cleanup val wb = PoiUtil.getClassPathWorkbook("cepingtmpl.xlsx");


        val templateName = Collects.isEmpty(table.getItemComments())
                ? StringUtils.isEmpty(table.getMatchComment()) ? "无总评语-模板" : "无评语-模板" : "有评语-模板";

        @Cleanup val newWb = new BeansToExcelOnTemplate(wb.getSheet(templateName)).create(table);

        PoiUtil.addImage(newWb.getSheetAt(0), "p.png", "G5");

        PoiUtil.writeExcel(newWb, templateName + "-" + table.getName() + ".xlsx");
    }


    // 结论表导出模型
    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    @Builder
    public static class ExportCepingResultTable {
        @ExcelCell
        private String title;           // 模板名称
        @ExcelCell(sheetName = true)
        private String sheetName;       // 表单名称

        @ExcelCell
        private String name;           // 身份证姓名
        @ExcelCell
        private String gender;         // 性别
        @ExcelCell
        private String age;            // 年龄

        @ExcelCell
        private String position;       // 应聘职位
        @ExcelCell
        private String level;          // 推荐职级
        @ExcelCell
        private String education;      // 学历

        @ExcelCell
        private double matchScore;     // 岗位匹配度
        @ExcelCell(maxLineLen = 40)
        private String matchComment;   // 岗位匹配度评语

        @ExcelRows(fromRef = "C", searchKey = "心理健康",   // 起点单元格在C列包含有"心理健康"关键字的单元格
                mergeRows = {
                        @MergeRow(fromRef = "A4", type = MergeType.Direct), // 从A5单元格开始向下直接合并
                        @MergeRow(fromRef = "B", type = MergeType.Direct),  // 从B列单元格开始向下直接合并
                        @MergeRow(fromRef = "C"), @MergeRow(fromRef = "D")},                          // 从C列单元格开始向下按值合并
                mergeCols = {
                        @MergeCol(fromColRef = "D", toColRef = "G")})
        @Singular
        private List<ItemComment> itemComments; // 四种工具测评结论

        @Data
        @NoArgsConstructor
        @AllArgsConstructor
        @Builder
        public static class ItemComment {
            private String item;
            @ExcelCell
            private String comment;
        }

        @ExcelRows(fromRef = "A", searchKey = "核心素质", // 起点在A列的包含有"核心素质"关键字的单元格
                mergeRows = {
                        @MergeRow(fromRef = "A"),                        // 从A列单元格开始向下按值合并
                        @MergeRow(fromRef = "B", moreCols = 1),          // 从B列单元格开始向下按值合并，并且多合并1列
                        @MergeRow(fromRef = "F", prefixSeperate = "^"),  // 从F列单元格开始向下按值合并，并且合并后去除^分割的前缀
                        @MergeRow(fromRef = "G", prefixSeperate = "^")}, // 从G列单元格开始向下按值合并
                mergeCols = {
                        @MergeCol(fromColRef = "D", toColRef = "E")})    // 每一行，合并单元格：从D行到E行
        @Singular
        private List<Item> items;    // 附录

        @Data
        @NoArgsConstructor
        @AllArgsConstructor
        @Builder
        public static class Item {
            private String category;     // 类别
            private String quality;      // 素质项
            private String _1;           // 留空，方便合并
            @ExcelCell(maxLineLen = 15)
            private String dimension;    // 测评维度
            private String _2;           // 留空，方便合并
            @ExcelCell(templateCells = {"FAIL:F12", "DEFAULT:F9"})
            private String score;        // 得分
            private String scoreTmpl;    // 得分单元格套用模板单元格名称，传PASS/FAIL
            private String remark;       // 备注
        }
    }
}


