package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColIgnore;
import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.google.common.collect.ImmutableMap;
import com.google.common.collect.Lists;
import lombok.*;
import org.apache.commons.lang3.StringUtils;
import org.junit.Test;

import java.util.List;
import java.util.Map;

import static com.google.common.truth.Truth.assertThat;

public class MultipleColumnsTest {
    @SneakyThrows
    @Test public void testAf() {
        @Cleanup val workbook = ExcelToBeansUtils.getClassPathWorkbook("af-tvplays.xlsx");
        val excelToBeans = new ExcelToBeans(workbook);
        val beans = excelToBeans.convert(AfTvPlayBean.class);
        assertThat(beans).hasSize(2);

        assertThat(beans.get(0)).isEqualTo(AfTvPlayBean.builder()
                .playName("大风车")
                .playDescs(Lists.newArrayList("蹲蹲蹲", "跳跳跳", "转转转"))
                .playUrls(Lists.newArrayList("aaa", null, "ccc"))
                .build());
        assertThat(beans.get(1)).isEqualTo(AfTvPlayBean.builder()
                .playName("小风车")
                .playDescs(Lists.newArrayList("蹲蹲蹲", "跳跳跳", "转转转"))
                .playUrls(Lists.newArrayList("aaa", "bbb", "ccc"))
                .build());
    }

    @Data @Builder
    public static class AfTvPlayBean {
        @ExcelColTitle("节目名称")
        private String playName;
        @ExcelColTitle("剧集描述")
        private List<String> playDescs;
        @ExcelColTitle("URL")
        private List<String> playUrls;
    }

    @SneakyThrows
    @Test public void testAfAttach() {
        @Cleanup val workbook = ExcelToBeansUtils.getClassPathWorkbook("af-tvplays.xlsx");
        val excelToBeans = new ExcelToBeans(workbook);
        val beans = excelToBeans.convert(AfTvPlayBeanAttach.class);
        assertThat(beans).hasSize(2);

        val bean0 = beans.get(0);
        assertThat(bean0).isEqualTo(AfTvPlayBeanAttach.builder()
                .playName("大风车")
                .playDescs(Lists.newArrayList("蹲蹲蹲", "跳跳跳", "转转转"))
                .playUrls(Lists.newArrayList("aaa", null, "ccc"))
                .cellDataMap(ImmutableMap.<String, CellData>builder()
                        .put("playName", CellData.builder().row(1).col(1).value("大风车").build())

                        .put("playDescs_0", CellData.builder().row(1).col(6).value("蹲蹲蹲").build())
                        .put("playUrls_0", CellData.builder().row(1).col(7).value("aaa").build())
                        .put("playDescs_1", CellData.builder().row(1).col(8).value("跳跳跳").build())
                        .put("playUrls_1", CellData.builder().row(1).col(9).value(null).build())
                        .put("playDescs_2", CellData.builder().row(1).col(10).value("转转转").build())
                        .put("playUrls_2", CellData.builder().row(1).col(11).value("ccc").build())
                        .build())
                .build());

        val bean1 = beans.get(1);
        assertThat(bean1).isEqualTo(AfTvPlayBeanAttach.builder()
                .playName("小风车")
                .playDescs(Lists.newArrayList("蹲蹲蹲", "跳跳跳", "转转转"))
                .playUrls(Lists.newArrayList("aaa", "bbb", "ccc"))
                .cellDataMap(ImmutableMap.<String, CellData>builder()
                        .put("playName", CellData.builder().row(2).col(1).value("小风车").build())

                        .put("playDescs_0", CellData.builder().row(2).col(6).value("蹲蹲蹲").build())
                        .put("playUrls_0", CellData.builder().row(2).col(7).value("aaa").build())
                        .put("playDescs_1", CellData.builder().row(2).col(8).value("跳跳跳").build())
                        .put("playUrls_1", CellData.builder().row(2).col(9).value("bbb").build())
                        .put("playDescs_2", CellData.builder().row(2).col(10).value("转转转").build())
                        .put("playUrls_2", CellData.builder().row(2).col(11).value("ccc").build())
                        .build())
                .build());

        val cellDataMap = bean0.getCellDataMap();
        val cellData = cellDataMap.get("playUrls_1");
        cellData.setComment("URL不能为空");
        cellData.setCommentAuthor("AF导入程序");
        ExcelToBeansUtils.writeRedComments(workbook, cellData);

        bean0.setError("error");
        excelToBeans.removeOkRows(AfTvPlayBeanAttach.class, beans);
        ExcelToBeansUtils.writeExcel(workbook, "af.xlsx");
    }

    @Data @Builder
    public static class AfTvPlayBeanAttach extends ExcelRowRef implements CellDataMapAttachable {
        @ExcelColTitle("节目名称")
        private String playName;
        @ExcelColTitle("剧集描述")
        private List<String> playDescs;
        @ExcelColTitle("URL")
        private List<String> playUrls;
        @ExcelColIgnore
        private Map<String, CellData> cellDataMap;

        @Override public void attachCellDataMap(Map<String, CellData> cellDataMap) {
            this.cellDataMap = cellDataMap;
        }
    }

    @SneakyThrows
    @Test public void test() {
        @Cleanup val workbook = ExcelToBeansUtils.getClassPathWorkbook("listColumns.xlsx");
        val excelToBeans = new ExcelToBeans(workbook);
        val beans = excelToBeans.convert(MultipleColumnsBeanWithTitle.class);
        assertThat(beans).hasSize(4);

        assertThat(beans.get(0).getRowNum()).isEqualTo(7);
        assertThat(beans.get(1).getRowNum()).isEqualTo(8);
        assertThat(beans.get(2).getRowNum()).isEqualTo(9);
        assertThat(beans.get(3).getRowNum()).isEqualTo(10);

        assertThat(beans.get(0)).isEqualTo(MultipleColumnsBeanWithTitle.builder().memberName("张小凡")
                .mobiles(Lists.newArrayList(null, "18795952311", "18795952311", "18795952311"))
                .homeareas(Lists.newArrayList("南京", "北京", "上海", "广东"))
                .build());
        assertThat(beans.get(1)).isEqualTo(MultipleColumnsBeanWithTitle.builder().memberName("李红")
                .mobiles(Lists.newArrayList("18676952432", null, "18676952432", "18676952432"))
                .homeareas(Lists.newArrayList("北京", "天津", "西安", "广西"))
                .build());
        assertThat(beans.get(2)).isEqualTo(MultipleColumnsBeanWithTitle.builder().memberName("李红")
                .mobiles(Lists.newArrayList("18676952432", "18676952432", null, "18676952432"))
                .homeareas(Lists.newArrayList("西安", "郑州", "福建", "湖南"))
                .build());
        assertThat(beans.get(3)).isEqualTo(MultipleColumnsBeanWithTitle.builder().memberName("张晓")
                .mobiles(Lists.newArrayList("13745367698", "13745367698", "13745367698", "13745367698"))
                .homeareas(Lists.newArrayList("杭州", "福州", "西宁", "湖北"))
                .build());
    }

    @Data @Builder
    public static class MultipleColumnsBeanWithTitle extends ExcelRowRef implements ExcelRowIgnorable {
        @ExcelColTitle("会员姓名") String memberName;
        @ExcelColTitle("手机号") List<String> mobiles;
        @ExcelColTitle("归属地") List<String> homeareas;

        @Override public boolean ignoreRow() {
            return StringUtils.startsWith(memberName, "示例-");
        }
    }
}
