package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.google.common.collect.Lists;
import lombok.*;
import org.apache.commons.lang3.StringUtils;
import org.junit.Test;

import java.util.List;

import static com.google.common.truth.Truth.assertThat;

public class MultipleColumnsTest {
    @SneakyThrows
    @Test public void test() {
        @Cleanup val workbook = ExcelToBeansUtils.getClassPathWorkbook("listColumns.xlsx");
        val excelToBeans = new ExcelSheetToBeans(workbook, MultipleColumnsBeanWithTitle.class);
        List<MultipleColumnsBeanWithTitle> beans = excelToBeans.convert();
        assertThat(beans).hasSize(4);

        assertThat(beans.get(0).getRowNum()).isEqualTo(6);
        assertThat(beans.get(1).getRowNum()).isEqualTo(7);
        assertThat(beans.get(2).getRowNum()).isEqualTo(8);
        assertThat(beans.get(3).getRowNum()).isEqualTo(9);

        assertThat(beans.get(0)).isEqualTo(MultipleColumnsBeanWithTitle.builder().memberName("张小凡")
                .mobiles(Lists.newArrayList(null, "18795952311", "18795952311", "18795952311"))
                .build());
        assertThat(beans.get(1)).isEqualTo(MultipleColumnsBeanWithTitle.builder().memberName("李红")
                .mobiles(Lists.newArrayList("18676952432", null, "18676952432", "18676952432"))
                .build());
        assertThat(beans.get(2)).isEqualTo(MultipleColumnsBeanWithTitle.builder().memberName("李红")
                .mobiles(Lists.newArrayList("18676952432", "18676952432", null, "18676952432"))
                .build());
        assertThat(beans.get(3)).isEqualTo(MultipleColumnsBeanWithTitle.builder().memberName("张晓")
                .mobiles(Lists.newArrayList("13745367698", "13745367698", "13745367698", "13745367698"))
                .build());
    }

    @Data @Builder
    public static class MultipleColumnsBeanWithTitle extends ExcelRowRef implements ExcelRowIgnorable {
        @ExcelColTitle("会员姓名") String memberName;
        @ExcelColTitle("手机号") List<String> mobiles;

        @Override public boolean ignoreRow() {
            return StringUtils.startsWith(memberName, "示例-");
        }
    }
}
