package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import lombok.*;
import org.apache.commons.lang3.StringUtils;
import org.junit.Test;

import static com.github.bingoohuang.excel2beans.ExcelToBeansUtils.getClassPathWorkbook;
import static com.google.common.truth.Truth.assertThat;

@SuppressWarnings("unchecked")
public class BeanWithTitleTest {
    @SneakyThrows
    @Test public void test() {
        @Cleanup val workbook = getClassPathWorkbook("member.xlsx");
        val excelToBeans = new ExcelToBeans(workbook);
        val beans = excelToBeans.convert(BeanWithTitle.class);
        assertThat(beans).hasSize(4);

        assertThat(beans.get(0).getRowNum()).isEqualTo(6);
        assertThat(beans.get(1).getRowNum()).isEqualTo(7);
        assertThat(beans.get(2).getRowNum()).isEqualTo(8);
        assertThat(beans.get(3).getRowNum()).isEqualTo(9);

        assertThat(beans.get(0)).isEqualTo(BeanWithTitle.builder()
                .memberName("张小凡").sex("女").cardPrice("2880").cardName("示例次卡（100次次卡）").effective("2016-08-01").build());
        assertThat(beans.get(1)).isEqualTo(BeanWithTitle.builder()
                .memberName("李红").sex("男").cardPrice("3000").cardName("示例年卡（一周3次年卡）").effective("2016-01-01").build());
        assertThat(beans.get(2)).isEqualTo(BeanWithTitle.builder()
                .memberName("李红").sex("男").cardPrice("0").cardName("示例私教卡（60次私教卡）").effective("2015-05-01").build());
        assertThat(beans.get(3)).isEqualTo(BeanWithTitle.builder()
                .memberName("张晓").sex("女").cardPrice(null).cardName(null).effective(null).build());
    }

    @Data @Builder
    public static class BeanWithTitle extends ExcelRowRef implements ExcelRowIgnorable {
        @ExcelColTitle("会员姓名") String memberName;
        @ExcelColTitle("卡名称") String cardName;
        @ExcelColTitle("办卡价格") String cardPrice;
        @ExcelColTitle String sex;
        @ExcelColTitle(value = "地址", required = false) String addr;
        @ExcelColTitle("有效期开始日") String effective;

        @Override public boolean ignoreRow() {
            return StringUtils.startsWith(memberName, "示例-");
        }
    }
}
