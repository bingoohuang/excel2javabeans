package com.github.bingoohuang.excel2beans;

import lombok.val;
import org.junit.Test;

import java.util.List;

import static com.github.bingoohuang.excel2beans.ExcelToBeansUtils.getClassPathWorkbook;
import static com.google.common.truth.Truth.assertThat;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
@SuppressWarnings("unchecked")
public class BeanWithTitleTest {
    @Test public void test() {
        val workbook = getClassPathWorkbook("member.xlsx");
        val excelToBeans = new ExcelToBeans(BeanWithTitle.class);
        List<BeanWithTitle> beans = excelToBeans.convert(workbook);
        assertThat(beans).hasSize(4);

        assertThat(beans.get(0).getRowNum()).isEqualTo(6);
        assertThat(beans.get(1).getRowNum()).isEqualTo(7);
        assertThat(beans.get(2).getRowNum()).isEqualTo(8);
        assertThat(beans.get(3).getRowNum()).isEqualTo(9);

        assertThat(beans.get(0)).isEqualTo(BeanWithTitle.builder().memberName("张小凡").sex("女").cardPrice("2880").cardName("示例次卡（100次次卡）").build());
        assertThat(beans.get(1)).isEqualTo(BeanWithTitle.builder().memberName("李红").sex("男").cardPrice("3000").cardName("示例年卡（一周3次年卡）").build());
        assertThat(beans.get(2)).isEqualTo(BeanWithTitle.builder().memberName("李红").sex("男").cardPrice("0").cardName("示例私教卡（60次私教卡）").build());
        assertThat(beans.get(3)).isEqualTo(BeanWithTitle.builder().memberName("张晓").sex("女").cardPrice(null).cardName(null).build());
    }
}
