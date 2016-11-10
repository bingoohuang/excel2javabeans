package com.github.bingoohuang.excel2javabeans;

import lombok.val;
import org.junit.Test;

import java.util.List;

import static com.github.bingoohuang.excel2javabeans.ExcelToBeansUtils.getClassPathWorkbook;
import static com.google.common.truth.Truth.assertThat;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
public class BeanWithTitleTest {
    @Test
    public void test() {
        val workbook = getClassPathWorkbook("member.xlsx");
        val excelToBeans = new ExcelToBeans(BeanWithTitle.class);
        List<BeanWithTitle> beans = excelToBeans.convert(workbook);
        assertThat(beans).hasSize(4);

        assertThat(beans.get(0).getRowNum()).isEqualTo(6);
        assertThat(beans.get(1).getRowNum()).isEqualTo(7);
        assertThat(beans.get(2).getRowNum()).isEqualTo(8);
        assertThat(beans.get(3).getRowNum()).isEqualTo(9);

        assertThat(beans.get(0)).isEqualTo(new BeanWithTitle(
                "张小凡", "女", "示例次卡（100次次卡）", "2880"));
        assertThat(beans.get(1)).isEqualTo(new BeanWithTitle(
                "李红", "男", "示例年卡（一周3次年卡）", "3000"));
        assertThat(beans.get(2)).isEqualTo(new BeanWithTitle(
                "李红", "男", "示例私教卡（60次私教卡）", "0"));
        assertThat(beans.get(3)).isEqualTo(new BeanWithTitle(
                "张晓", "女", null, null));
    }
}
