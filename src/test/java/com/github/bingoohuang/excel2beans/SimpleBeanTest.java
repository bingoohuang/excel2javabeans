package com.github.bingoohuang.excel2beans;

import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.util.List;

import static com.github.bingoohuang.excel2beans.ExcelToBeansUtils.getClassPathWorkbook;
import static com.google.common.truth.Truth.assertThat;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
@SuppressWarnings("unchecked")
public class SimpleBeanTest {
    @Test @SneakyThrows public void test() {
        val workbook = getClassPathWorkbook("simplebeans.xlsx");

        val excelToBeans = new ExcelSheetToBeans(workbook, SimpleBean.class);
        List<SimpleBean> beans = excelToBeans.convert();

        assertThat(beans).hasSize(2);
        assertThat(beans.get(0)).isEqualTo(SimpleBean.builder().name("bingoo").addr("nanjing").build());
        assertThat(beans.get(1)).isEqualTo(SimpleBean.builder().name("huang").addr("beijing").build());
    }

    @Test @SneakyThrows public void testWithBlanklines() {
        val workbook = getClassPathWorkbook("simplebeans-blanklines.xlsx");

        val excelToBeans = new ExcelSheetToBeans(workbook, SimpleBean.class);
        List<SimpleBean> beans = excelToBeans.convert();

        assertThat(beans).hasSize(2);
        assertThat(beans.get(0)).isEqualTo(SimpleBean.builder().name("bingoo").addr("nanjing").build());
        assertThat(beans.get(1)).isEqualTo(SimpleBean.builder().name("huang").addr("beijing").build());
    }
}
