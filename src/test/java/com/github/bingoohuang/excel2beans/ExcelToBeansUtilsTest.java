package com.github.bingoohuang.excel2beans;

import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.val;
import org.junit.Test;

import static com.google.common.truth.Truth.assertThat;

public class ExcelToBeansUtilsTest {
    @Test @SneakyThrows
    public void test() {
        @Cleanup val workbook = ExcelToBeansUtils.getClassPathWorkbook("af-comments.xlsx");
        ExcelToBeansUtils.removeAllComments(workbook);

//        ExcelToBeansUtils.writeExcel(workbook, "af-without-comments.xlsx");

        int total = 0;
        for (int i = 0, ii = workbook.getNumberOfSheets(); i < ii; ++i) {
            val sheet = workbook.getSheetAt(i);
            val comments = sheet.getCellComments();
            total += comments.size();
        }

        assertThat(total).isEqualTo(0);
    }
}
