package com.github.bingoohuang.excel2beans;

import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.val;
import org.junit.Test;

import static com.google.common.truth.Truth.assertThat;

public class ExcelToBeansUtilsTest {
    @Test @SneakyThrows
    public void test() {
        @Cleanup val wb = PoiUtil.getClassPathWorkbook("af-comments.xlsx");
        ExcelToBeansUtils.removeAllComments(wb);

//        ExcelToBeansUtils.writeExcel(workbook, "af-without-comments.xlsx");

        int total = 0;
        for (int i = 0, ii = wb.getNumberOfSheets(); i < ii; ++i) {
            val sheet = wb.getSheetAt(i);
            val comments = sheet.getCellComments();
            total += comments.size();
        }

        assertThat(total).isEqualTo(0);
    }

    @Test @SneakyThrows
    public void getWorkbookBytes() {
        @Cleanup val wb = PoiUtil.getClassPathWorkbook("af-comments.xlsx");
        byte[] bytes = PoiUtil.getWorkbookBytes(wb);
        assertThat(bytes).isNotEmpty();
    }
}
