package com.github.bingoohuang.beans2excel;

import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.junit.Test;

import static com.github.bingoohuang.excel2beans.ExcelToBeansUtils.getClassPathWorkbook;

public class HanergyCepingTest {
    @Test @SneakyThrows
    public void test() {
        @Cleanup val wb = getClassPathWorkbook("hanergy-ceping.xlsx");
        val cr = new CellReference("C5");
        val row = wb.getSheetAt(0).getRow(cr.getRow());
        val cell = row.getCell(cr.getCol());
        cell.setCellValue(7.7);

//        val name = "ceping-result.xlsx";
//        ExcelToBeansUtils.writeExcel(wb, name);
//        new File(name).delete();
    }

    public static String getCellValue(Workbook workbook, String cellReference) {
        val cr = new CellReference(cellReference);
        val row = workbook.getSheetAt(0).getRow(cr.getRow());
        val cell = row.getCell(cr.getCol());
        return cell.getStringCellValue();
    }
}
