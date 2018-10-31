package com.github.bingoohuang.beans2excel;

import com.github.bingoohuang.excel2beans.ExcelToBeansUtils;
import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.junit.Test;

import static com.github.bingoohuang.excel2beans.ExcelToBeansUtils.getClassPathWorkbook;

public class HanergyCepingTest {
    @Test @SneakyThrows
    public void testInsertRows() {
        @Cleanup val wb = getClassPathWorkbook("test.xlsx");
        val sh = wb.getSheetAt(0);

//        sh.removeRow(sh.getRow(0));
        sh.shiftRows(1, sh.getLastRowNum(), 1);

//        copyRow(sh, 3);
        val name = "test-result.xlsx";
        ExcelToBeansUtils.writeExcel(wb, name);
    }

    void copyRow(Sheet worksheet, int rowNum) {
        Row sourceRow = worksheet.getRow(rowNum);

        //Save the text of any formula before they are altered by row shifting
        String[] formulasArray = new String[sourceRow.getLastCellNum()];
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            if (sourceRow.getCell(i) != null && sourceRow.getCell(i).getCellTypeEnum() == CellType.FORMULA)
                formulasArray[i] = sourceRow.getCell(i).getCellFormula();
        }

        worksheet.shiftRows(rowNum, worksheet.getLastRowNum(), 1);
        Row newRow = sourceRow;  //Now sourceRow is the empty line, so let's rename it
        sourceRow = worksheet.getRow(rowNum + 1);  //Now the source row is at rowNum+1

        // Loop through source columns to add to new row
        for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
            // Grab a copy of the old/new cell
            Cell oldCell = sourceRow.getCell(i);
            Cell newCell;

            // If the old cell is null jump to next cell
            if (oldCell == null) {
                continue;
            } else {
                newCell = newRow.createCell(i);
            }

            // Copy style from old cell and apply to new cell
            CellStyle newCellStyle = worksheet.getWorkbook().createCellStyle();
            newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
            newCell.setCellStyle(newCellStyle);

            // If there is a cell comment, copy
            if (oldCell.getCellComment() != null) {
                newCell.setCellComment(oldCell.getCellComment());
            }

            // If there is a cell hyperlink, copy
            if (oldCell.getHyperlink() != null) {
                newCell.setHyperlink(oldCell.getHyperlink());
            }

            // Set the cell data type
            newCell.setCellType(oldCell.getCellTypeEnum());

            // Set the cell data value
            switch (oldCell.getCellTypeEnum()) {
                case BLANK:
                    break;
                case BOOLEAN:
                    newCell.setCellValue(oldCell.getBooleanCellValue());
                    break;
                case ERROR:
                    newCell.setCellErrorValue(oldCell.getErrorCellValue());
                    break;
                case FORMULA:
                    newCell.setCellFormula(formulasArray[i]);
                    break;
                case NUMERIC:
                    newCell.setCellValue(oldCell.getNumericCellValue());
                    break;
                case STRING:
                    newCell.setCellValue(oldCell.getRichStringCellValue());
                    break;
                default:
                    break;
            }
        }

        // If there are any merged regions in the source row, copy to new row
        for (int i = 0; i < worksheet.getNumMergedRegions(); i++) {
            CellRangeAddress cellRangeAddress = worksheet.getMergedRegion(i);
            if (cellRangeAddress.getFirstRow() == sourceRow.getRowNum()) {
                CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.getRowNum(),
                        (newRow.getRowNum() +
                                (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow()
                                )),
                        cellRangeAddress.getFirstColumn(),
                        cellRangeAddress.getLastColumn());
                worksheet.addMergedRegion(newCellRangeAddress);
            }
        }
    }

    @Test @SneakyThrows
    public void test() {
        @Cleanup val wb = getClassPathWorkbook("hanergy-ceping.xlsx");
        val cr = new CellReference("C5");
        val sh = wb.getSheetAt(0);
        val row = sh.getRow(cr.getRow());
        val cell = row.getCell(cr.getCol());
        cell.setCellValue(7.7);

        sh.shiftRows(6, sh.getLastRowNum(), 1);

        val name = "ceping-result.xlsx";
        ExcelToBeansUtils.writeExcel(wb, name);
//        new File(name).delete();
    }

    public static String getCellValue(Workbook workbook, String cellReference) {
        val cr = new CellReference(cellReference);
        val row = workbook.getSheetAt(0).getRow(cr.getRow());
        val cell = row.getCell(cr.getCol());
        return cell.getStringCellValue();
    }
}
