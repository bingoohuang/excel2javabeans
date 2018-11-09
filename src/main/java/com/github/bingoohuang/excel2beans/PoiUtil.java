package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelSheet;
import com.google.common.collect.Lists;
import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFDrawing;

import java.io.FileOutputStream;
import java.util.List;

public class PoiUtil {
    public static void blankCell(Row row, int col) {
        val cell = row.getCell(col);
        if (cell == null) return;

        if (cell.getCellTypeEnum() != CellType.BLANK) {
            cell.setCellType(CellType.BLANK);
        }
    }

    public static String getCellValue(Row row, int col) {
        val cell = row.getCell(col);
        if (cell == null) return null;

        val cellValue = cell.getStringCellValue();
        return StringUtils.trim(cellValue);
    }

    /**
     * 删除其它Sheet页，只保留指定模板的Sheet页输出。
     *
     * @param sheet 在工作簿需要保留的表单页。
     */
    public static void removeOtherSheets(Sheet sheet) {
        val wb = sheet.getWorkbook();

        List<String> deletedNames = Lists.newArrayList();
        for (int i = 0; i < wb.getNumberOfSheets(); ++i) {
            val sheetName = wb.getSheetName(i);
            if (!sheetName.equals(sheet.getSheetName())) {
                deletedNames.add(sheetName);
            }
        }

        deletedNames.forEach(x -> wb.removeSheetAt(wb.getSheetIndex(x)));
    }

    /**
     * 是否是完整的单元格索引。例如A5是完整的，A不是完整的，缺少行标。
     *
     * @param cellReference EXCEL中的单元格索引。
     * @return 完整的单元格索引时返回true。
     */
    public static boolean isFullCellReference(String cellReference) {
        return cellReference.matches("\\w+\\d+");
    }

    /**
     * 插入新的行（向下移动表格中的行），方便写入数据。
     *
     * @param sheet          表单页。
     * @param list           写入JavaBean列表。
     * @param templateRowNum 模板行号。
     * @return JavaBean列表大小。
     */
    public static int shiftRows(Sheet sheet, List<?> list, int templateRowNum) {
        val itemSize = list == null ? 0 : list.size();
        int lastRowNum = sheet.getLastRowNum();
        if (itemSize == 0 && templateRowNum == lastRowNum) { // 没有写入行，直接删除模板行
            sheet.removeRow(sheet.getRow(templateRowNum));
        } else if (itemSize != 1 && templateRowNum < lastRowNum) {
            sheet.shiftRows(templateRowNum + 1, lastRowNum, itemSize - 1);
        }

        return itemSize;
    }

    /**
     * 向单元格写入值，处理值为整型时的写入情况。
     *
     * @param cell           单元格。
     * @param fv             单元格值。
     * @param fixNumberValue 是否修正整型值写入。
     * @return 单元格字符串取值。
     */
    public static String writeCellValue(Cell cell, Object fv, boolean fixNumberValue) {
        if (fv instanceof Number) {
            val value = ((Number) fv).doubleValue();
            cell.setCellValue(value);
            return "" + value;
        }

        if (fv instanceof String && fixNumberValue) {
            val s = (String) fv;
            if (ExcelToBeansUtils.isNumeric(s)) {
                val value = Double.parseDouble(s);
                cell.setCellValue(value);
                return "" + value;
            }

            cell.setCellValue(s);
            return s;
        }

        final String value = "" + fv;
        cell.setCellValue(value);
        return value;
    }

    public static void removeRow(Sheet sheet, int rowIndex) {
        val lastRowNum = sheet.getLastRowNum();
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
        } else if (rowIndex == lastRowNum) {
            val removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }

    @SneakyThrows
    public static void writeExcel(Workbook workbook, String name) {
        @Cleanup val fileOut = new FileOutputStream(name);
        workbook.write(fileOut);
    }

    public static Sheet findSheet(Workbook workbook, Class<?> beanClass) {
        val excelSheet = beanClass.getAnnotation(ExcelSheet.class);
        if (excelSheet == null) return workbook.getSheetAt(0);

        for (int i = 0, ii = workbook.getNumberOfSheets(); i < ii; ++i) {
            val sheetName = workbook.getSheetName(i);
            if (sheetName.contains(excelSheet.name())) {
                return workbook.getSheetAt(i);
            }
        }

        throw new IllegalArgumentException("Unable to find sheet with name " + excelSheet.name());
    }

    /**
     * 根据单元格索引，找到单元格。
     *
     * @param sheet        EXCEL表单。
     * @param cellRefValue 单元格索引，例如A1, AB12等。
     * @return 单元格。
     */
    public static Cell findCell(Sheet sheet, String cellRefValue) {
        val cellRef = new CellReference(cellRefValue);
        val row = sheet.getRow(cellRef.getRow());
        return row.getCell(cellRef.getCol());
    }

    /**
     * 修正图表中对于表单名字的引用。
     *
     * @param sheet        EXCEL表单。
     * @param oldSheetName 旧的表单名字。
     * @param newSheetName 新的表单名字。
     */
    public static void fixChartSheetNameRef(Sheet sheet, String oldSheetName, String newSheetName) {
        val drawing = sheet.getDrawingPatriarch();
        if (!(drawing instanceof XSSFDrawing)) return;

        for (val chart : ((XSSFDrawing) drawing).getCharts()) {
            for (val barChart : chart.getCTChart().getPlotArea().getBarChartList()) {
                for (val ser : barChart.getSerList()) {
                    val val = ser.getVal();
                    if (val == null) continue;

                    val numRef = val.getNumRef();
                    if (numRef == null) continue;

                    val f = numRef.getF();
                    if (f == null) continue;

                    if (f.contains(oldSheetName)) {
                        numRef.setF(f.replace(oldSheetName, newSheetName));
                    }
                }
            }
        }
    }
}
