package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelCell;
import com.github.bingoohuang.excel2beans.annotations.ExcelRows;
import com.google.common.collect.Lists;
import lombok.RequiredArgsConstructor;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.List;

@RequiredArgsConstructor
public class BeansToExcelOnTemplate {
    private final Sheet templateSheet;

    public Workbook create(Object bean) {
        for (val field : bean.getClass().getDeclaredFields()) {
            processExcelCell(field, bean);
            processExcelRows(field, bean);
        }

        removeOtherSheets();
        return templateSheet.getWorkbook();
    }

    @SneakyThrows
    private void processExcelRows(Field field, Object bean) {
        val excelRows = field.getAnnotation(ExcelRows.class);
        if (excelRows == null) return;


        final Type genericType = field.getGenericType();
        if (!(genericType instanceof ParameterizedType)) return;

        val pt = (ParameterizedType) genericType;

        if (pt.getRawType() != List.class) return;
        Object fieldValue = invokeField(field, bean);
        List list = (List) fieldValue;
        if (list == null || list.isEmpty()) return;

        val itemSize = list.size();
        val cellRef = new CellReference(excelRows.fromRef());
        final int fromRow = cellRef.getRow();
        if (itemSize > 1) {
            templateSheet.shiftRows(fromRow + 1, templateSheet.getLastRowNum(), itemSize - 1);
        }

        int rowOffset = 0;

        val templateRow = templateSheet.getRow(fromRow);
        for (val item : list) {
            val itemClass = item.getClass();

            val row = rowOffset == 0 ? templateRow
                    : templateSheet.createRow(fromRow + rowOffset);

            int colOffset = 0;
            for (val itemField : itemClass.getDeclaredFields()) {
                val itemFieldValue = invokeField(itemField, item);

                Cell cell = null;
                if (rowOffset == 0) {
                    cell = row.getCell(cellRef.getCol() + colOffset);
                }

                if (cell == null) {
                    cell = row.createCell(cellRef.getCol() + colOffset);
                    cell.setCellStyle(templateRow.getCell(cellRef.getCol() + colOffset).getCellStyle());
                }

                cell.setCellValue("" + itemFieldValue);
                ++colOffset;
            }

            ++rowOffset;
        }

    }

    @SneakyThrows
    private void processExcelCell(Field field, Object bean) {
        val excelCell = field.getAnnotation(ExcelCell.class);
        if (excelCell == null) return;

        val cell = findCell(excelCell.value());

        Object fieldValue = invokeField(field, bean);
        val fv = fieldValue == null ? "" : fieldValue;

        if (StringUtils.isNotEmpty(excelCell.replace())) {
            val stringCellValue = cell.getStringCellValue();
            val newValue = stringCellValue.replace(excelCell.replace(), "" + fv);
            cell.setCellValue(newValue);
        } else {
            if (fv instanceof Number) {
                cell.setCellValue(((Number) fv).doubleValue());
            } else {
                cell.setCellValue("" + fv);
            }
        }
    }

    private Object invokeField(Field field, Object bean) throws IllegalAccessException {
        if (!field.isAccessible()) field.setAccessible(true);
        return field.get(bean);
    }

    private Cell findCell(String cellRefValue) {
        val cellRef = new CellReference(cellRefValue);
        val row = templateSheet.getRow(cellRef.getRow());
        return row.getCell(cellRef.getCol());
    }

    private void removeOtherSheets() {
        val workbook = templateSheet.getWorkbook();

        List<String> deleteSheetNames = Lists.newArrayList();
        for (int i = 0; i < workbook.getNumberOfSheets(); ++i) {
            val sheetName = workbook.getSheetName(i);
            if (!sheetName.equals(templateSheet.getSheetName())) {
                deleteSheetNames.add(sheetName);
            }
        }

        for (val name : deleteSheetNames) {
            workbook.removeSheetAt(workbook.getSheetIndex(name));
        }
    }
}
