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
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
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

        val genericType = field.getGenericType();
        if (!(genericType instanceof ParameterizedType)) return;

        val pt = (ParameterizedType) genericType;
        if (pt.getRawType() != List.class) return;

        val templateCell = findTemplateCell(excelRows);
        val list = (List) invokeField(field, bean);

        val itemSize = shiftRows(templateCell, list);
        if (itemSize > 0) {
            writeRows(templateCell, list);
            mergeRows(excelRows, templateCell, itemSize);
        }
    }

    private void mergeRows(ExcelRows excelRows, Cell templateCell, int itemSize) {
        for (val mergeRow : excelRows.mergeRows()) {
            val fromCell = findCell(mergeRow.fromRef());
            val fromRow = fromCell.getRow().getRowNum();
            val lastRow = templateCell.getRowIndex() + itemSize - 1;
            val col = fromCell.getColumnIndex();
            switch (mergeRow.type()) {
                case Direct:
                    templateSheet.addMergedRegion(new CellRangeAddress(fromRow, lastRow, col, col));
                    break;
                case SameValue:
//                    final String lastValue = fromCell.getStringCellValue();
//                    int preRow = fromRow;
//                    for (int i = 1; i < itemSize; ++i) {
//                        val c = templateSheet.getRow(fromRow +i).getCell(col);
//                        if (!c.getStringCellValue().equals(lastValue)) {
//                            templateSheet.addMergedRegion(new CellRangeAddress(preRow, fromRow + i - 1, col, col));
//                        }
//
//                    }

                    break;
            }
        }
    }

    private int shiftRows(Cell templateCell, List list) {
        val itemSize = list == null ? 0 : list.size();
        int fromRow = templateCell.getRow().getRowNum();
        if (fromRow == templateSheet.getLastRowNum()) {
            templateSheet.removeRow(templateSheet.getRow(fromRow));
        } else {
            templateSheet.shiftRows(fromRow + 1, templateSheet.getLastRowNum(), itemSize - 1);
        }

        return itemSize;
    }

    private void writeRows(Cell templateCell, List list) {
        val templateRow = templateCell.getRow();
        int fromRow = templateRow.getRowNum();

        val templateCol = templateCell.getColumnIndex();
        int rowOffset = 0;
        for (val item : list) {
            val itemClass = item.getClass();

            val row = rowOffset == 0 ? templateRow
                    : templateSheet.createRow(fromRow + rowOffset);

            int colOffset = 0;
            for (val itemField : itemClass.getDeclaredFields()) {
                val itemFieldValue = invokeField(itemField, item);

                newCell(templateRow, templateCol, rowOffset, row, colOffset, itemFieldValue);
                ++colOffset;
            }

            if (rowOffset > 0) {
                for (int i = templateRow.getFirstCellNum(); i < templateCol; ++i) {
                    newCell(templateRow, templateCol, rowOffset, row, i - templateCol, "");
                }
                for (int i = templateCol + colOffset, ii = templateRow.getLastCellNum(); i <= ii; ++i) {
                    newCell(templateRow, templateCol, rowOffset, row, i - templateCol, "");
                }
            }

            ++rowOffset;
        }
    }

    private void newCell(Row templateRow, int templateCol, int rowOffset, Row row, int colOffset, Object itemFieldValue) {
        Cell cell = null;
        if (rowOffset == 0) {
            cell = row.getCell(templateCol + colOffset);
        }

        if (cell == null) {
            cell = row.createCell(templateCol + colOffset);
            val templateRowCell = templateRow.getCell(templateCol + colOffset);
            if (templateRowCell != null) cell.setCellStyle(templateRowCell.getCellStyle());
        }

        cell.setCellValue("" + itemFieldValue);
    }

    private Cell findTemplateCell(ExcelRows excelRows) {
        if (StringUtils.isNotEmpty(excelRows.fromRef())) {
            val cellRef = new CellReference(excelRows.fromRef());
            return templateSheet.getRow(cellRef.getRow()).getCell(cellRef.getCol());
        }

        val cellRef = new CellReference(excelRows.fromColRef() + "1");
        for (int i = cellRef.getRow(); i <= templateSheet.getLastRowNum(); ++i) {
            val row = templateSheet.getRow(i);
            if (row == null) continue;

            val cell = row.getCell(cellRef.getCol());
            if (cell == null) continue;

            if (excelRows.fromKey().equals(cell.getStringCellValue())) {
                return cell;
            }
        }

        throw new RuntimeException("unable to find template row for fromColRef="
                + excelRows.fromColRef() + " and fromKey=" + excelRows.fromKey());
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

    @SneakyThrows
    private Object invokeField(Field field, Object bean) {
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
