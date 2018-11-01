package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelCell;
import com.github.bingoohuang.excel2beans.annotations.ExcelRows;
import com.github.bingoohuang.excel2beans.annotations.MergeRow;
import com.github.bingoohuang.excel2beans.annotations.MergeType;
import com.github.bingoohuang.util.GenericType;
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
import java.util.List;

@RequiredArgsConstructor
public class BeansToExcelOnTemplate {
    // 模板表单
    private final Sheet sheet;

    // 根据JavaBean，在模板页基础上生成Excel.
    public Workbook create(Object bean) {
        for (val field : bean.getClass().getDeclaredFields()) {
            processExcelCellAnnotation(field, bean);
            processExcelRowsAnnotation(field, bean);
        }

        removeOtherSheets();
        return sheet.getWorkbook();
    }

    /**
     * 处理@ExcelRows注解的属性。
     *
     * @param field JavaBean反射字段。
     * @param bean  字段所在的JavaBean。
     */
    @SneakyThrows
    private void processExcelRowsAnnotation(Field field, Object bean) {
        val excelRows = field.getAnnotation(ExcelRows.class);
        if (excelRows == null) return;

        if (!GenericType.of(field.getGenericType()).isRawType(List.class)) return;

        val templateCell = findTemplateCell(excelRows);
        @SuppressWarnings("unchecked")
        val list = (List<Object>) invokeField(field, bean);

        val itemSize = shiftRows(templateCell, list);
        if (itemSize > 0) {
            writeRows(templateCell, list);
            mergeRows(excelRows, templateCell, itemSize);
            mergeCols(excelRows, templateCell, itemSize);
        }
    }

    /**
     * 根据@ExcelRows注解的指示，合并横向合并单元格。
     *
     * @param excelRowsAnn @ExcelRows注解值。
     * @param templateCell 模板单元格。
     * @param itemSize     在多少行上横向合并单元格。
     */
    private void mergeCols(ExcelRows excelRowsAnn, Cell templateCell, int itemSize) {
        val tmplRowIndexRef = templateCell.getRowIndex() + 1;
        for (val mergeColAnn : excelRowsAnn.mergeCols()) {
            val from = findCell(mergeColAnn.fromColRef() + tmplRowIndexRef);
            val to = findCell(mergeColAnn.toColRef() + tmplRowIndexRef);

            val fromRow = from.getRowIndex();
            val fromCol = from.getColumnIndex();
            val toCol = to.getColumnIndex();

            for (int i = fromRow; i < fromRow + itemSize; ++i) {
                sheet.addMergedRegion(new CellRangeAddress(i, i, fromCol, toCol));
            }
        }
    }

    /**
     * 根据@ExcelRows注解的指示，纵向合并单元格。
     *
     * @param excelRowsAnn @ExcelRows注解值。
     * @param templateCell 模板单元格。
     * @param itemSize     纵向合并多少行的单元格。
     */
    private void mergeRows(ExcelRows excelRowsAnn, Cell templateCell, int itemSize) {
        val tmplCellRowIndex = templateCell.getRowIndex();

        for (val mergeRowAnn : excelRowsAnn.mergeRows()) {
            val fr = mergeRowAnn.fromRef();
            val cellRef = isFullCellReference(fr) ? fr : fr + (tmplCellRowIndex + 1);
            val fromCell = findCell(cellRef);


            val lastRow = tmplCellRowIndex + itemSize - 1;
            if (mergeRowAnn.type() == MergeType.Direct) {
                val fromRow = fromCell.getRow().getRowNum();
                directMergeRows(mergeRowAnn, fromRow, lastRow, fromCell.getColumnIndex());
            } else if (mergeRowAnn.type() == MergeType.SameValue) {
                sameValueMergeRows(mergeRowAnn, itemSize, fromCell);
            }
        }
    }

    /**
     * 是否是完整的单元格索引。例如A5是完整的，A不是完整的，缺少行标。
     *
     * @param cellReference EXCEL中的单元格索引。
     * @return 完整的单元格索引时返回true。
     */
    private boolean isFullCellReference(String cellReference) {
        return cellReference.matches("\\w+\\d+");
    }

    /**
     * 同值纵向合并单元格。
     *
     * @param mergeRowAnn 纵向合并单元格注解。
     * @param itemSize    纵向合并行数。
     * @param fromCell    开始合并单元格。
     */
    private void sameValueMergeRows(MergeRow mergeRowAnn, int itemSize, Cell fromCell) {
        String lastValue = fromCell.getStringCellValue();
        val col = fromCell.getColumnIndex();
        int preRow = fromCell.getRowIndex();
        int i = preRow + 1;
        for (int ii = preRow + itemSize; i < ii; ++i) {
            val c = sheet.getRow(i).getCell(col);
            val cellValue = c.getStringCellValue();
            if (cellValue.equals(lastValue)) continue;

            directMergeRows(mergeRowAnn, preRow, i - 1, col);

            lastValue = cellValue;
            preRow = i;
        }

        directMergeRows(mergeRowAnn, preRow, i - 1, col);
    }

    /**
     * 直接纵向合并单元格。
     *
     * @param mergeRowAnn 纵向合并单元格注解。
     * @param fromRow     开始合并行索引。
     * @param lastRow     结束合并行索引。
     * @param col         纵向列索引。
     */
    private void directMergeRows(MergeRow mergeRowAnn, int fromRow, int lastRow, int col) {
        fixCellValue(mergeRowAnn, fromRow, col);

        if (lastRow < fromRow) return;
        if (lastRow == fromRow && mergeRowAnn.moreCols() == 0) return;

        val c1 = new CellRangeAddress(fromRow, lastRow, col, col + mergeRowAnn.moreCols());
        sheet.addMergedRegion(c1);
    }

    /**
     * 修复合并单元格的数据。
     * 1. 去除指导合并的取值前缀，例如1^a中的1^;
     * 2. 修正合并后单元格的数值型取值。
     *
     * @param mergeRowAnn 纵向合并单元格注解。
     * @param fromRow     开始合并行索引。
     * @param col         纵向列索引。
     */
    private void fixCellValue(MergeRow mergeRowAnn, int fromRow, int col) {
        val sep = mergeRowAnn.prefixSeperate();
        if (StringUtils.isEmpty(sep)) return;

        val cell = sheet.getRow(fromRow).getCell(col);
        val old = cell.getStringCellValue();
        val indexOf = old.indexOf(sep);
        val fixed = indexOf < 0 ? old : old.substring(indexOf + sep.length());

        val numeric = ExcelToBeansUtils.isNumeric(fixed);
        writeCellValue(cell, numeric ? Double.parseDouble(fixed) : fixed);
    }

    /**
     * 插入新的行（向下移动表格中的行），方便写入数据。
     *
     * @param templateCell 模板单元格。
     * @param list         写入JavaBean列表。
     * @return JavaBean列表大小。
     */
    private int shiftRows(Cell templateCell, List<Object> list) {
        val itemSize = list == null ? 0 : list.size();
        int fromRow = templateCell.getRow().getRowNum();
        int lastRowNum = sheet.getLastRowNum();
        if (itemSize == 0 && fromRow == lastRowNum) { // 没有写入行，直接删除模板行
            sheet.removeRow(sheet.getRow(fromRow));
        } else if (itemSize != 1 && fromRow < lastRowNum) {
            sheet.shiftRows(fromRow + 1, lastRowNum, itemSize - 1);
        }

        return itemSize;
    }

    /**
     * 根据JavaBean列表，向Excel中写入多行。
     *
     * @param templateCell 模板单元格。
     * @param items        写入JavaBean列表。
     */
    private void writeRows(Cell templateCell, List<Object> items) {
        val tmplRow = templateCell.getRow();
        val fromRow = tmplRow.getRowNum();

        val tmplCol = templateCell.getColumnIndex();
        for (int i = 0, ii = items.size(); i < ii; ++i) {
            val item = items.get(i);
            val row = i == 0 ? tmplRow : sheet.createRow(fromRow + i);

            val fields = item.getClass().getDeclaredFields();
            for (int j = 0; j < fields.length; ++j) {
                val fv = invokeField(fields[j], item);
                newCell(tmplRow, tmplCol + j, i, row, fv);
            }

            emptyEndsCells(tmplRow, tmplCol, i, row, fields.length);
        }
    }

    /**
     * 置空写入行两端单元格。
     *
     * @param tmplRow   模板行。
     * @param tmplCol   模板单元格所在列。
     * @param rowOffset 写入行偏移号。
     * @param row       写入行。
     * @param fieldsNum 写入JavaBean属性的数量。
     */
    private void emptyEndsCells(Row tmplRow, int tmplCol, int rowOffset, Row row, int fieldsNum) {
        if (rowOffset <= 0) return;

        emptyCells(tmplRow, rowOffset, row, tmplRow.getFirstCellNum(), tmplCol - 1);
        emptyCells(tmplRow, rowOffset, row, tmplCol + fieldsNum, tmplRow.getLastCellNum());
    }

    /**
     * 置空非JavaBean属性字段关联的单元格。
     *
     * @param tmplRow   模板行。
     * @param rowOffset 写入行偏移号。
     * @param row       需要创建新单元格所在的行。
     * @param colStart  开始列索引。
     * @param colEnd    结束列索引。
     */
    private void emptyCells(Row tmplRow, int rowOffset, Row row, int colStart, int colEnd) {
        for (int i = colStart; i <= colEnd; ++i) {
            newCell(tmplRow, i, rowOffset, row, "");
        }
    }

    /**
     * 创建新的单元格。
     *
     * @param tmplRow   模板行。
     * @param cellCol   单元格所在的列索引。
     * @param rowOffset 行偏移号。
     * @param row       需要创建新单元格所在的行。
     * @param cellValue 新单元格取值。
     */
    private void newCell(Row tmplRow, int cellCol, int rowOffset, Row row, Object cellValue) {
        Cell cell = null;
        if (rowOffset == 0) { // 偏移量为0，说明当前在模板行上，尝试直接获取单元格
            cell = row.getCell(cellCol);
        }

        if (cell == null) { // 获取不到时，创建新的单元格
            cell = row.createCell(cellCol);

            // 并且从对应列的模板单元格套取样式
            val tmplRowCell = tmplRow.getCell(cellCol);
            if (tmplRowCell != null) cell.setCellStyle(tmplRowCell.getCellStyle());
        }

        writeCellValue(cell, cellValue);
    }

    /**
     * 查找模板单元格。
     *
     * @param excelRowsAnn ExcelRows注解
     * @return 模板单元格。
     */
    private Cell findTemplateCell(ExcelRows excelRowsAnn) {
        if (isFullCellReference(excelRowsAnn.fromRef())) {
            val cellRef = new CellReference(excelRowsAnn.fromRef());
            return sheet.getRow(cellRef.getRow()).getCell(cellRef.getCol());
        }

        val cellRef = new CellReference(excelRowsAnn.fromRef() + "1");
        for (int i = cellRef.getRow(); i <= sheet.getLastRowNum(); ++i) {
            val row = sheet.getRow(i);
            if (row == null) continue;

            val cell = row.getCell(cellRef.getCol());
            if (cell == null) continue;

            val cellValue = cell.getStringCellValue();
            if (StringUtils.contains(cellValue, excelRowsAnn.searchKey())) return cell;
        }

        throw new RuntimeException("unable to find template row for fromColRef="
                + excelRowsAnn.fromRef() + " and searchKey=" + excelRowsAnn.searchKey());
    }

    /**
     * 处理@ExcelCell注解的字段。
     *
     * @param field JavaBean反射字段。
     * @param bean  字段所在的JavaBean。
     */
    @SneakyThrows
    private void processExcelCellAnnotation(Field field, Object bean) {
        val ann = field.getAnnotation(ExcelCell.class);
        if (ann == null) return;

        Object fv = invokeField(field, bean);
        if (ann.sheetName()) {
            if (StringUtils.isNotEmpty(ann.replace())) { // 有内容需要替换
                fv = sheet.getSheetName().replace(ann.replace(), "" + fv);
            }

            val wb = sheet.getWorkbook();
            wb.setSheetName(wb.getSheetIndex(sheet), "" + fv);
        } else {
            val cell = findCell(ann.value());
            if (StringUtils.isNotEmpty(ann.replace())) { // 有内容需要替换
                val old = cell.getStringCellValue();
                fv = old.replace(ann.replace(), "" + fv);
            }

            writeCellValue(cell, fv);
        }
    }

    /**
     * 向单元格写入值，处理值为整型时的写入情况。
     *
     * @param cell 单元格。
     * @param fv   单元格值。
     */
    private void writeCellValue(Cell cell, Object fv) {
        if (fv instanceof Number) {
            cell.setCellValue(((Number) fv).doubleValue());
        } else {
            cell.setCellValue("" + fv);
        }
    }

    /**
     * 获取字段取值（null时，转换为长度为空字符串）。
     *
     * @param field JavaBean反射字段。
     * @param bean  字段所在的JavaBean。
     * @return 字段取值。
     */
    @SneakyThrows
    private Object invokeField(Field field, Object bean) {
        if (!field.isAccessible()) field.setAccessible(true);
        val fieldValue = field.get(bean);

        return fieldValue == null ? "" : fieldValue;
    }

    /**
     * 根据单元格索引，找到单元格。
     *
     * @param cellRefValue 单元格索引，例如A1, AB12等。
     * @return 单元格。
     */
    private Cell findCell(String cellRefValue) {
        val cellRef = new CellReference(cellRefValue);
        val row = sheet.getRow(cellRef.getRow());
        return row.getCell(cellRef.getCol());
    }

    /**
     * 删除其它Sheet页，只保留指定模板的Sheet页输出。
     */
    private void removeOtherSheets() {
        val workbook = sheet.getWorkbook();

        List<String> deleteSheetNames = Lists.newArrayList();
        for (int i = 0; i < workbook.getNumberOfSheets(); ++i) {
            val sheetName = workbook.getSheetName(i);
            if (!sheetName.equals(sheet.getSheetName())) {
                deleteSheetNames.add(sheetName);
            }
        }

        deleteSheetNames.forEach(x -> workbook.removeSheetAt(workbook.getSheetIndex(x)));
    }
}
