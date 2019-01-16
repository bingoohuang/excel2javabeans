package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.*;
import com.github.bingoohuang.utils.lang.Mapp;
import com.github.bingoohuang.utils.lang.Str;
import com.github.bingoohuang.utils.reflect.Fields;
import com.github.bingoohuang.utils.type.Generic;
import com.google.common.collect.Maps;
import lombok.RequiredArgsConstructor;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import lombok.val;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.joor.Reflect;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

@RequiredArgsConstructor @Slf4j
public class BeansToExcelOnTemplate {
    private final Sheet sheet;   // 模板表单
    private Sheet optionsSheet;  // 选项表单
    private final Map<Integer, Integer> rowHeightRatioMap = Maps.newHashMap();   // RowIndex -> RowHeight ratio

    // 根据JavaBean，在模板页基础上生成Excel.
    @SuppressWarnings("unchecked")
    public Workbook create(Object bean) {
        optionsSheet = sheet.getWorkbook().getSheet("Options");

        for (val field : bean.getClass().getDeclaredFields()) {
            if (Fields.shouldIgnored(field, ExcelColIgnore.class)) continue;
            if (field.getName().endsWith("Tmpl")) continue;

            processExcelCellAnnotation(field, bean);
            processExcelRowsAnnotation(field, bean);
        }

        fixRowsHeight();
        PoiUtil.removeOtherSheets(sheet);


        val wb = sheet.getWorkbook();
        wb.setPrintArea(0, 0, PoiUtil.findMaxCol(sheet), 0, sheet.getLastRowNum());
        return wb;
    }

    /**
     * 修正行高。
     */
    private void fixRowsHeight() {
        for (int i = sheet.getFirstRowNum(), ii = sheet.getLastRowNum(); i <= ii; ++i) {
            val maxRows = rowHeightRatioMap.getOrDefault(i, 1);
            if (maxRows <= 1) continue;

            val row = sheet.getRow(i);
            if (row == null) continue;

            row.setHeight((short) (maxRows * row.getHeight()));
        }
    }

    /**
     * 处理@ExcelRows注解的属性。
     *
     * @param field JavaBean反射字段。
     * @param bean  字段所在的JavaBean。
     */
    @SneakyThrows @SuppressWarnings("unchecked")
    private void processExcelRowsAnnotation(Field field, Object bean) {
        val excelRows = field.getAnnotation(ExcelRows.class);
        if (excelRows == null) return;

        if (!Generic.of(field.getGenericType()).isRawType(List.class)) return;

        val templateCell = findTemplateCell(excelRows);
        if (templateCell == null) {
            log.warn("unable to locate template cell for field {}", field);
            return; // 找不到模板单元格，直接忽略字段处理。
        }

        val list = (List<Object>) Fields.invokeField(field, bean);
        val itemSize = PoiUtil.shiftRows(sheet, list, templateCell.getRowIndex());
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
            val from = PoiUtil.findCell(sheet, mergeColAnn.fromColRef() + tmplRowIndexRef);
            val to = PoiUtil.findCell(sheet, mergeColAnn.toColRef() + tmplRowIndexRef);

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

        for (val ann : excelRowsAnn.mergeRows()) {
            val fr = ann.fromRef();
            val cellRef = PoiUtil.isFullCellReference(fr) ? fr : fr + (tmplCellRowIndex + 1);
            val fromCell = PoiUtil.findCell(sheet, cellRef);

            val lastRow = tmplCellRowIndex + itemSize - 1;
            if (ann.type() == MergeType.Direct) {
                val fromRow = fromCell.getRow().getRowNum();
                directMergeRows(ann, fromRow, lastRow, fromCell.getColumnIndex());
            } else if (ann.type() == MergeType.SameValue) {
                sameValueMergeRows(ann, itemSize, fromCell);
            }
        }
    }

    /**
     * 同值纵向合并单元格。
     *
     * @param mergeRowAnn 纵向合并单元格注解。
     * @param itemSize    纵向合并行数。
     * @param fromCell    开始合并单元格。
     */
    private void sameValueMergeRows(MergeRow mergeRowAnn, int itemSize, Cell fromCell) {
        String lastValue = PoiUtil.getCellStringValue(fromCell);
        int preRow = fromCell.getRowIndex();
        val col = fromCell.getColumnIndex();
        int i = preRow + 1;

        for (final int ii = preRow + itemSize; i < ii; ++i) {
            val cell = sheet.getRow(i).getCell(col);
            val cellValue = PoiUtil.getCellStringValue(cell);
            if (StringUtils.equals(cellValue, lastValue)) continue;

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
     *
     * @param mergeRowAnn 纵向合并单元格注解。
     * @param fromRow     开始合并行索引。
     * @param col         纵向列索引。
     */
    private void fixCellValue(MergeRow mergeRowAnn, int fromRow, int col) {
        val sep = mergeRowAnn.prefixSeperate();
        if (StringUtils.isEmpty(sep)) return;

        val cell = sheet.getRow(fromRow).getCell(col);
        val old = PoiUtil.getCellStringValue(cell);
        val fixed = substringAfterSep(old, sep);

        PoiUtil.writeCellValue(cell, fixed);
    }

    public static String substringAfterSep(String str, String sep) {
        if (StringUtils.isEmpty(str)) return str;

        int pos = str.indexOf(sep);
        return pos < 0 ? str : str.substring(pos + sep.length());
    }

    /**
     * 根据JavaBean列表，向Excel中写入多行。
     *
     * @param templateCell 模板单元格。
     * @param items        写入JavaBean列表。
     */
    @SuppressWarnings("unchecked")
    private void writeRows(Cell templateCell, List<Object> items) {
        val tmplRow = templateCell.getRow();
        val fromRow = tmplRow.getRowNum();

        val tmplCol = templateCell.getColumnIndex();
        for (int i = 0, ii = items.size(); i < ii; ++i) {
            val item = items.get(i);
            val row = i == 0 ? tmplRow : sheet.createRow(fromRow + i);
            if (row != tmplRow) {
                row.setHeight(tmplRow.getHeight());
            }

            val fields = item.getClass().getDeclaredFields();
            int cutoff = 0;
            for (int j = 0; j < fields.length; ++j) {
                val field = fields[j];
                if (Fields.shouldIgnored(field, ExcelColIgnore.class)) {
                    ++cutoff;
                    continue;
                }
                if (field.getName().endsWith("Tmpl")) {
                    ++cutoff;
                    continue;
                }

                val fv = Fields.invokeField(field, item);
                val excelCell = field.getAnnotation(ExcelCell.class);
                int maxLen = excelCell == null ? 0 : excelCell.maxLineLen();
                val cell = newCell(tmplRow, tmplCol + j - cutoff, i, row, fv, maxLen);
                applyTemplateCellStyle(field, item, excelCell, cell);
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
            newCell(tmplRow, i, rowOffset, row, "", 0);
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
     * @return 新的单元格。
     */
    private Cell newCell(Row tmplRow, int cellCol, int rowOffset, Row row, Object cellValue, int maxLen) {
        val cell = getOrCreateCell(tmplRow, cellCol, rowOffset, row);

        val value = PoiUtil.writeCellValue(cell, cellValue);

        fixMaxRowHeightRatio(row, maxLen, value);
        return cell;
    }

    private Cell getOrCreateCell(Row tmplRow, int cellCol, int rowOffset, Row row) {
        // 偏移量为0，说明当前在模板行上，尝试直接获取单元格
        if (rowOffset == 0) return row.getCell(cellCol);

        val cell = row.createCell(cellCol);
        // 从对应列的模板单元格套取样式
        val tmplRowCell = tmplRow.getCell(cellCol);
        if (tmplRowCell != null) cell.setCellStyle(tmplRowCell.getCellStyle());

        return cell;
    }

    private void fixMaxRowHeightRatio(Row row, int maxLen, String value) {
        if (maxLen <= 0) return;

        val max = rowHeightRatioMap.getOrDefault(row.getRowNum(), 1);
        int rows = (int) Math.ceil(value.length() * 1.0 / maxLen);
        if (rows > max) rowHeightRatioMap.put(row.getRowNum(), rows);
    }

    /**
     * 查找模板单元格。
     *
     * @param excelRowsAnn ExcelRows注解
     * @return 模板单元格。返回null，没有找到。
     */
    private Cell findTemplateCell(ExcelRows excelRowsAnn) {
        val fromRef = excelRowsAnn.fromRef();
        if (PoiUtil.isFullCellReference(fromRef)) {
            val cellRef = new CellReference(fromRef);
            return sheet.getRow(cellRef.getRow()).getCell(cellRef.getCol());
        }

        val cellRef = new CellReference(fromRef + "1");
        val col = cellRef.getCol();
        val searchKey = excelRowsAnn.searchKey();
        for (int i = cellRef.getRow(), ii = sheet.getLastRowNum(); i <= ii; ++i) {
            val row = sheet.getRow(i);
            if (row == null) continue;
            val cell = row.getCell(col);
            if (cell == null) continue;

            val cellValue = PoiUtil.getCellStringValue(cell);
            if (StringUtils.contains(cellValue, searchKey)) return cell;
        }

        return null;
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

        Object fv = Fields.invokeField(field, bean);

        if (ann.sheetName()) {
            if (StringUtils.isNotEmpty(ann.replace())) { // 有内容需要替换
                fv = sheet.getSheetName().replace(ann.replace(), Str.nullThen(fv, ""));
            }

            val wb = sheet.getWorkbook();
            val oldSheetName = sheet.getSheetName();
            val newSheetName = "" + fv;
            wb.setSheetName(wb.getSheetIndex(sheet), newSheetName);

            PoiUtil.fixChartSheetNameRef(sheet, oldSheetName, newSheetName);
        } else {
            val cell = PoiUtil.findCell(sheet, ann.value(), StringUtils.defaultIfEmpty(ann.searchKey(), "{" + field.getName() + "}"));
            if (cell == null) {
                log.warn("unable to find cell for {} in field {}", ann, field);
                return;
            }

            if (StringUtils.isNotEmpty(ann.replace())) { // 有内容需要替换
                val old = PoiUtil.getCellStringValue(cell);
                fv = old.replace(ann.replace(), Str.nullThen(fv, ""));
            }

            applyTemplateCellStyle(field, bean, ann, cell);

            val strCellValue = PoiUtil.writeCellValue(cell, fv);
            fixMaxRowHeightRatio(cell.getRow(), ann.maxLineLen(), strCellValue);
        }
    }

    private void applyTemplateCellStyle(Field field, Object bean, ExcelCell excelCell, Cell cell) {
        if (excelCell == null || optionsSheet == null) return;

        val templateCells = excelCell.templateCells();
        if (templateCells.length == 0) return;

        val templateCellMap = Mapp.createMap(":", templateCells);

        String tmplName = Reflect.on(bean).get(field.getName() + "Tmpl");
        val tmplCellReference = Mapp.firstNonNull(templateCellMap, tmplName, "DEFAULT");

        val tmplCell = PoiUtil.findCell(optionsSheet, tmplCellReference);
        cell.setCellStyle(tmplCell.getCellStyle());
    }
}
