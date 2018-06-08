package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.CellData.CellDataBuilder;
import com.github.bingoohuang.util.instantiator.BeanInstantiator;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Table;
import lombok.RequiredArgsConstructor;
import lombok.val;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;

import java.text.SimpleDateFormat;
import java.util.List;
import java.util.Map;

class RowObjectCreator<T> {
    private final List<ExcelBeanField> beanFields;
    private final boolean cellDataMapAttachable;
    private final Sheet sheet;
    private final Row row;
    private final Table<Integer, Integer, ImageData> imageDataTable;
    private final DataFormatter cellFormatter;
    private final Map<String, CellData> cellDataMap;
    private final T object;

    private int emptyNum;

    public RowObjectCreator(BeanInstantiator<T> instantiator,
                            List<ExcelBeanField> beanFields,
                            boolean cellDataMapAttachable,
                            Sheet sheet,
                            Table<Integer, Integer, ImageData> imageDataTable,
                            DataFormatter cellFormatter,
                            int rowNum) {
        this.beanFields = beanFields;
        this.cellDataMapAttachable = cellDataMapAttachable;
        this.cellDataMap = cellDataMapAttachable ? Maps.newHashMap() : null;
        this.sheet = sheet;
        this.imageDataTable = imageDataTable;
        this.cellFormatter = cellFormatter;
        this.row = sheet.getRow(rowNum);
        this.object = this.row == null ? null : (T) instantiator.newInstance();
    }

    public T createObject() {
        if (object == null) return null;

        processRow();
        if (emptyNum == beanFields.size()) return null;
        if (cellDataMapAttachable)
            ((CellDataMapAttachable) object).attachCellDataMap(cellDataMap);

        return object;
    }


    private void processRow() {
        for (val beanField : beanFields) {
            val fieldValue = new BeanFieldValueCreator(beanField).parseFieldValue();

            if (fieldValue == null) {
                ++emptyNum;
            } else {
                beanField.setFieldValue(object, fieldValue);
            }
        }
    }

    @RequiredArgsConstructor
    private class BeanFieldValueCreator {
        private final ExcelBeanField beanField;

        public Object parseFieldValue() {
            return beanField.isMultipleColumns() ? parseMultipleFieldValue()
                    : processSingleColumn(beanField.getColumnIndex(), -1);
        }

        private Object parseMultipleFieldValue() {
            int nonEmptyFieldValues = 0;
            List<Object> fieldValues = Lists.newArrayList();
            for (int columnIndex : beanField.getMultipleColumnIndexes()) {
                val value = processSingleColumn(columnIndex, fieldValues.size());
                fieldValues.add(value);

                if (value != null) ++nonEmptyFieldValues;
            }

            return nonEmptyFieldValues > 0 ? fieldValues : null;
        }


        private Object processSingleColumn(int columnIndex, int fieldNameIndex) {
            if (columnIndex < 0) return null;

            val cell = row.getCell(columnIndex);
            if (beanField.isImageDataField()) {
                attachCellDataMap(columnIndex, fieldNameIndex, cell);
                return imageDataTable.get(row.getRowNum(), columnIndex);
            } else {
                val cellValue = getCellValue(cell);
                return convertCellValue(cell, cellValue, row.getRowNum(), columnIndex, fieldNameIndex);
            }
        }


        private void attachCellDataMap(int columnIndex, int fieldNameIndex, Cell cell) {
            if (!cellDataMapAttachable) return;

            val attachFieldName = createAttachFieldName(fieldNameIndex);
            val cellData = createCellData(cell, null, row.getRowNum(), columnIndex);
            cellDataMap.put(attachFieldName, cellData);
        }

        private String createAttachFieldName(int fieldNameIndex) {
            val fieldName = beanField.getFieldName();
            return fieldNameIndex < 0 ? fieldName : fieldName + "_" + fieldNameIndex;
        }

        private Object convertCellValue(Cell cell, String cellValue, int rowNum, int colIndex, int fieldNameIndex) {
            CellData cd = null;
            if (beanField.isCellDataType() || cellDataMapAttachable) {
                cd = createCellData(cell, cellValue, rowNum, colIndex);
            }

            if (cellDataMapAttachable)
                cellDataMap.put(createAttachFieldName(fieldNameIndex), cd);

            if (StringUtils.isEmpty(cellValue)) return null;

            return beanField.isCellDataType() ? cd : beanField.convert(cellValue);
        }
    }


    private CellData createCellData(Cell cell, String cellValue, int rowNum, int colNum) {
        val builder = CellData.builder().value(cellValue).row(rowNum).col(colNum)
                .sheetIndex(sheet.getWorkbook().getSheetIndex(sheet));

        return applyComment(cell, builder).build();
    }

    private CellDataBuilder applyComment(Cell cell, CellDataBuilder cellData) {
        if (cell == null) return cellData;

        val comment = cell.getCellComment();
        if (comment == null) return cellData;

        return cellData.comment(comment.getString().getString()).commentAuthor(comment.getAuthor());
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return null;

        if (cell.getCellTypeEnum() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            return new SimpleDateFormat("yyyy-MM-dd").format(cell.getDateCellValue());
        }

        return StringUtils.trim(cellFormatter.formatCellValue(cell));
    }
}



