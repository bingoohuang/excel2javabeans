package com.github.bingoohuang.excel2beans;

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

        if (cellDataMapAttachable) cellDataMap = Maps.newHashMap();
        else cellDataMap = null;

        this.sheet = sheet;
        this.imageDataTable = imageDataTable;
        this.cellFormatter = cellFormatter;
        this.row = sheet.getRow(rowNum);
        this.object = this.row == null ? null : (T) instantiator.newInstance();
    }

    public T createObject() {
        if (object == null) return null;

        processRow();
        if (emptyNum == beanFields.size()) {
            return null;
        }

        if (cellDataMapAttachable) {
            ((CellDataMapAttachable) object).attachCellDataMap(cellDataMap);
        }

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
            if (beanField.isMultipleColumns()) {
                return parseMultipleFieldValue();
            } else {
                return processSingleColumn(beanField.getColumnIndex(), -1);
            }
        }

        private Object parseMultipleFieldValue() {
            int nonEmptyFieldValues = 0;
            val fieldValues = Lists.<Object>newArrayList();
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

        private Object convertCellValue(Cell cell, String cellValue, int rowNum,
                                        int columnIndex, int fieldNameIndex) {

            CellData cellData = null;
            if (beanField.isCellDataType() || cellDataMapAttachable) {
                cellData = createCellData(cell, cellValue, rowNum, columnIndex);
            }

            if (cellDataMapAttachable) {
                val attachFieldName = createAttachFieldName(fieldNameIndex);
                cellDataMap.put(attachFieldName, cellData);
            }

            if (StringUtils.isEmpty(cellValue)) return null;

            return beanField.isCellDataType() ? cellData : beanField.convert(cellValue);
        }
    }


    private CellData createCellData(Cell cell, String cellValue, int rowNum, int colNum) {
        val cellDataBuilder = CellData.builder()
                .value(cellValue).row(rowNum).col(colNum)
                .sheetIndex(sheet.getWorkbook().getSheetIndex(sheet));
        applyComment(cell, cellDataBuilder);
        return cellDataBuilder.build();
    }

    private void applyComment(Cell cell, CellData.CellDataBuilder cellData) {
        if (cell == null) return;

        val comment = cell.getCellComment();
        if (comment == null) return;

        cellData.comment(comment.getString().getString())
                .commentAuthor(comment.getAuthor());
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return null;

        val cellType = cell.getCellTypeEnum();
        if (cellType == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            val dateCellValue = cell.getDateCellValue();
            val sdf = new SimpleDateFormat("yyyy-MM-dd");
            return sdf.format(dateCellValue);
        }

        val cellValue = cellFormatter.formatCellValue(cell);
        return StringUtils.trim(cellValue);
    }
}



