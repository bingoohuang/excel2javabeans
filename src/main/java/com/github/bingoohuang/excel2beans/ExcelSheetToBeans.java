package com.github.bingoohuang.excel2beans;

import com.esotericsoftware.reflectasm.FieldAccess;
import com.esotericsoftware.reflectasm.MethodAccess;
import com.github.bingoohuang.excel2beans.annotations.ExcelSheet;
import com.google.common.collect.Lists;
import lombok.val;
import org.apache.poi.ss.usermodel.*;
import org.objenesis.ObjenesisStd;
import org.objenesis.instantiator.ObjectInstantiator;

import java.util.List;

import static org.apache.commons.lang3.StringUtils.isEmpty;
import static org.apache.commons.lang3.StringUtils.trim;

public class ExcelSheetToBeans<T> {
    private final Workbook workbook;

    private final FieldAccess fieldAccess;
    private final MethodAccess methodAccess;
    private final ObjectInstantiator<T> instantiator;
    private final ExcelBeanField[] beanFields;
    private final boolean hasTitle;
    private final DataFormatter cellFormatter = new DataFormatter();
    private final Sheet sheet;

    public ExcelSheetToBeans(Workbook workbook, Class<T> beanClass) {
        this.workbook = workbook;
        this.fieldAccess = FieldAccess.get(beanClass);
        this.methodAccess = MethodAccess.get(beanClass);
        this.instantiator = new ObjenesisStd().getInstantiatorOf(beanClass);
        this.sheet = selectSheet(beanClass);
        this.beanFields = ExcelToBeansUtils.parseBeanFields(beanClass, null);
        this.hasTitle = hasTitle();
    }

    private Sheet selectSheet(Class<T> beanClass) {
        val excelSheet = beanClass.getAnnotation(ExcelSheet.class);
        if (excelSheet == null) {
            return workbook.getSheetAt(0);
        }

        for (int i = 0, ii = workbook.getNumberOfSheets(); i < ii; ++i) {
            val sheetName = workbook.getSheetName(i);
            if (sheetName.contains(excelSheet.name())) {
                return workbook.getSheetAt(i);
            }
        }

        throw new IllegalArgumentException("Unable to find sheet with name " + excelSheet.name());
    }

    public List<T> convert() {
        List<T> beans = Lists.newArrayList();

        val startRowNum = jumpToStartDataRow(sheet);
        for (int i = startRowNum, ii = sheet.getLastRowNum(); i <= ii; ++i) {
            T object = createObject(sheet, i);
            if (object != null) {
                addToBeans(beans, i, object);
            }
        }

        return beans;
    }

    private T createObject(Sheet sheet, int i) {
        T object = null;

        val row = sheet.getRow(i);
        if (row != null) {
            object = instantiator.newInstance();
            int emptyNum = processRow(object, row);
            if (emptyNum == beanFields.length) {
                object = null;
            }
        }

        return object;
    }

    private void addToBeans(List<T> beans, int i, T object) {
        if (object instanceof ExcelRowIgnorable) {
            val ignore = (ExcelRowIgnorable) object;
            if (ignore.ignoreRow()) {
                return;
            }
        }

        if (object instanceof ExcelRowRef) {
            val ref = (ExcelRowRef) object;
            ref.setRowNum(i);
        }

        beans.add(object);
    }

    private int processRow(T object, Row row) {
        int emptyNum = 0;
        for (int j = 0; j < beanFields.length; ++j) {
            val beanField = beanFields[j];
            val cell = row.getCell(beanField.getColumnIndex());
            val cellStringValue = getCellValue(cell);
            if (isEmpty(cellStringValue)) {
                ++emptyNum;
            } else {
                val cellValue = convertCellValue(beanField, cell, cellStringValue, row.getRowNum());
                beanField.setFieldValue(fieldAccess, methodAccess, object, cellValue);
            }
        }

        return emptyNum;
    }

    private Object convertCellValue(ExcelBeanField beanField, Cell cell, String cellValue, int rowNum) {
        if (beanField.isCellDataType()) {
            val cellData = CellData.builder()
                    .value(cellValue)
                    .row(rowNum)
                    .col(cell.getColumnIndex())
                    .sheetIndex(workbook.getSheetIndex(sheet));
            applyComment(cell, cellData);
            return cellData.build();
        } else {
            Class<?> type = beanField.getField().getType();
            if (type == int.class || type == Integer.class) {
                return Integer.parseInt(cellValue);
            }
        }

        return cellValue;
    }

    private void applyComment(Cell cell, CellData.CellDataBuilder cellData) {
        val comment = cell.getCellComment();
        if (comment == null) {
            return;
        }

        cellData.comment(comment.getString().getString())
                .commentAuthor(comment.getAuthor());
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return null;

        val cellValue = cellFormatter.formatCellValue(cell);
        return trim(cellValue);
    }


    private int jumpToStartDataRow(Sheet sheet) {
        int i = sheet.getFirstRowNum();
        if (!hasTitle) return i;

        // try to find the title row
        for (int ii = sheet.getLastRowNum(); i <= ii; ++i) {
            val row = sheet.getRow(i);

            boolean containsTitle = false;
            for (int j = 0; j < beanFields.length; ++j) {
                val beanField = beanFields[j];
                if (!beanField.hasTitle()) {
                    beanField.setColumnIndex(j + row.getFirstCellNum());
                } else {
                    if (findColumn(row, beanField)) {
                        containsTitle = true;
                    }
                }
            }

            if (containsTitle) {
                checkTitleColumnsAllFound();
                return i + 1;
            }
        }

        throw new IllegalArgumentException("找不到标题行");
    }

    private void checkTitleColumnsAllFound() {
        for (int j = 0; j < beanFields.length; ++j) {
            val beanField = beanFields[j];
            if (beanField.hasTitle() && !beanField.isTitleColumnFound()) {
                throw new IllegalArgumentException("找不到[" + beanField.getTitle() + "]的列");
            }
        }
    }

    private boolean findColumn(Row row, ExcelBeanField beanField) {
        for (int k = row.getFirstCellNum(), kk = row.getLastCellNum(); k <= kk; ++k) {
            val cell = row.getCell(k);
            if (cell == null) continue;

            val cellValue = cell.getStringCellValue();
            if (beanField.containTitle(cellValue)) {
                beanField.setColumnIndex(cell.getColumnIndex());
                beanField.setTitleColumnFound(true);
                return true;
            }
        }

        return false;
    }

    private boolean hasTitle() {
        for (val beanField : beanFields) {
            if (beanField.hasTitle()) return true;
        }

        return false;
    }
}
