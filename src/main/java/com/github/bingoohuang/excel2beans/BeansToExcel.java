package com.github.bingoohuang.excel2beans;

import com.esotericsoftware.reflectasm.FieldAccess;
import com.esotericsoftware.reflectasm.MethodAccess;
import com.github.bingoohuang.excel2beans.annotations.ExcelSheet;
import com.google.common.collect.Maps;
import lombok.val;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;
import java.util.Map;

/**
 * Created by bingoohuang on 2017/3/20.
 */
public class BeansToExcel {
    private Workbook templateWorkbook;

    public BeansToExcel() {
    }

    public BeansToExcel(Workbook templateWorkbook) {
        this.templateWorkbook = templateWorkbook;
    }

    public Workbook create(List<?>... lists) {
        val wb = new XSSFWorkbook();

        Map<Class, BeanClassBag> sheets = Maps.newHashMap();

        for (val list : lists) {
            for (val bean : list) {
                val bag = insureBagCreated(wb, sheets, bean);

                Sheet sheet = bag.getSheet();
                Row row = sheet.createRow(sheet.getLastRowNum() + 1);
                writeRowCells(bean, bag, row);
            }
        }

        autoSizeColumn(sheets);

        return wb;
    }

    private void autoSizeColumn(Map<Class, BeanClassBag> sheets) {
        for (val bag : sheets.values()) {
            Sheet sheet = bag.getSheet();
            short lastCellNum = sheet.getRow(0).getLastCellNum();
            for (int i = 0; i <= lastCellNum; ++i) {
                sheet.autoSizeColumn(i); // adjust width of the column
            }
        }
    }

    private void writeRowCells(Object bean, BeanClassBag bag, Row row) {
        for (int i = 0, ii = bag.getBeanFields().length; i < ii; ++i) {
            Cell cell = row.createCell(i);

            val field = bag.getBeanFields()[i];
            Object fieldValue = field.getFieldValue(bag.getFieldAccess(), bag.getMethodAccess(), bean);
            cell.setCellValue(String.valueOf(fieldValue));
            cell.setCellStyle(field.getCellStyle());
        }
    }

    private BeanClassBag insureBagCreated(Workbook wb, Map<Class, BeanClassBag> map, Object bean) {
        Class<?> beanClass = bean.getClass();
        BeanClassBag bag = map.get(beanClass);
        if (bag != null) return bag;

        bag = new BeanClassBag();
        bag.setBeanClass(beanClass);
        bag.setSheet(createSheet(wb, beanClass));
        bag.setBeanFields(ExcelToBeansUtils.parseBeanFields(beanClass));
        bag.setMethodAccess(MethodAccess.get(beanClass));
        bag.setFieldAccess(FieldAccess.get(beanClass));

        addTitleToSheet(bag);

        map.put(beanClass, bag);

        return bag;
    }

    private void addTitleToSheet(BeanClassBag bag) {
        Row row = bag.getSheet().createRow(0);
        ExcelBeanField[] beanFields = bag.getBeanFields();
        for (int i = 0, ii = beanFields.length; i < ii; ++i) {
            ExcelBeanField beanField = beanFields[i];
            Cell cell = row.createCell(i);
            cell.setCellValue(beanField.getTitle());
        }

        cloneCellStyle(bag, row, beanFields);
    }

    private void cloneCellStyle(BeanClassBag bag, Row row, ExcelBeanField[] beanFields) {
        if (templateWorkbook == null) return;

        String sheetName = bag.getSheet().getSheetName();
        Sheet templateSheet = templateWorkbook.getSheet(sheetName);
        if (templateSheet == null) {
            templateSheet = templateWorkbook.getSheetAt(0);
        }

        for (int i = 0, ii = beanFields.length; i < ii; ++i) {
            val titleStyle = cloneCellStyle(row, templateSheet, 0, i);
            row.getCell(i).setCellStyle(titleStyle);

            val dataStyle = cloneCellStyle(row, templateSheet, 1, i);
            beanFields[i].setCellStyle(dataStyle);
        }
    }

    private CellStyle cloneCellStyle(Row row, Sheet templateSheet, int rowIndex, int colIndex) {
        Row templateRow = templateSheet.getRow(rowIndex);
        CellStyle cellStyle = templateRow.getCell(colIndex).getCellStyle();
        CellStyle cloneStyle = row.getSheet().getWorkbook().createCellStyle();
        cloneStyle.cloneStyleFrom(cellStyle);
        return cloneStyle;
    }

    private Sheet createSheet(Workbook wb, Class<?> beanClass) {
        val excelSheet = beanClass.getAnnotation(ExcelSheet.class);
        String sheetName = null;
        if (excelSheet != null) {
            sheetName = excelSheet.name();
        }

        if (StringUtils.isBlank(sheetName)) {
            sheetName = beanClass.getSimpleName();
        }

        Sheet sheet = wb.createSheet(sheetName);

        return sheet;
    }
}
