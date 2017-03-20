package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelSheet;
import com.google.common.collect.Maps;
import lombok.experimental.var;
import lombok.val;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;
import java.util.Map;

/**
 * Created by bingoohuang on 2017/3/20.
 */
public class BeansToExcel {
    private final XSSFWorkbook workbook;
    private final Workbook templateWorkbook;

    public BeansToExcel() {
        this(null);
    }

    public BeansToExcel(Workbook templateWorkbook) {
        this.templateWorkbook = templateWorkbook;
        this.workbook = new XSSFWorkbook();
    }

    public Workbook create(List<?>... lists) {
        Map<Class, BeanClassBag> beanClassBag = Maps.newHashMap();

        for (val list : lists) {
            for (val bean : list) {
                writeBeanToExcel(beanClassBag, bean);
            }
        }

        autoSizeColumn(beanClassBag);

        return workbook;
    }

    private void writeBeanToExcel(Map<Class, BeanClassBag> beanClassBag, Object bean) {
        val bag = insureBagCreated(beanClassBag, bean);

        val sheet = bag.getSheet();
        val row = sheet.createRow(sheet.getLastRowNum() + 1);
        writeRowCells(bean, bag, row);
    }

    private void autoSizeColumn(Map<Class, BeanClassBag> sheets) {
        for (val bag : sheets.values()) {
            val lastCellNum = bag.getSheet().getRow(0).getLastCellNum();
            for (int i = 0; i <= lastCellNum; ++i) {
                bag.getSheet().autoSizeColumn(i); // adjust width of the column
            }
        }
    }

    private void writeRowCells(Object bean, BeanClassBag bag, Row row) {
        for (int i = 0, ii = bag.getBeanFields().length; i < ii; ++i) {
            val cell = row.createCell(i);
            val field = bag.getBeanFields()[i];

            val fieldValue = field.getFieldValue(bag.getFieldAccess(), bag.getMethodAccess(), bean);
            cell.setCellValue(String.valueOf(fieldValue));
            cell.setCellStyle(field.getCellStyle());
        }
    }

    private BeanClassBag insureBagCreated(Map<Class, BeanClassBag> beanClassBag, Object bean) {
        val beanClass = bean.getClass();
        var bag = beanClassBag.get(beanClass);
        if (bag != null) return bag;

        bag = new BeanClassBag(beanClass);
        beanClassBag.put(beanClass, bag);

        bag.setSheet(createSheet(beanClass));
        bag.setBeanFields(ExcelToBeansUtils.parseBeanFields(beanClass));

        addTitleToSheet(bag);

        return bag;
    }

    private void addTitleToSheet(BeanClassBag bag) {
        val row = bag.getSheet().createRow(0);
        val beanFields = bag.getBeanFields();
        for (int i = 0, ii = beanFields.length; i < ii; ++i) {
            row.createCell(i).setCellValue(beanFields[i].getTitle());
        }

        cloneCellStyle(bag, row, beanFields);
    }

    private void cloneCellStyle(BeanClassBag bag, Row row, ExcelBeanField[] beanFields) {
        if (templateWorkbook == null) return;

        val templateSheet = parseTemplateSheet(bag);
        for (int colIndex = 0, ii = beanFields.length; colIndex < ii; ++colIndex) {
            val headStyle = cloneCellStyle(templateSheet, 0, colIndex);
            row.getCell(colIndex).setCellStyle(headStyle);

            val dataStyle = cloneCellStyle(templateSheet, 1, colIndex);
            beanFields[colIndex].setCellStyle(dataStyle);
        }
    }

    private Sheet parseTemplateSheet(BeanClassBag bag) {
        val sheetName = bag.getSheet().getSheetName();
        var templateSheet = templateWorkbook.getSheet(sheetName);

        return templateSheet != null ? templateSheet : templateWorkbook.getSheetAt(0);
    }

    private CellStyle cloneCellStyle(Sheet templateSheet, int rowIndex, int colIndex) {
        val templateRow = templateSheet.getRow(rowIndex);
        val cellStyle = templateRow.getCell(colIndex).getCellStyle();
        val cloneStyle = workbook.createCellStyle();
        cloneStyle.cloneStyleFrom(cellStyle);
        return cloneStyle;
    }

    private Sheet createSheet(Class<?> beanClass) {
        val excelSheet = beanClass.getAnnotation(ExcelSheet.class);
        val sheetName = parseSheetName(beanClass, excelSheet);

        return workbook.createSheet(sheetName);
    }

    private String parseSheetName(Class<?> beanClass, ExcelSheet excelSheet) {
        if (excelSheet == null) {
            return beanClass.getSimpleName();
        }

        if (StringUtils.isNotBlank(excelSheet.name())) {
            return excelSheet.name();
        }

        return beanClass.getSimpleName();
    }
}
