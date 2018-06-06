package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelSheet;
import com.google.common.collect.Maps;
import lombok.val;
import lombok.var;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.stream.IntStream;

public class BeansToExcel {
    private final Workbook workbook;
    private final Workbook styleTemplate;

    public BeansToExcel() {
        this(null);
    }

    /**
     * Construct BeansToExcel with a style template.
     *
     * @param styleTemplate style template
     */
    public BeansToExcel(Workbook styleTemplate) {
        this.styleTemplate = styleTemplate;
        this.workbook = new XSSFWorkbook();
    }

    public Workbook create(List<?>... lists) {
        return create(Maps.newHashMap(), lists);
    }

    public Workbook create(Map<String, Object> props, List<?>... lists) {
        Map<Class, BeanClassBag> beanClassBag = Maps.newHashMap();
        Arrays.stream(lists).forEach(x -> x.forEach(y -> writeBeanToExcel(props, beanClassBag, y)));
        autoSizeColumn(beanClassBag);

        return workbook;
    }

    private void writeBeanToExcel(Map<String, Object> props, Map<Class, BeanClassBag> bagMap, Object bean) {
        val bag = insureBagCreated(props, bagMap, bean);
        writeRowCells(bean, bag, createRow(bag));
    }

    private Row createRow(BeanClassBag bag) {
        val sheet = bag.getSheet();
        if (bag.isFirstRowCreated()) {
            return sheet.createRow(sheet.getLastRowNum() + 1);
        }

        bag.setFirstRowCreated(true);
        return sheet.createRow(0);
    }

    private void autoSizeColumn(Map<Class, BeanClassBag> sheets) {
        sheets.values().forEach(bag -> {
            val sheet = bag.getSheet();
            val lastCellNum = sheet.getRow(sheet.getLastRowNum()).getLastCellNum();
            IntStream.rangeClosed(0, lastCellNum).forEach(i -> sheet.autoSizeColumn(i));
        });
    }

    private void writeRowCells(Object bean, BeanClassBag bag, Row row) {
        IntStream.range(0, bag.getBeanFields().size()).forEach(i -> {
            val cell = row.createCell(i);
            val field = bag.getBeanField(i);

            val fieldValue = field.getFieldValue(bean);
            cell.setCellValue(String.valueOf(fieldValue));
            val cellStyle = field.getCellStyle();
            if (cellStyle != null) cell.setCellStyle(cellStyle);
        });
    }

    private BeanClassBag insureBagCreated(
            Map<String, Object> props, Map<Class, BeanClassBag> beanClassBag, Object bean) {
        val beanClass = bean.getClass();
        var bag = beanClassBag.get(beanClass);
        if (bag != null) return bag;

        bag = new BeanClassBag(beanClass);
        beanClassBag.put(beanClass, bag);

        bag.setSheet(createSheet(beanClass));
        bag.setBeanFields(new ExcelBeanFieldParser(beanClass, bag.getSheet()).parseBeanFields());

        addHeadToSheet(props, bag);
        addTitleToSheet(bag);

        return bag;
    }

    private void addHeadToSheet(Map<String, Object> props, BeanClassBag bag) {
        val excelSheet = bag.getBeanClass().getAnnotation(ExcelSheet.class);
        if (excelSheet == null) return;

        val headKey = excelSheet.headKey();
        if (!props.containsKey(headKey)) return;

        val head = String.valueOf(props.get(headKey));
        if (StringUtils.isEmpty(head)) return;

        val lastCol = bag.getBeanFields().size() - 1;
        val region = new CellRangeAddress(0, 0, 0, lastCol);
        bag.getSheet().addMergedRegion(region);

        val row = createRow(bag);
        row.createCell(0).setCellValue(head);
    }

    private void addTitleToSheet(BeanClassBag bag) {
        val row = createRow(bag);
        val beanFields = bag.getBeanFields();
        IntStream.range(0, beanFields.size()).forEach(i ->
                row.createCell(i).setCellValue(beanFields.get(i).getTitle()));

        cloneCellStyle(bag, row, beanFields);
    }

    private void cloneCellStyle(BeanClassBag bag, Row row, List<ExcelBeanField> beanFields) {
        if (styleTemplate == null) return;

        val templateSheet = parseTemplateSheet(bag);
        IntStream.range(0, beanFields.size()).forEach(colIndex -> {
            val headStyle = cloneCellStyle(templateSheet, 0, colIndex);
            row.getCell(colIndex).setCellStyle(headStyle);

            row.setHeight(templateSheet.getRow(0).getHeight());

            val dataStyle = cloneCellStyle(templateSheet, 1, colIndex);
            beanFields.get(colIndex).setCellStyle(dataStyle);
        });
    }

    private Sheet parseTemplateSheet(BeanClassBag bag) {
        val sheetName = bag.getSheet().getSheetName();
        var sheet = styleTemplate.getSheet(sheetName);

        return sheet != null ? sheet : styleTemplate.getSheetAt(0);
    }

    private CellStyle cloneCellStyle(Sheet styleSheet, int rowIndex, int colIndex) {
        val styleRow = styleSheet.getRow(rowIndex);
        val cellStyle = styleRow.getCell(colIndex).getCellStyle();
        val cloneStyle = workbook.createCellStyle();
        cloneStyle.cloneStyleFrom(cellStyle);
        return cloneStyle;
    }

    private Sheet createSheet(Class<?> beanClass) {
        val excelSheet = beanClass.getAnnotation(ExcelSheet.class);
        val sheetName = excelSheet != null && StringUtils.isNotBlank(excelSheet.name())
                ? excelSheet.name() : beanClass.getSimpleName();

        return workbook.createSheet(sheetName);
    }

}
