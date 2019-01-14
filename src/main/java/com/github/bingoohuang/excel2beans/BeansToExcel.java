package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelSheet;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import lombok.val;
import lombok.var;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.*;
import java.util.stream.IntStream;

public class BeansToExcel {
    private final Workbook workbook;
    private final Workbook styleTemplate;
    private final ReflectAsmCache reflectAsmCache = new ReflectAsmCache();
    private final Map<Class<?>, Set<String>> mapIncludedFields = Maps.newHashMap();

    public BeansToExcel() {
        this(null);
    }

    public void includes(Class<?> beanClass, String... includedFields) {
        mapIncludedFields.put(beanClass, new HashSet<>(Lists.newArrayList(includedFields)));
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

    private boolean isAllEmpty(List<?>... lists) {
        for (val list : lists) {
            if (!list.isEmpty()) return false;
        }
        return true;
    }

    public Workbook create(Map<String, Object> props, List<?>... lists) {
        if (isAllEmpty(lists)) {
            return workbook.createSheet("Empty").getWorkbook();
        }

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
            IntStream.rangeClosed(0, lastCellNum).forEach(sheet::autoSizeColumn);
        });
    }

    private void writeRowCells(Object bean, BeanClassBag bag, Row row) {
        IntStream.range(0, bag.getBeanFields().size()).forEach(i -> {
            val cell = row.createCell(i);
            val field = bag.getBeanField(i);

            val value = field.getFieldValue(bean);
            PoiUtil.writeCellValue(cell, value);

            val style = field.getCellStyle();
            if (style != null) cell.setCellStyle(style);
        });
    }

    protected BeanClassBag insureBagCreated(Map<String, Object> props, Map<Class, BeanClassBag> bagMap, Object bean) {
        val beanClass = bean.getClass();
        var bag = bagMap.get(beanClass);
        if (bag != null) return bag;

        bag = new BeanClassBag(beanClass);
        bagMap.put(beanClass, bag);

        bag.setSheet(createSheet(beanClass));
        Set<String> includedFields = mapIncludedFields.get(beanClass);
        bag.setBeanFields(new ExcelBeanFieldParser(beanClass, bag.getSheet())
                .parseBeanFields(includedFields, reflectAsmCache));

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

    private void cloneCellStyle(BeanClassBag bag, Row row, List<ExcelBeanField> fields) {
        if (styleTemplate == null) return;

        val templateSheet = parseTemplateSheet(bag);

        IntStream.range(0, fields.size()).forEach(colIndex -> {
            val excelBeanField = fields.get(colIndex);
            val titledColIndex = findTitledColIndex(colIndex, excelBeanField, templateSheet);

            val headStyle = cloneCellStyle(templateSheet, 0, titledColIndex);
            row.getCell(colIndex).setCellStyle(headStyle);
            row.setHeight(templateSheet.getRow(0).getHeight());

            val dataStyle = cloneCellStyle(templateSheet, 1, titledColIndex);
            excelBeanField.setCellStyle(dataStyle);
        });
    }

    private int findTitledColIndex(int colIndex, ExcelBeanField excelBeanField, Sheet templateSheet) {
        if (excelBeanField.getTitleColumnIndex() >= 0) return excelBeanField.getTitleColumnIndex();

        excelBeanField.setTitleColumnIndex(colIndex);
        if (!excelBeanField.hasTitle()) return colIndex;

        val row = templateSheet.getRow(0);
        val colCell = row.getCell(colIndex);
        if (PoiUtil.getCellStringValue(colCell).contains(excelBeanField.getTitle())) return colIndex;

        for (short cn = row.getFirstCellNum(), cellMax = row.getLastCellNum(); cn < cellMax; ++cn ) {
            val cell = row.getCell(cn);
            if (PoiUtil.getCellStringValue(cell).contains(excelBeanField.getTitle())) {
                excelBeanField.setTitleColumnIndex(cn);
                return cn;
            }
        }

        return colIndex;
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
