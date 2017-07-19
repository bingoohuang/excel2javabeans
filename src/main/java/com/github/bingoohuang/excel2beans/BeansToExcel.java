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
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;
import java.util.Map;

/**
 * Created by bingoohuang on 2017/3/20.
 */
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
        Map<String, Object> props = Maps.newHashMap();
        return create(props, lists);
    }

    public Workbook create(Map<String, Object> props, List<?>... lists) {
        Map<Class, BeanClassBag> beanClassBag = Maps.newHashMap();

        for (val list : lists) {
            for (val bean : list) {
                writeBeanToExcel(props, beanClassBag, bean);
            }
        }

        autoSizeColumn(beanClassBag);

        return workbook;
    }

    private void writeBeanToExcel(
            Map<String, Object> props, Map<Class, BeanClassBag> beanClassBagMap, Object bean) {
        val bag = insureBagCreated(props, beanClassBagMap, bean);

        val row = createRow(bag);
        writeRowCells(bean, bag, row);
    }

    private Row createRow(BeanClassBag bag) {
        Sheet sheet = bag.getSheet();
        if (bag.isFirstRowCreated()) {
            return sheet.createRow(sheet.getLastRowNum() + 1);
        } else {
            bag.setFirstRowCreated(true);
            return sheet.createRow(0);
        }
    }

    private void autoSizeColumn(Map<Class, BeanClassBag> sheets) {
        for (val bag : sheets.values()) {
            Sheet sheet = bag.getSheet();
            val lastCellNum = sheet.getRow(sheet.getLastRowNum()).getLastCellNum();
            for (int i = 0; i <= lastCellNum; ++i) {
                sheet.autoSizeColumn(i); // adjust width of the column
            }
        }
    }

    private void writeRowCells(Object bean, BeanClassBag bag, Row row) {
        for (int i = 0, ii = bag.getBeanFields().size(); i < ii; ++i) {
            val cell = row.createCell(i);
            val field = bag.getBeanField(i);

            val fieldValue = field.getFieldValue(bag.getFieldAccess(), bag.getMethodAccess(), bean);
            cell.setCellValue(String.valueOf(fieldValue));
            cell.setCellStyle(field.getCellStyle());
        }
    }

    private BeanClassBag insureBagCreated(
            Map<String, Object> props, Map<Class, BeanClassBag> beanClassBag, Object bean) {
        val beanClass = bean.getClass();
        var bag = beanClassBag.get(beanClass);
        if (bag != null) return bag;

        bag = new BeanClassBag(beanClass);
        beanClassBag.put(beanClass, bag);

        bag.setSheet(createSheet(beanClass));
        bag.setBeanFields(ExcelToBeansUtils.parseBeanFields(beanClass, bag.getSheet()));

        addHeadToSheet(props, bag);
        addTitleToSheet(bag);

        return bag;
    }

    private void addHeadToSheet(Map<String, Object> props, BeanClassBag bag) {
        val excelSheet = bag.getBeanClass().getAnnotation(ExcelSheet.class);
        if (excelSheet == null) {
            return;
        }

        String headKey = excelSheet.headKey();
        if (!props.containsKey(headKey)) {
            return;
        }

        String head = String.valueOf(props.get(headKey));
        if (StringUtils.isEmpty(head)) {
            return;
        }

        val cra = new CellRangeAddress(0, 0,
                0, bag.getBeanFields().size() - 1);
        bag.getSheet().addMergedRegion(cra);

        val row = createRow(bag);
        row.createCell(0).setCellValue(head);
    }

    private void addTitleToSheet(BeanClassBag bag) {
        val row = createRow(bag);
        val beanFields = bag.getBeanFields();
        for (int i = 0, ii = beanFields.size(); i < ii; ++i) {
            row.createCell(i).setCellValue(beanFields.get(i).getTitle());
        }

        cloneCellStyle(bag, row, beanFields);
    }

    private void cloneCellStyle(BeanClassBag bag, Row row, List<ExcelBeanField> beanFields) {
        if (styleTemplate == null) return;

        val templateSheet = parseTemplateSheet(bag);
        for (int colIndex = 0, ii = beanFields.size(); colIndex < ii; ++colIndex) {
            val headStyle = cloneCellStyle(templateSheet, 0, colIndex);
            row.getCell(colIndex).setCellStyle(headStyle);

            val dataStyle = cloneCellStyle(templateSheet, 1, colIndex);
            beanFields.get(colIndex).setCellStyle(dataStyle);
        }
    }

    private Sheet parseTemplateSheet(BeanClassBag bag) {
        val sheetName = bag.getSheet().getSheetName();
        var styleSheet = styleTemplate.getSheet(sheetName);
        if (styleSheet == null) {
            styleSheet = styleTemplate.getSheetAt(0);
        }

        return styleSheet;
    }

    private CellStyle cloneCellStyle(Sheet styleSheet, int rowIndex, int colIndex) {
        val templateRow = styleSheet.getRow(rowIndex);
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
