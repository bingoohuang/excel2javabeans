package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColIgnore;
import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.github.bingoohuang.excel2beans.annotations.ExcelTemplateSheet;
import com.github.bingoohuang.utils.reflect.Fields;
import com.github.bingoohuang.utils.type.Generic;
import com.google.common.collect.HashMultimap;
import com.google.common.collect.Maps;
import com.google.common.collect.Multimap;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import lombok.val;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import static org.apache.commons.lang3.StringUtils.isNotEmpty;

@RequiredArgsConstructor @Slf4j
public class BeansToExcelOnTitle {
    // 模板表单
    private final Sheet sheet;
    // Javabean class
    private final Class<?> javabeanClass;
    // 模板行索引
    private int templateRowNum = 0;
    // 标题行
    private Row titleRow;
    // 模板行
    private Row tmplRow;

    public BeansToExcelOnTitle(String classpathTemplateExcelFileName, Class<?> javabeanClass) {
        this(findTemplateSheet(classpathTemplateExcelFileName, javabeanClass), javabeanClass);
    }

    private static Sheet findTemplateSheet(String classpathTemplateExcelFileName, Class<?> javabeanClass) {
        val excelTemplateSheet = javabeanClass.getAnnotation(ExcelTemplateSheet.class);
        val sheetName = excelTemplateSheet.templateSheetName();
        val workbook = PoiUtil.getClassPathWorkbook(classpathTemplateExcelFileName);
        return isNotEmpty(sheetName) ? workbook.getSheet(sheetName) : workbook.getSheetAt(0);
    }

    public Workbook create(List<?> records) {
        val titledColMap = parseTitledMap(records);

        for (int i = 0, ii = records.size(); i < ii; ++i) {
            val row = createRow(i);
            val record = records.get(i);

            writeRecordToRow(record, row, titledColMap);
        }

        PoiUtil.removeOtherSheets(sheet);
        return sheet.getWorkbook();
    }

    @SuppressWarnings("unchecked")
    private void writeRecordToRow(Object record, Row row, Map<String, Integer> titledColMap) {
        for (val field : record.getClass().getDeclaredFields()) {
            if (Fields.shouldIgnored(field, ExcelColIgnore.class)) continue;

            writeFieldValue(record, row, titledColMap, field);
        }
    }

    @SuppressWarnings("unchecked")
    private void writeFieldValue(Object record, Row row, Map<String, Integer> titledColMap, Field field) {
        val colTitle = field.getAnnotation(ExcelColTitle.class);
        if (colTitle != null && isNotEmpty(colTitle.value())) {
            val col = titledColMap.get(colTitle.value());
            if (col == null) {
                log.warn("@ExcelColTitle({}) for {} does not exists in template excel sheet",
                        colTitle.value(), field.getName());
                return;
            }

            val fv = Fields.invokeField(field, record);
            val cell = row.getCell(col);
            PoiUtil.writeCellValue(cell, fv);
        } else if (Generic.of(field.getGenericType()).isRawType(Map.class)) {
            val fv = Fields.invokeField(field, record);
            if (fv == null) return;

            val map = (Map<String, String>) fv;
            for (val entry : map.entrySet()) {
                val col = titledColMap.get(entry.getKey());
                if (col == null) {
                    log.warn("Map key title {} for {} does not exists in template excel sheet",
                            entry.getKey(), field.getName());
                    continue;
                }

                val cell = row.getCell(col);
                PoiUtil.writeCellValue(cell, entry.getValue());
            }
        }
    }

    private Row createRow(int offset) {
        if (offset == 0) return tmplRow; // 模板行,不需要新建，直接跳过

        val row = sheet.createRow(templateRowNum + offset);
        for (int j = titleRow.getFirstCellNum(), jj = titleRow.getLastCellNum(); j < jj; ++j) {
            val cell = row.createCell(j);

            cell.setCellStyle(tmplRow.getCell(j).getCellStyle());
        }

        return row;
    }

    private Map<String, Integer> parseTitledMap(List<?> records) {
        val excelTemplateSheet = javabeanClass.getAnnotation(ExcelTemplateSheet.class);

        titleRow = sheet.getRow(excelTemplateSheet.titleRowRef() - 1);
        templateRowNum = excelTemplateSheet.templateRowRef() - 1;
        tmplRow = sheet.getRow(templateRowNum);

        PoiUtil.shiftRows(sheet, records, tmplRow.getRowNum());

        Multimap<String, Integer> duplicateTitles = HashMultimap.create();

        for (int i = titleRow.getFirstCellNum(), ii = titleRow.getLastCellNum(); i < ii; ++i) {
            val title = findTitle(i);
            PoiUtil.blankCell(tmplRow, i);

            if (StringUtils.isEmpty(title)) continue;

            duplicateTitles.put(title, i);
        }

        return logDuplicateTitles(duplicateTitles);
    }

    private Map<String, Integer> logDuplicateTitles(Multimap<String, Integer> duplicateTitles) {
        Map<String, Integer> map = Maps.newHashMap();
        for (val title : duplicateTitles.keySet()) {
            val columns = duplicateTitles.get(title);
            map.put(title, columns.iterator().next());
            if (columns.size() > 1) {
                log.warn("duplicate titles {} found at {}", title, columns.stream()
                        .map(x -> new CellReference(titleRow.getCell(x)).formatAsString())
                        .collect(Collectors.joining(", ")));
            }
        }

        return map;
    }

    private String findTitle(int col) {
        val t = PoiUtil.getCellValue(tmplRow, col);
        return isNotEmpty(t) ? t : PoiUtil.getCellValue(titleRow, col);
    }

}
