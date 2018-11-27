package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColIgnore;
import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.github.bingoohuang.excel2beans.annotations.ExcelTemplateSheet;
import com.github.bingoohuang.utils.reflect.Fields;
import com.github.bingoohuang.utils.type.Generic;
import com.google.common.collect.Maps;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import lombok.val;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

@RequiredArgsConstructor @Slf4j
public class BeansToExcelOnTitle {
    // 模板表单
    private final Sheet sheet;
    // 模板行索引
    private int templateRowNum = 0;
    // 标题行
    private Row titleRow;
    // 模板行
    private Row tmplRow;

    public Workbook create(List<?> records) {
        Map<String, Integer> titledColMap = Maps.newHashMap();
        parseTitledMap(records, titledColMap);

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
        if (colTitle != null && StringUtils.isNotEmpty(colTitle.value())) {
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

    private void parseTitledMap(List<?> records, Map<String, Integer> titledColMap) {
        if (records.isEmpty()) return;

        val recordClass = records.get(0).getClass();
        val excelTemplateSheet = recordClass.getAnnotation(ExcelTemplateSheet.class);

        titleRow = sheet.getRow(excelTemplateSheet.titleRowRef() - 1);
        templateRowNum = excelTemplateSheet.templateRowRef() - 1;
        tmplRow = sheet.getRow(templateRowNum);

        PoiUtil.shiftRows(sheet, records, tmplRow.getRowNum());

        for (int i = titleRow.getFirstCellNum(), ii = titleRow.getLastCellNum(); i < ii; ++i) {
            val title = findTitle(i);

            PoiUtil.blankCell(tmplRow, i);

            if (StringUtils.isEmpty(title)) continue;

            if (titledColMap.containsKey(title)) {
                log.warn("duplicate title {} found at [{}]", title, new CellReference(titleRow.getCell(i)).formatAsString());
                continue;
            }

            titledColMap.put(title, i);
        }
    }

    private String findTitle(int col) {
        val t = PoiUtil.getCellValue(tmplRow, col);
        return StringUtils.isNotEmpty(t) ? t : PoiUtil.getCellValue(titleRow, col);
    }

}
