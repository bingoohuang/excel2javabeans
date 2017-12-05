package com.github.bingoohuang.excel2beans;

import lombok.Cleanup;
import lombok.Getter;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;

import java.io.ByteArrayOutputStream;
import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/**
 * Mapping excel rows to java beans.
 */
public class ExcelToBeans implements Closeable {
    private @Getter final Workbook workbook;
    private final boolean shouldBeClosedByMe;

    @SneakyThrows
    public ExcelToBeans(InputStream excelInputStream) {
        this.workbook = WorkbookFactory.create(excelInputStream);
        this.shouldBeClosedByMe = true;
    }

    @SneakyThrows
    public ExcelToBeans(Workbook workbook) {
        this.workbook = workbook;
        this.shouldBeClosedByMe = false;
    }

    @SneakyThrows
    public <T> List<T> convert(Class<T> beanClass) {
        val converter = new ExcelSheetToBeans(workbook, beanClass);
        return converter.convert();
    }

    public void writeError(Class<?> beanClass, List<? extends ExcelRowRef> rowRefs) {
        val sheet = ExcelToBeansUtils.findSheet(workbook, beanClass);
        val converter = new ExcelSheetToBeans(workbook, beanClass);
        int lastCellNum = getLastCellNum(converter, rowRefs, sheet);
        if (lastCellNum <= 0) return;

        val redCellStyle = createRedCellStyle();

        for (val rowRef : rowRefs) {
            if (StringUtils.isEmpty(rowRef.error())) continue;

            val row = sheet.getRow(rowRef.getRowNum());
            val cell = row.createCell(lastCellNum);
            cell.setCellStyle(redCellStyle);
            cell.setCellValue(rowRef.error());
        }
    }

    public int getLastCellNum(ExcelSheetToBeans converter, List<? extends ExcelRowRef> rowRefs, Sheet sheet) {
        if (converter.isHasTitle()) {
            return sheet.getRow(converter.findTitleRowNum()).getLastCellNum();
        }

        return rowRefs.isEmpty() ? 0 : sheet.getRow(rowRefs.get(0).getRowNum()).getLastCellNum();
    }

    private CellStyle createRedCellStyle() {
        val cellStyle = workbook.createCellStyle();
        val font = workbook.createFont();
        font.setColor(IndexedColors.RED.getIndex());
        cellStyle.setFont(font);

        return cellStyle;
    }

    public void removeOkRows(Class<?> beanClass, List<? extends ExcelRowRef> rowRefs) {
        val sheet = ExcelToBeansUtils.findSheet(workbook, beanClass);
        int count = 0;
        for (val rowRef : rowRefs) {
            if (StringUtils.isNotEmpty(rowRef.error())) continue;

            ExcelToBeansUtils.removeRow(sheet, rowRef.getRowNum() - count);
            ++count;
        }
    }

    public void writeExcel(String name) {
        ExcelToBeansUtils.writeExcel(workbook, name);
    }

    @SneakyThrows
    public byte[] getWorkbookBytes() {
        @Cleanup val bout = new ByteArrayOutputStream();
        workbook.write(bout);
        return bout.toByteArray();
    }

    @Override public void close() throws IOException {
        if (shouldBeClosedByMe) workbook.close();
    }
}
