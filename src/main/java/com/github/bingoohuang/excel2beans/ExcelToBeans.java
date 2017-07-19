package com.github.bingoohuang.excel2beans;

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
 * Mapping excel cell values to java beans.
 */
public class ExcelToBeans implements Closeable {
    private final Workbook workbook;
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
        int lastCellNum = getLastCellNum(rowRefs, sheet);
        if (lastCellNum <= 0) {
            return;
        }

        val redCellStyle = createRedCellStyle();

        for (val rowRef : rowRefs) {
            if (StringUtils.isEmpty(rowRef.error())) {
                continue;
            }

            val row = sheet.getRow(rowRef.getRowNum());
            val cell = row.createCell(lastCellNum);
            cell.setCellStyle(redCellStyle);
            cell.setCellValue(rowRef.error());
        }
    }

    public int getLastCellNum(List<? extends ExcelRowRef> rowRefs, Sheet sheet) {
        for (val rowRef : rowRefs) {
            val row = sheet.getRow(rowRef.getRowNum());
            return row.getLastCellNum();
        }

        return 0;
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
            if (StringUtils.isEmpty(rowRef.error())) {
                ExcelToBeansUtils.removeRow(sheet, rowRef.getRowNum() - count);
                ++count;
            }
        }
    }

    public void writeExcel(String name) {
        ExcelToBeansUtils.writeExcel(workbook, name);
    }

    @SneakyThrows
    public byte[] getWorkbookBytes() {
        val bout = new ByteArrayOutputStream();
        workbook.write(bout);
        bout.close();

        return bout.toByteArray();
    }

    @Override public void close() throws IOException {
        if (shouldBeClosedByMe) {
            workbook.close();
        }
    }
}
