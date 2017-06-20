package com.github.bingoohuang.excel2maps;

import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.Closeable;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

public class ExcelToMaps implements Closeable {
    private final Workbook workbook;
    private final boolean shouldBeClosedByMe;

    @SneakyThrows
    public ExcelToMaps(InputStream excelInputStream) {
        this.workbook = WorkbookFactory.create(excelInputStream);
        this.shouldBeClosedByMe = true;
    }

    @SneakyThrows
    public ExcelToMaps(Workbook workbook) {
        this.workbook = workbook;
        this.shouldBeClosedByMe = false;
    }


    @SneakyThrows
    public List<Map<String, String>> convert(ExcelToMapsConfig excelToMapsConfig, int sheetIndex) {
        val converter = new ExcelSheetToMaps(workbook, excelToMapsConfig);
        return converter.convert(sheetIndex);
    }

    @Override public void close() throws IOException {
        if (shouldBeClosedByMe) {
            workbook.close();
        }
    }
}
