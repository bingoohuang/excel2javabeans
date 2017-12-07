package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelSheet;
import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.experimental.var;
import lombok.val;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

public class ExcelToBeansUtils {
    @SneakyThrows
    public static Workbook getClassPathWorkbook(String classPathExcelName) {
        @Cleanup val is = getClassPathInputStream(classPathExcelName);
        return WorkbookFactory.create(is);
    }

    @SneakyThrows
    public static InputStream getClassPathInputStream(String classPathExcelName) {
        val classLoader = ExcelToBeansUtils.class.getClassLoader();
        return classLoader.getResourceAsStream(classPathExcelName);
    }


    public static void removeRow(Sheet sheet, int rowIndex) {
        val lastRowNum = sheet.getLastRowNum();
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
        } else if (rowIndex == lastRowNum) {
            val removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }

    @SneakyThrows
    public static byte[] getWorkbookBytes(Workbook workbook) {
        @Cleanup val bout = new ByteArrayOutputStream();
        workbook.write(bout);
        return bout.toByteArray();
    }


    public static void writeRedComments(Workbook workbook, Collection<CellData> cellDatas) {
        val factory = workbook.getCreationHelper();
        val globalNewCellStyle = reddenBorder(workbook.createCellStyle());

        // 重用cell style，提升性能
        val cellStyleMap = new HashMap<CellStyle, CellStyle>();

        for (val cellData : cellDatas) {
            val sheet = workbook.getSheetAt(cellData.getSheetIndex());
            val row = sheet.getRow(cellData.getRow());
            var cell = row.getCell(cellData.getCol());
            if (cell == null) cell = row.createCell(cellData.getCol());

            setCellStyle(globalNewCellStyle, cellStyleMap, cell);

            addComment(factory, cellData, cell);
        }
    }

    public static CellStyle reddenBorder(CellStyle cellStyle) {
        val borderStyle = BorderStyle.THIN;

        cellStyle.setBorderLeft(borderStyle);
        cellStyle.setBorderRight(borderStyle);
        cellStyle.setBorderTop(borderStyle);
        cellStyle.setBorderBottom(borderStyle);

        val redColorIndex = IndexedColors.RED.getIndex();

        cellStyle.setBottomBorderColor(redColorIndex);
        cellStyle.setTopBorderColor(redColorIndex);
        cellStyle.setLeftBorderColor(redColorIndex);
        cellStyle.setRightBorderColor(redColorIndex);

        return cellStyle;
    }

    private static void setCellStyle(CellStyle defaultCellStyle, Map<CellStyle, CellStyle> cellStyleMap, Cell cell) {
        val cellStyle = cell.getCellStyle();
        if (cellStyle == null) {
            cell.setCellStyle(defaultCellStyle);
            return;
        }

        var newCellStyle = cellStyleMap.get(cellStyle);
        if (newCellStyle == null) {
            newCellStyle = cell.getSheet().getWorkbook().createCellStyle();
            newCellStyle.cloneStyleFrom(cellStyle);
            cellStyleMap.put(cellStyle, reddenBorder(newCellStyle));
        }

        cell.setCellStyle(newCellStyle);
    }

    private static void addComment(CreationHelper factory, CellData cellData, Cell cell) {
        var comment = cell.getCellComment();
        if (comment == null) {
            val drawing = cell.getSheet().createDrawingPatriarch();
            // When the comment box is visible, have it show in a 1x3 space
            val anchor = factory.createClientAnchor();
            anchor.setCol1(cell.getColumnIndex());
            anchor.setCol2(cell.getColumnIndex() + 1);
            anchor.setRow1(cell.getRow().getRowNum());
            anchor.setRow2(cell.getRow().getRowNum() + 3);

            // Create the comment and set the text+author
            comment = drawing.createCellComment(anchor);

            cell.setCellComment(comment);
        }

        val str = factory.createRichTextString(cellData.getComment());
        comment.setString(str);

        val author = cellData.getCommentAuthor();
        if (StringUtils.isNotEmpty(author)) comment.setAuthor(author);
    }

    @SneakyThrows
    public static void writeExcel(Workbook workbook, String name) {
        @Cleanup val fileOut = new FileOutputStream(name);
        workbook.write(fileOut);
    }

    public static Sheet findSheet(Workbook workbook, Class<?> beanClass) {
        val excelSheet = beanClass.getAnnotation(ExcelSheet.class);
        if (excelSheet == null) return workbook.getSheetAt(0);

        for (int i = 0, ii = workbook.getNumberOfSheets(); i < ii; ++i) {
            val sheetName = workbook.getSheetName(i);
            if (sheetName.contains(excelSheet.name())) {
                return workbook.getSheetAt(i);
            }
        }

        throw new IllegalArgumentException("Unable to find sheet with name " + excelSheet.name());
    }
}
