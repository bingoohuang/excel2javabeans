package com.github.bingoohuang.excel2beans;

import lombok.experimental.UtilityClass;
import lombok.val;
import lombok.var;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;

import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

@UtilityClass
public class ExcelToBeansUtils {

    public void writeRedComments(Workbook workbook, Collection<CellData> cellDatas) {
        writeRedComments(workbook, cellDatas, 3, 5);
    }

    public void writeRedComments(Workbook workbook, Collection<CellData> cellDatas,
                                 int commentColSpan, int commentRowSpan) {
        removeAllComments(workbook);

        val globalNewCellStyle = reddenBorder(workbook.createCellStyle());
        val factory = workbook.getCreationHelper();

        // 重用cell style，提升性能
        val cellStyleMap = new HashMap<CellStyle, CellStyle>();

        for (val cellData : cellDatas) {
            val sheet = workbook.getSheetAt(cellData.getSheetIndex());
            val row = sheet.getRow(cellData.getRow());
            var cell = row.getCell(cellData.getCol());
            if (cell == null) cell = row.createCell(cellData.getCol());

            setCellStyle(globalNewCellStyle, cellStyleMap, cell);

            addComment(factory, cellData, cell, commentColSpan, commentRowSpan);
        }
    }

    public static void removeAllComments(Workbook workbook) {
        val cellStyle = workbook.createCellStyle();
        for (int i = 0, ii = workbook.getNumberOfSheets(); i < ii; ++i) {
            val sheet = workbook.getSheetAt(i);
            val comments = sheet.getCellComments();
            for (val entry : comments.entrySet()) {
                val comment = entry.getValue();
                val cell = sheet.getRow(comment.getRow()).getCell(comment.getColumn());
                cell.removeCellComment();
                cell.setCellStyle(cellStyle);
            }
        }
    }

    public CellStyle reddenBorder(CellStyle cellStyle) {
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

    private void setCellStyle(CellStyle defaultCellStyle, Map<CellStyle, CellStyle> cellStyleMap, Cell cell) {
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

    private void addComment(CreationHelper factory, CellData cellData, Cell cell, int colSpan, int rowSpan) {
        val col = cell.getColumnIndex();
        val row = cell.getRow().getRowNum();
        val drawing = cell.getSheet().createDrawingPatriarch();
        val anchor = drawing.createAnchor(0, 0, 0, 0,
                col, row, col + colSpan, row + rowSpan);
        val comment = drawing.createCellComment(anchor);
        comment.setString(factory.createRichTextString(cellData.getComment()));

        val author = cellData.getCommentAuthor();
        if (StringUtils.isNotEmpty(author)) comment.setAuthor(author);

        cell.setCellComment(comment);
    }
}
