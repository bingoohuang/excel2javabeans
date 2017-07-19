package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColIgnore;
import com.github.bingoohuang.excel2beans.annotations.ExcelColStyle;
import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.github.bingoohuang.excel2beans.annotations.ExcelSheet;
import com.google.common.io.ByteStreams;
import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.experimental.UtilityClass;
import lombok.experimental.var;
import lombok.val;
import org.apache.poi.ss.usermodel.*;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static com.github.bingoohuang.excel2beans.annotations.ExcelColAlign.*;
import static org.apache.commons.lang3.StringUtils.capitalize;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
@UtilityClass
public class ExcelToBeansUtils {
    @SneakyThrows
    public Workbook getClassPathWorkbook(String classPathExcelName) {
        val classLoader = ExcelToBeansUtils.class.getClassLoader();
        @Cleanup val is = classLoader.getResourceAsStream(classPathExcelName);
        return WorkbookFactory.create(is);
    }

    @SneakyThrows
    public InputStream getClassPathInputStream(String classPathExcelName) {
        val classLoader = ExcelToBeansUtils.class.getClassLoader();
        return classLoader.getResourceAsStream(classPathExcelName);
    }

    public List<ExcelBeanField> parseBeanFields(Class<?> beanClass, Sheet sheet) {
        val declaredFields = beanClass.getDeclaredFields();
        val fields = new ArrayList<ExcelBeanField>(declaredFields.length);

        for (val field : declaredFields) {
            processField(sheet, fields, field);
        }

        return fields;
    }

    public static void removeRow(Sheet sheet, int rowIndex) {
        int lastRowNum = sheet.getLastRowNum();
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
        } else if (rowIndex == lastRowNum) {
            val removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }

    private void processField(Sheet sheet, List<ExcelBeanField> fields, Field field) {
        if (Modifier.isStatic(field.getModifiers()) || Modifier.isTransient(field.getModifiers())) {
            return;
        }

        val rowIgnore = field.getAnnotation(ExcelColIgnore.class);
        if (rowIgnore != null) {
            return;
        }

        String fieldName = field.getName();
        if (fieldName.startsWith("$")) { // ignore un-normal fields like $jacocoData
            return;
        }

        val beanField = new ExcelBeanField();

        beanField.setColumnIndex(fields.size());
        beanField.setField(field);
        beanField.setName(fieldName);
        beanField.setSetter("set" + capitalize(fieldName));
        beanField.setGetter("get" + capitalize(fieldName));

        setTitle(field, beanField);
        setStyle(sheet, field, beanField);
        setIsCellData(field, beanField);

        fields.add(beanField);
    }

    private static void setIsCellData(Field field, ExcelBeanField beanField) {
        beanField.setCellDataType(field.getType() == CellData.class);
    }

    private void setStyle(Sheet sheet, Field field, ExcelBeanField beanField) {
        val colStyle = field.getAnnotation(ExcelColStyle.class);
        if (colStyle != null) {
            CellStyle style = setAlign(sheet, colStyle);
            if (style != null) {
                beanField.setCellStyle(style);
            }
        }
    }

    private void setTitle(Field field, ExcelBeanField beanField) {
        val colTitle = field.getAnnotation(ExcelColTitle.class);
        if (colTitle != null) {
            beanField.setTitle(colTitle.value());
        }
    }

    private CellStyle setAlign(Sheet sheet, ExcelColStyle colStyle) {
        var style = sheet.getWorkbook().createCellStyle();
        val align = colStyle.align();
        if (align == LEFT) {
            style.setAlignment(HorizontalAlignment.LEFT);
        } else if (align == CENTER) {
            style.setAlignment(HorizontalAlignment.CENTER);
        } else if (align == RIGHT) {
            style.setAlignment(HorizontalAlignment.RIGHT);
        } else {
            style = null;
        }

        return style;
    }

    @SneakyThrows
    public void download(HttpServletResponse response, Workbook workbook, String fileName) {
        @Cleanup val out = prepareDownload(response, fileName);
        workbook.write(out);
        workbook.close();
    }

    @SneakyThrows
    public void download(HttpServletResponse response, byte[] workbook, String fileName) {
        @Cleanup val out = prepareDownload(response, fileName);
        out.write(workbook);
    }

    @SneakyThrows
    public void download(HttpServletResponse response, InputStream workbook, String fileName) {
        @Cleanup val out = prepareDownload(response, fileName);
        ByteStreams.copy(workbook, out);
    }

    @SneakyThrows
    public static ServletOutputStream prepareDownload(HttpServletResponse response, String fileName) {
        response.setContentType("application/vnd.ms-excel;charset=UTF-8");
        val encodedFileName = URLEncoder.encode(fileName, "UTF-8");
        response.setHeader("Content-disposition", "attachment; " +
                "filename=\"" + encodedFileName + "\"; " +
                "filename*=utf-8'zh_cn'" + encodedFileName);
        return response.getOutputStream();
    }

    public static void writeRedComments(Workbook workbook, CellData... cellDatas) {
        val factory = workbook.getCreationHelper();

        val globalNewCellStyle = reddenBorder(workbook.createCellStyle());

        // 重用cell style，提升性能
        val cellStyleMap = new HashMap<CellStyle, CellStyle>();

        for (val cellData : cellDatas) {
            val sheet = workbook.getSheetAt(cellData.getSheetIndex());
            val row = sheet.getRow(cellData.getRow());
            val cell = row.getCell(cellData.getCol());

            setCellStyle(globalNewCellStyle, cellStyleMap, cell);

            addComment(factory, cellData, cell);
        }
    }

    public static CellStyle reddenBorder(CellStyle cellStyle) {
        val borderStyle = BorderStyle.MEDIUM;

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

    public static void setCellStyle(
            CellStyle globalRedBorderCellStyle, Map<CellStyle, CellStyle> cellStyleMap, Cell cell
    ) {
        val cellStyle = cell.getCellStyle();
        if (cellStyle == null) {
            cell.setCellStyle(globalRedBorderCellStyle);
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
        comment.setAuthor(cellData.getCommentAuthor());
    }

    @SneakyThrows
    public static void writeExcel(Workbook workbook, String name) {
        @Cleanup val fileOut = new FileOutputStream(name);
        workbook.write(fileOut);
    }

    public static Sheet findSheet(Workbook workbook, Class<?> beanClass) {
        val excelSheet = beanClass.getAnnotation(ExcelSheet.class);
        if (excelSheet == null) {
            return workbook.getSheetAt(0);
        }

        for (int i = 0, ii = workbook.getNumberOfSheets(); i < ii; ++i) {
            val sheetName = workbook.getSheetName(i);
            if (sheetName.contains(excelSheet.name())) {
                return workbook.getSheetAt(i);
            }
        }

        throw new IllegalArgumentException("Unable to find sheet with name " + excelSheet.name());
    }
}
