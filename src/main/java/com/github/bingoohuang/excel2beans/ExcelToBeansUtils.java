package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.*;
import com.google.common.base.Optional;
import com.google.common.collect.HashBasedTable;
import com.google.common.collect.Maps;
import com.google.common.collect.Table;
import com.google.common.io.ByteStreams;
import com.google.common.primitives.Primitives;
import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.experimental.UtilityClass;
import lombok.experimental.var;
import lombok.val;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.ImageUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.lang.reflect.ParameterizedType;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.apache.commons.lang3.StringUtils.capitalize;

@UtilityClass
public class ExcelToBeansUtils {
    public static int computeAxisRowIndex(Sheet sheet, Picture picture) {
        // Calculates the dimensions in EMUs for the anchor of the given picture
        val dimension = ImageUtils.getDimensionFromAnchor(picture);
        val halfHeight = dimension.getHeight() / Units.EMU_PER_POINT / 2;

        val clientAnchor = picture.getClientAnchor();
        val anchorRow1 = clientAnchor.getRow1();
        val fromRowHeight = sheet.getRow(anchorRow1).getHeightInPoints();
        val anchorDy1 = clientAnchor.getDy1();
        val anchorRow2 = clientAnchor.getRow2();
        val y1 = sheet instanceof HSSFSheet
                ? anchorDy1 / 256.0f * fromRowHeight // refer to HSSFClientAnchor.getAnchorHeightInPoints
                : anchorDy1 / Units.EMU_PER_POINT;

        var sumHeight = fromRowHeight - y1;
        if (sumHeight >= halfHeight) return anchorRow1;

        for (var i = anchorRow1 + 1; i < anchorRow2; ++i) {
            sumHeight += sheet.getRow(i).getHeightInPoints();
            if (sumHeight >= halfHeight) return i;
        }

        return anchorRow2;
    }

    public static int computeAxisColIndex(Sheet sheet, Picture picture) {
        // Calculates the dimensions in EMUs for the anchor of the given picture
        val dimension = ImageUtils.getDimensionFromAnchor(picture); //
        val halfWidth = dimension.getHeight() / Units.EMU_PER_PIXEL / 2;

        val clientAnchor = picture.getClientAnchor();
        val anchorCol1 = clientAnchor.getCol1();
        val anchorCol2 = clientAnchor.getCol2();
        val anchorDx1 = clientAnchor.getDx1();

        val fromColumnWidth = sheet.getColumnWidthInPixels(anchorCol1);
        var sumWidth = fromColumnWidth - anchorDx1 / Units.EMU_PER_PIXEL;
        if (sumWidth >= halfWidth) return anchorCol1;

        for (var i = anchorCol1 + 1; i < anchorCol2; ++i) {
            sumWidth += sheet.getColumnWidthInPixels(i);
            if (sumWidth >= halfWidth) return i;
        }

        return anchorCol2;
    }


    public static Table<Integer, Integer, ImageData> readAllCellImages(Sheet sheet) {
        val patriarch = sheet.getDrawingPatriarch();
        if (patriarch instanceof XSSFDrawing) {
            return readAllCellImages((XSSFDrawing) patriarch, sheet);
        } else if (patriarch instanceof HSSFPatriarch) {
            return readAllCellImages((HSSFPatriarch) patriarch, sheet);
        }

        return HashBasedTable.create();
    }

    private static Table<Integer, Integer, ImageData> readAllCellImages(HSSFPatriarch patriarch, Sheet sheet) {
        val images = HashBasedTable.<Integer, Integer, ImageData>create();
        val allPictures = sheet.getWorkbook().getAllPictures();
        for (val shape : patriarch.getChildren()) {
            if (!(shape instanceof HSSFPicture && shape.getAnchor() instanceof HSSFClientAnchor)) continue;

            val picture = (HSSFPicture) shape;
            val imageData = createImageData(allPictures.get(picture.getPictureIndex() - 1));

            val axisRow = computeAxisRowIndex(sheet, picture);
            val axisCol = computeAxisColIndex(sheet, picture);

            images.put(axisRow, axisCol, imageData);
        }

        return images;
    }

    private static Table<Integer, Integer, ImageData> readAllCellImages(XSSFDrawing drawing, Sheet sheet) {
        val images = HashBasedTable.<Integer, Integer, ImageData>create();
        for (val shape : drawing.getShapes()) {
            if (!(shape instanceof XSSFPicture)) continue;

            val picture = (XSSFPicture) shape;
            val imageData = createImageData(picture.getPictureData());

            val axisRow = computeAxisRowIndex(sheet, picture);
            val axisCol = computeAxisColIndex(sheet, picture);

            images.put(axisRow, axisCol, imageData);
        }

        return images;
    }

    public static ImageData createImageData(PictureData pic) {
        return new ImageData(pic.getData(), pic.suggestFileExtension(), pic.getMimeType(), pic.getPictureType());
    }

    @SneakyThrows
    public Workbook getClassPathWorkbook(String classPathExcelName) {
        @Cleanup val is = getClassPathInputStream(classPathExcelName);
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

    private void processField(Sheet sheet, List<ExcelBeanField> fields, Field field) {
        if (Modifier.isStatic(field.getModifiers())) return;
        if (Modifier.isTransient(field.getModifiers())) return;

        val rowIgnore = field.getAnnotation(ExcelColIgnore.class);
        if (rowIgnore != null) return;

        val fieldName = field.getName();
        if (fieldName.startsWith("$")) return; // ignore un-normal fields like $jacocoData

        val beanField = new ExcelBeanField();

        beanField.setColumnIndex(fields.size());
        beanField.setField(field);
        beanField.setFieldName(fieldName);
        beanField.setSetter("set" + capitalize(fieldName));
        beanField.setGetter("get" + capitalize(fieldName));

        setTitle(field, beanField);
        setStyle(sheet, field, beanField);
        setIsCellData(field, beanField);
        setMultipleColumns(field, beanField);
        setValueOfMethod(beanField);

        fields.add(beanField);
    }

    private static void setValueOfMethod(ExcelBeanField beanField) {
        val fieldType = beanField.getFieldType();
        if (fieldType == String.class) return;

        val valueOfMethod = getValueOfMethodFrom(fieldType);
        beanField.setValueOfMethod(valueOfMethod);
    }

    public static Method getValueOfMethodFrom(Class targetClazz) {
        val existsMethod = valueOfMethodCache.get(targetClazz);
        if (existsMethod != null) return existsMethod.orNull();

        val clazz = Primitives.wrap(targetClazz);
        try {
            val valueOfMethod = clazz.getMethod("valueOf", new Class<?>[]{String.class});
            if (Modifier.isStatic(valueOfMethod.getModifiers())
                    && Modifier.isPublic(valueOfMethod.getModifiers())
                    && valueOfMethod.getReturnType().isAssignableFrom(clazz)) {
                valueOfMethodCache.put(targetClazz, Optional.of(valueOfMethod));

                return valueOfMethod;
            }

        } catch (Exception e) {
            valueOfMethodCache.put(clazz, Optional.<Method>absent());
        }

        return null;
    }

    public static final Map<Class, Optional<Method>> valueOfMethodCache = Maps.newConcurrentMap();

    public static Object invokeValueOf(Class clazz, String value) {
        val valueOfMethod = valueOfMethodCache.get(clazz);
        if (valueOfMethod == null || !valueOfMethod.isPresent()) return null;

        try {
            return valueOfMethod.get().invoke(null, value);
        } catch (Exception e) {
            // ignore
        }

        return null;
    }

    private static void setMultipleColumns(Field field, ExcelBeanField beanField) {
        val genericType = field.getGenericType();
        val isCollectionGeneric = genericType instanceof ParameterizedType
                && List.class.isAssignableFrom(field.getType());
        if (!isCollectionGeneric) return;

        val parameterizedType = (ParameterizedType) genericType;
        val actualTypeArgs = parameterizedType.getActualTypeArguments();
        if (actualTypeArgs.length == 1) {
            beanField.setMultipleColumns(true);
            beanField.setElementType((Class) actualTypeArgs[0]);
        }
    }

    private static void setIsCellData(Field field, ExcelBeanField beanField) {
        beanField.setCellDataType(field.getType() == CellData.class);
    }

    private void setStyle(Sheet sheet, Field field, ExcelBeanField beanField) {
        val colStyle = field.getAnnotation(ExcelColStyle.class);
        if (colStyle == null) return;

        val style = setAlign(sheet, colStyle);
        if (style == null) return;

        beanField.setCellStyle(style);
    }

    private void setTitle(Field field, ExcelBeanField beanField) {
        val colTitle = field.getAnnotation(ExcelColTitle.class);
        if (colTitle == null) return;

        beanField.setTitleRequired(colTitle.required());
        beanField.setTitle(colTitle.value());
    }

    private CellStyle setAlign(Sheet sheet, ExcelColStyle colStyle) {
        var style = sheet.getWorkbook().createCellStyle();
        val align = colStyle.align();
        if (align == ExcelColAlign.LEFT) {
            style.setAlignment(HorizontalAlignment.LEFT);
        } else if (align == ExcelColAlign.CENTER) {
            style.setAlignment(HorizontalAlignment.CENTER);
        } else if (align == ExcelColAlign.RIGHT) {
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
