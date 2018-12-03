package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelSheet;
import com.github.bingoohuang.utils.lang.Classpath;
import com.google.common.collect.Lists;
import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import lombok.val;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

import static com.github.bingoohuang.utils.codec.Bytes.toByteArray;
import static org.apache.commons.lang3.StringUtils.endsWithIgnoreCase;

@Slf4j
public class PoiUtil {
    /**
     * 用密码保护工作簿。只有在xlsx格式才起作用。
     *
     * @param workbook 工作簿。
     * @param password 保护密码。
     */
    public static void protectWorkbook(Workbook workbook, String password) {
        if (StringUtils.isEmpty(password)) return;

        if (workbook instanceof XSSFWorkbook) {
            val xsswb = (XSSFWorkbook) workbook;
            for (int i = 0, ii = xsswb.getNumberOfSheets(); i < ii; ++i) {
                xsswb.getSheetAt(i).protectSheet(password);
            }
        }
    }

    public static void blankCell(Row row, int col) {
        val cell = row.getCell(col);
        if (cell == null) return;

        if (cell.getCellTypeEnum() != CellType.BLANK) {
            cell.setCellType(CellType.BLANK);
        }
    }

    public static String getCellValue(Row row, int col) {
        val cell = row.getCell(col);
        if (cell == null) return null;

        return getCellStringValue(cell);
    }

    /**
     * 删除其它Sheet页，只保留指定模板的Sheet页输出。
     *
     * @param sheet 在工作簿需要保留的表单页。
     */
    public static void removeOtherSheets(Sheet sheet) {
        val wb = sheet.getWorkbook();

        List<String> deletedNames = Lists.newArrayList();
        for (int i = 0; i < wb.getNumberOfSheets(); ++i) {
            val sheetName = wb.getSheetName(i);
            if (!sheetName.equals(sheet.getSheetName())) {
                deletedNames.add(sheetName);
            }
        }

        deletedNames.forEach(x -> wb.removeSheetAt(wb.getSheetIndex(x)));
    }

    /**
     * 是否是完整的单元格索引。例如A5是完整的，A不是完整的，缺少行标。
     *
     * @param cellReference EXCEL中的单元格索引。
     * @return 完整的单元格索引时返回true。
     */
    public static boolean isFullCellReference(String cellReference) {
        return cellReference.matches("\\w+\\d+");
    }

    /**
     * 是否是列索引。例如A是列索引。
     *
     * @param cellReference EXCEL中的单元格索引。
     * @return true时是列索引。
     */
    public static boolean isColReference(String cellReference) {
        return cellReference.matches("\\w+");
    }

    /**
     * 是否是行索引。例如5是列索引。
     *
     * @param cellReference EXCEL中的单元格索引。
     * @return true时是行索引。
     */
    public static boolean isRowReference(String cellReference) {
        return cellReference.matches("\\d+");
    }

    /**
     * 插入新的行（向下移动表格中的行），方便写入数据。
     *
     * @param sheet          表单页。
     * @param list           写入JavaBean列表。
     * @param templateRowNum 模板行号。
     * @return JavaBean列表大小。
     */
    public static int shiftRows(Sheet sheet, List<?> list, int templateRowNum) {
        val itemSize = list == null ? 0 : list.size();
        int lastRowNum = sheet.getLastRowNum();
        if (itemSize == 0 && templateRowNum == lastRowNum) { // 没有写入行，直接删除模板行
            sheet.removeRow(sheet.getRow(templateRowNum));
        } else if (itemSize != 1 && templateRowNum < lastRowNum) {
            sheet.shiftRows(templateRowNum + 1, lastRowNum, itemSize - 1);
        }

        return itemSize;
    }

    /**
     * 向单元格写入值，处理值为整型时的写入情况。
     *
     * @param cell 单元格。
     * @param fv   单元格值。
     * @return 单元格字符串取值。
     */
    public static String writeCellValue(Cell cell, Object fv) {
        if (fv == null) {
            cell.setCellType(CellType.BLANK);
            return "";
        }

        if (fv instanceof Number) {
            val value = ((Number) fv).doubleValue();
            cell.setCellValue(value);
            return "" + value;
        }

        if (fv instanceof String) {
            cell.setCellValue((String) fv);
            return (String) fv;
        }

        val value = "" + fv;
        cell.setCellValue(value);
        return value;
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
    public static void writeExcel(Workbook workbook, File name) {
        @Cleanup val fileOut = new FileOutputStream(name);
        workbook.write(fileOut);
    }

    public static void writeExcel(Workbook workbook, String name) {
        writeExcel(workbook, new File(name));
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

    /**
     * 根据单元格索引，找到单元格。
     *
     * @param sheet        EXCEL表单。
     * @param cellRefValue 单元格索引，例如A1, AB12等。
     * @return 单元格。
     */
    public static Cell findCell(Sheet sheet, String cellRefValue) {
        val cellRef = new CellReference(cellRefValue);
        val row = sheet.getRow(cellRef.getRow());
        if (row == null) {
            log.warn("unable to find row for " + cellRefValue);
            return null;
        }

        val cell = row.getCell(cellRef.getCol());
        if (cell == null) {
            log.warn("unable to find cell for " + cellRefValue);
        }

        return cell;
    }

    /**
     * 修正图表中对于表单名字的引用。
     *
     * @param sheet        EXCEL表单。
     * @param oldSheetName 旧的表单名字。
     * @param newSheetName 新的表单名字。
     */
    public static void fixChartSheetNameRef(Sheet sheet, String oldSheetName, String newSheetName) {
        val drawing = sheet.getDrawingPatriarch();
        if (!(drawing instanceof XSSFDrawing)) return;

        for (val chart : ((XSSFDrawing) drawing).getCharts()) {
            for (val barChart : chart.getCTChart().getPlotArea().getBarChartList()) {
                for (val ser : barChart.getSerList()) {
                    val val = ser.getVal();
                    if (val == null) continue;

                    val numRef = val.getNumRef();
                    if (numRef == null) continue;

                    val f = numRef.getF();
                    if (f == null) continue;

                    if (f.contains(oldSheetName)) {
                        numRef.setF(f.replace(oldSheetName, newSheetName));
                    }
                }
            }
        }
    }

    public static String getCellStringValue(Cell cell) {
        switch (cell.getCellTypeEnum()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BLANK:
                return "";
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                CellValue cv = getFormulaCellValue(cell);
                return cv.getStringValue();
            case ERROR:
                return FormulaError.forInt(cell.getErrorCellValue()).getString();
        }

        cell.setCellType(CellType.STRING);
        return StringUtils.trimToEmpty(cell.getStringCellValue());
    }

    public static CellValue getFormulaCellValue(Cell cell) {
        if (cell == null) return null;

        try {
            return cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator().evaluate(cell);
        } catch (Exception e) {
            log.warn("get formula cell value[{}, {}] error : ", cell.getRowIndex(), cell.getColumnIndex(), e);

            return null;
        }
    }

    /**
     * 查找单元格。
     *
     * @param sheet         表单
     * @param cellReference 单元格索引
     * @param searchKey     单元格中包含的关键字
     * @return 单元格，没有找到时返回null
     */
    public static Cell findCell(Sheet sheet, String cellReference, String searchKey) {
        if (isFullCellReference(cellReference)) {
            return findCell(sheet, cellReference);
        }

        if (isColReference(cellReference)) {
            val cellRef = new CellReference(cellReference + "1");
            return searchColCell(sheet, cellRef.getCol(), searchKey);
        }

        if (isRowReference(cellReference)) {
            val cellRef = new CellReference("A" + cellReference);
            return searchRowCell(sheet, cellRef.getRow(), searchKey);
        }

        return searchCell(sheet, searchKey);
    }

    /**
     * 查找单元格。
     *
     * @param sheet     表单
     * @param colIndex  列索引
     * @param searchKey 单元格中包含的关键字
     * @return 单元格，没有找到时返回null
     */
    public static Cell searchColCell(Sheet sheet, short colIndex, String searchKey) {
        if (StringUtils.isEmpty(searchKey)) return null;

        for (int i = sheet.getFirstRowNum(), ii = sheet.getLastRowNum(); i < ii; ++i) {
            val row = sheet.getRow(i);
            if (row == null) continue;

            val cell = matchCell(row, colIndex, searchKey);
            if (cell != null) return cell;
        }

        return null;
    }

    /**
     * 查找单元格。
     *
     * @param sheet     表单
     * @param rowIndex  行索引
     * @param searchKey 单元格中包含的关键字
     * @return 单元格，没有找到时返回null
     */
    public static Cell searchRowCell(Sheet sheet, int rowIndex, String searchKey) {
        if (StringUtils.isEmpty(searchKey)) return null;

        return searchRow(sheet.getRow(rowIndex), searchKey);
    }

    /**
     * 查找单元格。
     *
     * @param sheet     表单
     * @param searchKey 单元格中包含的关键字
     * @return 单元格，没有找到时返回null
     */
    public static Cell searchCell(Sheet sheet, String searchKey) {
        if (StringUtils.isEmpty(searchKey)) return null;

        for (int i = sheet.getFirstRowNum(), ii = sheet.getLastRowNum(); i < ii; ++i) {
            Cell cell = searchRow(sheet.getRow(i), searchKey);
            if (cell != null) return cell;
        }

        return null;
    }

    /**
     * 在行中查找。
     *
     * @param row       行
     * @param searchKey 单元格中包含的关键字
     * @return 单元格，没有找到时返回null
     */
    public static Cell searchRow(Row row, String searchKey) {
        if (row == null) return null;

        for (int j = row.getFirstCellNum(), jj = row.getLastCellNum(); j < jj; ++j) {
            Cell cell = matchCell(row, j, searchKey);
            if (cell != null) return cell;
        }
        return null;
    }

    /**
     * 匹配单元格
     *
     * @param row       行
     * @param colIndex  列索引
     * @param searchKey 单元格中包含的关键字
     * @return 单元格，没有找到时返回null
     */
    public static Cell matchCell(Row row, int colIndex, String searchKey) {
        val cell = row.getCell(colIndex);
        if (cell == null) return null;

        val value = getCellStringValue(cell);
        if (StringUtils.contains(value, searchKey)) return cell;

        return null;
    }

    @SneakyThrows
    public static Workbook getClassPathWorkbook(String classPathExcelName) {
        @Cleanup val is = Classpath.loadRes(classPathExcelName);
        return WorkbookFactory.create(is);
    }

    @SneakyThrows
    public static byte[] getWorkbookBytes(Workbook workbook) {
        @Cleanup val bout = new ByteArrayOutputStream();
        workbook.write(bout);
        return bout.toByteArray();
    }

    /**
     * 增加一张图片。
     *
     * @param sheet               表单
     * @param cpImageName         类路径中的图片文件名称
     * @param anchorCellReference 图片锚定单元格索引
     */
    @SneakyThrows
    public static void addImage(Sheet sheet, String cpImageName, String anchorCellReference) {
        // add a picture shape
        val anchor = sheet.getWorkbook().getCreationHelper().createClientAnchor();
        anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
        // subsequent call of Picture#resize() will operate relative to it
        val cr = new CellReference(anchorCellReference);
        anchor.setCol1(cr.getCol());
        anchor.setRow1(cr.getRow());

        // Create the drawing patriarch.  This is the top level container for all shapes.
        @Cleanup val p = Classpath.loadRes(cpImageName);

        val picIndex = sheet.getWorkbook().addPicture(toByteArray(p), getPictureType(cpImageName));
        val pic = sheet.createDrawingPatriarch().createPicture(anchor, picIndex);
        // auto-size picture relative to its top-left corner
        pic.resize();
    }

    private static int getPictureType(String classpathImageName) {
        if (endsWithIgnoreCase(classpathImageName, "png")) return Workbook.PICTURE_TYPE_PNG;
        if (endsWithIgnoreCase(classpathImageName, "jpg")) return Workbook.PICTURE_TYPE_JPEG;
        if (endsWithIgnoreCase(classpathImageName, "jpeg")) return Workbook.PICTURE_TYPE_JPEG;

        throw new RuntimeException("unknown format for image file " + classpathImageName);
    }

    /**
     * 查找有值得最大列索引。
     *
     * @param sheet 表单
     * @return 最大列索引
     */
    public static int findMaxCol(Sheet sheet) {
        int maxCol = 0;
        for (int i = 0, ii = sheet.getLastRowNum(); i <= ii; ++i) {
            val row = sheet.getRow(i);
            if (row == null) continue;

            for (int j = row.getLastCellNum() - 1; j > maxCol; --j) {
                val cell = row.getCell(j);
                if (cell == null) continue;

                val value = getCellStringValue(cell);
                if (StringUtils.isNotEmpty(value)) {
                    maxCol = j;
                    break;
                }
            }
        }

        return maxCol;
    }
}
