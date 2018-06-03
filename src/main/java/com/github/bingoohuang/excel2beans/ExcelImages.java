package com.github.bingoohuang.excel2beans;

import com.google.common.collect.HashBasedTable;
import com.google.common.collect.Table;
import lombok.val;
import lombok.var;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.ImageUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;

public class ExcelImages {
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
        HashBasedTable<Integer, Integer, ImageData> images = HashBasedTable.create();
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
        HashBasedTable<Integer, Integer, ImageData> images = HashBasedTable.create();
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
}
