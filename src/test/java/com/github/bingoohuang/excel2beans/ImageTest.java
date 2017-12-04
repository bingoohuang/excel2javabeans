package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.github.bingoohuang.excel2beans.annotations.ExcelSheet;
import com.github.bingoohuang.westid.WestId;
import lombok.Cleanup;
import lombok.Data;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.junit.Ignore;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

import static com.google.common.truth.Truth.assertThat;

public class ImageTest {
    @ExcelSheet(name = "工作表1") @Data
    public static class ImageBean {
        @ExcelColTitle("图片")
        private ImageData imageData;
        @ExcelColTitle("名字")
        private String name;
    }

    @Data
    public static class ImageListBean {
        @ExcelColTitle("图片")
        private List<ImageData> imageDatas;
        @ExcelColTitle("名字")
        private String name;
    }

    @Test @SneakyThrows
    public void testImageList() {
        @Cleanup val workbook = ExcelToBeansUtils.getClassPathWorkbook("multi-images.xlsx");
        val excelToBeans = new ExcelToBeans(workbook);
        val beans = excelToBeans.convert(ImageListBean.class);
        assertThat(beans.size()).isEqualTo(3);
        assertThat(beans.get(0).name).isEqualTo("健身男");
        assertThat(beans.get(0).imageDatas.get(0).getData().length).isEqualTo(255429);
        assertThat(beans.get(0).imageDatas.get(1).getData().length).isEqualTo(1682552);
        assertThat(beans.get(1).name).isEqualTo("健身女");
        assertThat(beans.get(1).imageDatas.get(0).getData().length).isEqualTo(373333);
        assertThat(beans.get(1).imageDatas.get(1).getData().length).isEqualTo(1560588);
        assertThat(beans.get(2).name).isEqualTo("越野赛");
        assertThat(beans.get(2).imageDatas.get(0).getData().length).isEqualTo(3700955);
        assertThat(beans.get(2).imageDatas.get(1).getData().length).isEqualTo(1663205);

//        val image0Name = createPicture(beans.get(0).imageDatas.get(1));
//        System.out.println("健身男：" + image0Name);
//
//        val image1Name = createPicture(beans.get(1).imageDatas.get(1));
//        System.out.println("健身女：" + image1Name);
//
//        val image2Name = createPicture(beans.get(2).imageDatas.get(1));
//        System.out.println("越野赛：" + image2Name);
    }


    @Test @SneakyThrows
    public void testXls() {
        @Cleanup val workbook = ExcelToBeansUtils.getClassPathWorkbook("images.xls");
        testImage(workbook);
    }

    @Test @SneakyThrows
    public void testXlsx() {
        @Cleanup val workbook = ExcelToBeansUtils.getClassPathWorkbook("images.xlsx");
        testImage(workbook);
    }

    @Test @SneakyThrows
    public void testCenterXls() {
        @Cleanup val workbook = ExcelToBeansUtils.getClassPathWorkbook("center-images.xls");
        testImage(workbook);
    }

    @Test @SneakyThrows
    public void testCenterXlsx() {
        @Cleanup val workbook = ExcelToBeansUtils.getClassPathWorkbook("center-images.xlsx");
        testImage(workbook);
    }

    public void testImage(Workbook workbook) {
        val excelToBeans = new ExcelToBeans(workbook);
        val beans = excelToBeans.convert(ImageBean.class);
        assertThat(beans.size()).isEqualTo(3);
        assertThat(beans.get(0).name).isEqualTo("健身男");
        assertThat(beans.get(0).imageData.getData().length).isEqualTo(255429);
        assertThat(beans.get(1).name).isEqualTo("健身女");
        assertThat(beans.get(1).imageData.getData().length).isEqualTo(373333);
        assertThat(beans.get(2).name).isEqualTo("越野赛");
        assertThat(beans.get(2).imageData.getData().length).isEqualTo(3700955);
    }

    @SneakyThrows
    public static String createPicture(PictureData picture) {
        if (picture == null) return "null";

        val extension = picture.suggestFileExtension();
        val imageFileName = String.valueOf(WestId.next()) + "." + extension;
        File file = new File(imageFileName);
        file.deleteOnExit(); // comment out this for human assertion
        @Cleanup val out = new FileOutputStream(file);
        out.write(picture.getData());

        return imageFileName;
    }

    @Test @SneakyThrows @Ignore
    public void test2() {
        @Cleanup val workbook1 = ExcelToBeansUtils.getClassPathWorkbook("center-images.xls");
        computeImagePosition(workbook1);
        System.out.println();
        System.out.println();
        @Cleanup val workbook2 = ExcelToBeansUtils.getClassPathWorkbook("center-images.xlsx");
        computeImagePosition(workbook2);
    }

    private void computeImagePosition(Workbook workbook) {
        val allPictures = workbook.getAllPictures();
        val sheet = workbook.getSheetAt(1);
        val drawingPatriarch = sheet.getDrawingPatriarch();
        if (drawingPatriarch instanceof XSSFDrawing) {
            val xssfDrawing = (XSSFDrawing) drawingPatriarch;
            for (val shape : xssfDrawing.getShapes()) {
                if (shape instanceof XSSFPicture) {
                    val picture = (XSSFPicture) shape;
                    val clientAnchor = picture.getClientAnchor();

                    val from = clientAnchor.getFrom();
                    val to = clientAnchor.getTo();
                    val pictureName = createPicture(picture.getPictureData());
                    System.out.println("from row:" + from.getRow() + ", from col:" + from.getCol() + ",pictureName:" + pictureName);

                    int axisRowIndex = ExcelToBeansUtils.computeAxisRowIndex(sheet, picture);
                    int axisColIndex = ExcelToBeansUtils.computeAxisColIndex(sheet, picture);

                    System.out.println("axisRowIndex: " + axisRowIndex + ",axisColIndex:" + axisColIndex);
                }
            }
        } else if (drawingPatriarch instanceof HSSFPatriarch) {
            val hssfPatriarch = (HSSFPatriarch) drawingPatriarch;
            for (val shape : hssfPatriarch.getChildren()) {
                if (shape instanceof HSSFPicture) {
                    val hssfPicture = (HSSFPicture) shape;
                    val pictureIndex = hssfPicture.getPictureIndex();
                    val picture = allPictures.get(pictureIndex - 1);
                    val anchor = hssfPicture.getAnchor();
                    if (anchor instanceof HSSFClientAnchor) {
                        val hssfClientAnchor = (HSSFClientAnchor) anchor;
                        val pictureName = createPicture(picture);
                        System.out.println("row1:" + hssfClientAnchor.getRow1() + ", col1:" + hssfClientAnchor.getCol1() + ",pictureName:" + pictureName);

                        int axisRowIndex = ExcelToBeansUtils.computeAxisRowIndex(sheet, hssfPicture);
                        int axisColIndex = ExcelToBeansUtils.computeAxisColIndex(sheet, hssfPicture);

                        System.out.println("axisRowIndex: " + axisRowIndex + ",axisColIndex:" + axisColIndex);
                    }
                }
            }
        }
    }

}
