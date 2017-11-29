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

import static com.google.common.truth.Truth.assertThat;

public class ImageTest {
    @ExcelSheet(name = "工作表1") @Data
    public static class ImageBean {
        @ExcelColTitle("图片")
        private ImageData imageData;
        @ExcelColTitle("名字")
        private String name;
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

    public void testImage(Workbook workbook) {
        val excelToBeans = new ExcelToBeans(workbook);
        val beans = excelToBeans.convert(ImageBean.class);
        assertThat(beans.size()).isEqualTo(3);
        assertThat(beans.get(0).name).isEqualTo("健身男");
        assertThat(beans.get(1).name).isEqualTo("健身女");
        assertThat(beans.get(2).name).isEqualTo("越野赛");

        val image0Name = createPicture(beans.get(0).imageData);
        System.out.println("健身男：" + image0Name);

        val image1Name = createPicture(beans.get(1).imageData);
        System.out.println("健身女：" + image1Name);

        val image2Name = createPicture(beans.get(2).imageData);
        System.out.println("越野赛：" + image2Name);
    }

    @SneakyThrows
    public static String createPicture(PictureData picture) {
        if (picture == null) {
            return "null";
        }

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
        @Cleanup val workbook = ExcelToBeansUtils.getClassPathWorkbook("images.xlsx");

        val allPictures = workbook.getAllPictures();
        val sheet = workbook.getSheetAt(1);
        val drawingPatriarch = sheet.getDrawingPatriarch();
        if (drawingPatriarch instanceof XSSFDrawing) {
            val xssfDrawing = (XSSFDrawing) drawingPatriarch;
            for (val shape : xssfDrawing.getShapes()) {
                if (shape instanceof XSSFPicture) {
                    val picture = (XSSFPicture) shape;
                    val clientAnchor = picture.getPreferredSize();
                    val from = clientAnchor.getFrom();
                    val pictureName = createPicture(picture.getPictureData());
                    System.out.println("row:" + from.getRow() + ", col:" + from.getCol() + ",pictureName:" + pictureName);
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
                        System.out.println("row:" + hssfClientAnchor.getRow1() + ", col:" + hssfClientAnchor.getCol1() + ",pictureName:" + pictureName);
                    }
                }
            }
        }
    }
}
