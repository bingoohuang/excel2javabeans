package com.github.bingoohuang.excel2beans;

import com.google.common.io.ByteStreams;
import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.experimental.UtilityClass;
import lombok.val;
import org.apache.poi.ss.usermodel.Workbook;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.net.URLEncoder;

@UtilityClass
public class ExcelDownloads {
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
    public ServletOutputStream prepareDownload(HttpServletResponse response, String fileName) {
        response.setContentType("application/vnd.ms-excel;charset=UTF-8");
        val encodedFileName = URLEncoder.encode(fileName, "UTF-8");
        response.setHeader("Content-disposition", "attachment; " +
                "filename=\"" + encodedFileName + "\"; " +
                "filename*=utf-8'zh_cn'" + encodedFileName);
        return response.getOutputStream();
    }
}
