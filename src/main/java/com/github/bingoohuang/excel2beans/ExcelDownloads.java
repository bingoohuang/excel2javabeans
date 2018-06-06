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
    public void download(HttpServletResponse r, Workbook wb, String fileName) {
        @Cleanup val out = prepareDownload(r, fileName);
        wb.write(out);
        wb.close();
    }

    @SneakyThrows
    public void download(HttpServletResponse r, byte[] wb, String fileName) {
        @Cleanup val out = prepareDownload(r, fileName);
        out.write(wb);
    }

    @SneakyThrows
    public void download(HttpServletResponse r, InputStream wb, String fileName) {
        @Cleanup val out = prepareDownload(r, fileName);
        ByteStreams.copy(wb, out);
        wb.close();
    }

    @SneakyThrows
    public ServletOutputStream prepareDownload(HttpServletResponse r, String fileName) {
        r.setContentType("application/vnd.ms-excel;charset=UTF-8");
        val f = URLEncoder.encode(fileName, "UTF-8");
        r.setHeader("Content-disposition", "attachment; filename=\"" + f + "\"; filename*=utf-8'zh_cn'" + f);
        return r.getOutputStream();
    }
}
