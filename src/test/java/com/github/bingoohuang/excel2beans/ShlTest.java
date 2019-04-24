package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import lombok.*;
import org.junit.Test;

import static com.google.common.truth.Truth.assertThat;

public class ShlTest {
  @Test
  @SneakyThrows
  public void member() {
    @Cleanup val workbook = PoiUtil.getClassPathWorkbook("shl.xls");
    val excelToBeans = new ExcelToBeans(workbook);
    val beans = excelToBeans.convert(ShlTest.ShlBean.class);
    assertThat(beans).hasSize(12);
  }

  @Data
  @Builder
  public static class ShlBean {
    @ExcelColTitle("项目名称")
    private String name;

    @ExcelColTitle("Deciding and Initiating Action")
    private int action;
  }
}
