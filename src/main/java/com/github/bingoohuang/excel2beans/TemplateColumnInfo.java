package com.github.bingoohuang.excel2beans;

import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.ArrayList;

@Data
@Accessors(fluent = true)
public class TemplateColumnInfo {
  int seq = -1;
  ArrayList<CellStyle> styles;
  String title;
  String example;

  public CellStyle tryStyle(int index) {
    if (index >= 0 && index < styles.size()) {
      return styles.get(index);
    }

    return null;
  }
}
