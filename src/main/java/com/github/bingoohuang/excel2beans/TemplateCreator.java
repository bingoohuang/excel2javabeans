package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.github.bingoohuang.utils.type.Generic;
import lombok.Getter;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.ss.usermodel.*;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.apache.commons.lang3.StringUtils.isNotEmpty;

public class TemplateCreator {
  @Getter private final Workbook workbook;

  public TemplateCreator(Workbook workbook) {
    this.workbook = workbook;
  }

  @SneakyThrows
  private List<TemplateColumnInfo> parseColumnInfos(
      Object columnBean, Sheet sheet, Map<String, Integer> titleColumnIndexes, int titleRowIndex) {

    ArrayList<CellStyle> styles = null;
    List<TemplateColumnInfo> columnInfos = new ArrayList<>(30);

    for (val field : columnBean.getClass().getDeclaredFields()) {
      val excelColTitle = field.getAnnotation(ExcelColTitle.class);
      if (excelColTitle == null) {
        continue;
      }

      field.setAccessible(true);

      TemplateColumnInfo v;
      String title = excelColTitle.value();
      if (isNotEmpty(title)) { // 定位列
        Integer cellIndex = titleColumnIndexes.get(title);

        styles = createStyles(sheet, titleRowIndex, cellIndex);
        v =
            new TemplateColumnInfo()
                .seq(cellIndex)
                .styles(styles)
                .title(title)
                .example(tryExample(columnBean, field, sheet, titleRowIndex + 1, cellIndex));
        columnInfos.add(v);
        continue;
      }

      if (styles == null) {
        styles = createStyles(sheet, titleRowIndex, 0);
      }

      if (field.getType() == TitleColumn.class) {

        TitleColumn o = (TitleColumn) field.get(columnBean);
        if (o != null) {
          v = new TemplateColumnInfo().styles(styles).title(o.title()).example(o.example());
          columnInfos.add(v);
        }

        continue;
      }

      val gtu = Generic.of(field.getGenericType());
      if (!gtu.isParameterized()
          || !List.class.isAssignableFrom(field.getType())
          || gtu.getActualTypeArg(0) != TitleColumn.class) {
        continue;
      }

      @SuppressWarnings("unchecked")
      List<TitleColumn> os = (List<TitleColumn>) field.get(columnBean);
      if (os == null) {
        continue;
      }

      for (val o : os) {
        v = new TemplateColumnInfo().styles(styles).title(o.title()).example(o.example());
        columnInfos.add(v);
      }
    }

    int lastSeq = columnInfos.get(0).seq();
    if (lastSeq < 0) {
      lastSeq = 0;
    }

    for (int i = 0; i < columnInfos.size(); i++) {
      columnInfos.get(i).seq(lastSeq + i);
    }

    return columnInfos;
  }

  @SneakyThrows
  private String tryExample(Object columnBean, Field field, Sheet sheet, int row, int column) {
    Object o = field.get(columnBean);
    if (o != null) {
      String os = o.toString();
      if (isNotEmpty(os)) {
        return os;
      }
    }

    Row r = sheet.getRow(row);
    if (r == null) {
      return "";
    }

    Cell c = r.getCell(column);
    if (c == null) {
      return "";
    }

    return c.getStringCellValue();
  }

  private ArrayList<CellStyle> createStyles(Sheet sheet, int titleRowIndex, int cellIndex) {
    ArrayList<CellStyle> styles = new ArrayList<>();

    for (int k = titleRowIndex, kk = sheet.getLastRowNum(); k <= kk; ++k) {
      val row = sheet.getRow(k);
      Cell cell = row.getCell(cellIndex);
      styles.add(cell != null ? cell.getCellStyle() : null);
    }

    return styles;
  }

  private int locationTitleRow(
      Sheet sheet, Class<?> columnBeanClass, Map<String, Integer> titleColumnIndexes) {
    val declaredFields = columnBeanClass.getDeclaredFields();
    List<String> titles = new ArrayList<>(declaredFields.length);

    for (val field : declaredFields) {
      val excelColTitle = field.getAnnotation(ExcelColTitle.class);
      if (excelColTitle != null && isNotEmpty(excelColTitle.value())) {
        titles.add(excelColTitle.value());
      }
    }

    return findTitleRowNum(sheet, titles, titleColumnIndexes);
  }

  public int findTitleRowNum(
      Sheet sheet, List<String> titles, Map<String, Integer> titleColumnIndexes) {
    int i = sheet.getFirstRowNum();

    for (int ii = sheet.getLastRowNum(); i <= ii; ++i) {
      val row = sheet.getRow(i);

      if (findRowWithTitles(row, titles, titleColumnIndexes)) {
        return i;
      }
    }

    throw new IllegalArgumentException("Unable to find title row.");
  }

  private boolean findRowWithTitles(
      Row row, List<String> titles, Map<String, Integer> titleColumnIndexes) {
    for (val title : titles) {
      if (!findColumn(row, title, titleColumnIndexes)) {
        return false;
      }
    }
    return true;
  }

  private boolean findColumn(Row row, String title, Map<String, Integer> titleColumnIndexes) {
    for (int k = row.getFirstCellNum(), kk = row.getLastCellNum(); k < kk; ++k) {
      val cell = row.getCell(k);
      if (cell == null) {
        continue;
      }

      val cellValue = cell.getStringCellValue();
      if (containTitle(cellValue, title)) {
        titleColumnIndexes.put(title, k);
        return true;
      }
    }

    return false;
  }

  public boolean containTitle(String cellValue, String title) {
    return cellValue != null && cellValue.toUpperCase().contains(title);
  }

  public void create(Object columnBean) {
    val sheet = PoiUtil.findSheet(workbook, columnBean.getClass());

    Map<String, Integer> titleColumnIndexes = new HashMap<>();

    // Step 1: 定位Title行
    int titleRowIndex = locationTitleRow(sheet, columnBean.getClass(), titleColumnIndexes);
    // Step 2: 准备好要写的Title行Bean
    List<TemplateColumnInfo> columnInfos =
        parseColumnInfos(columnBean, sheet, titleColumnIndexes, titleRowIndex);

    int rowOff = 0;

    for (int i = titleRowIndex, kk = sheet.getLastRowNum(); i <= kk; ++i, ++rowOff) {
      Row row = sheet.getRow(i);

      for (val columnInfo : columnInfos) {
        Cell cell = row.createCell(columnInfo.seq());

        if (rowOff == 0) {
          cell.setCellValue(columnInfo.title());
        } else if (rowOff == 1) {
          cell.setCellValue(columnInfo.example());
        }

        CellStyle style = columnInfo.tryStyle(rowOff);
        if (style != null) {
          cell.setCellStyle(style);
        }
      }
    }
  }
}
