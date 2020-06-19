package com.github.bingoohuang.beans2excel;

import com.github.bingoohuang.excel2beans.PoiUtil;
import com.github.bingoohuang.excel2beans.TemplateCreator;
import com.github.bingoohuang.excel2beans.TitleColumn;
import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import lombok.Data;
import lombok.experimental.Accessors;
import lombok.val;
import org.junit.Test;

import java.util.List;

import static com.github.bingoohuang.excel2beans.PoiUtil.getClassPathWorkbook;
import static com.google.common.collect.Lists.newArrayList;

public class ColumnsExtendTest {
  @Data
  @Accessors(fluent = true)
  public static class ColumnBean {
    // 模板中提供的固定字段，用于定位列。下面动态字段将在该列后面创建，并且复制此列的样式
    @ExcelColTitle("姓名")
    String name;

    // 需要添加的单个动态列，样式拷贝前面提供的定位列。
    @ExcelColTitle TitleColumn titleCol0;
    // 需要添加的多个动态列，样式拷贝前面提供的定位列。
    @ExcelColTitle List<TitleColumn> titleCol1;
    @ExcelColTitle List<TitleColumn> titleCol2;

    @ExcelColTitle("地区")
    String area;

    @ExcelColTitle List<TitleColumn> titleCol3;
  }

  @Test
  public void createTemplates() {
    ColumnBean b =
        new ColumnBean()
            .name("示例-张三") // 模板中提供的固定字段，用于定位列
            .titleCol0(new TitleColumn().title("年龄").example("18")) // 单个动态列
            .titleCol1(
                newArrayList(
                    new TitleColumn().title("性别").example("女"),
                    new TitleColumn().title("城市").example("梵蒂冈"))) // 多个动态列
            .titleCol2(
                newArrayList(
                    new TitleColumn().title("学历").example("博士"),
                    new TitleColumn().title("学校").example("西南联大"))) // 多个动态列
            .area("示例-东京") // 模板中提供的固定字段，用于定位列
            .titleCol3(
                newArrayList(
                    new TitleColumn().title("血型").example("O"),
                    new TitleColumn().title("血压").example("140/90"))) // 多个动态列
        ;

    val creator = new TemplateCreator(getClassPathWorkbook("column-extend-meta.xlsx"));
    creator.create(b);

    PoiUtil.writeExcel(creator.getWorkbook(), "column-extend.xlsx");
  }
}
