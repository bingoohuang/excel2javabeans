package com.github.bingoohuang.excel2maps;

import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.val;
import org.junit.Test;

import static com.github.bingoohuang.excel2beans.ExcelToBeansUtils.getClassPathWorkbook;
import static com.google.common.collect.ImmutableMap.of;
import static com.google.common.truth.Truth.assertThat;

public class ExcelSheetToMapsTest {
    @Test @SneakyThrows public void test1() {
        val excel2MapsConfig = new ExcelToMapsConfig();
        excel2MapsConfig.add(new ColumnDef("会员姓名", "memberName", "示例-*"));
        excel2MapsConfig.add(new ColumnDef("性别", "sex"));
        val workbook = getClassPathWorkbook("member.xlsx");
        @Cleanup val excel2Maps = new ExcelToMaps(workbook);

        val maps = excel2Maps.convert(excel2MapsConfig, 0);
        assertThat(maps).hasSize(4);
        assertThat(maps.get(0)).isEqualTo(of("memberName", "张小凡", "sex", "女", "_row", "6"));
        assertThat(maps.get(1)).isEqualTo(of("memberName", "李红", "sex", "男", "_row", "7"));
        assertThat(maps.get(2)).isEqualTo(of("memberName", "李红", "sex", "男", "_row", "8"));
        assertThat(maps.get(3)).isEqualTo(of("memberName", "张晓", "sex", "女", "_row", "9"));
    }
}
