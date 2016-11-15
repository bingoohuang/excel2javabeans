package com.github.bingoohuang.excel2maps;

import lombok.val;
import org.junit.Test;

import static com.github.bingoohuang.excel2beans.ExcelToBeansUtils.getClassPathInputStream;
import static com.google.common.collect.ImmutableMap.of;
import static com.google.common.truth.Truth.assertThat;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/15.
 */
public class Excel2MapsTest {
    @Test public void test1() {
        val excel2MapsConfig = new Excel2MapsConfig();
        excel2MapsConfig.add(new ColumnDef("会员姓名", "memberName", "示例-*"));
        excel2MapsConfig.add(new ColumnDef("性别", "sex"));
        val excel2Maps = new Excel2Maps(excel2MapsConfig);
        val workbookInputStream = getClassPathInputStream("member.xlsx");

        val maps = excel2Maps.convert(workbookInputStream);
        assertThat(maps).hasSize(4);
        assertThat(maps.get(0)).isEqualTo(of("memberName", "张小凡", "sex", "女", "_rowNum", "6"));
        assertThat(maps.get(1)).isEqualTo(of("memberName", "李红", "sex", "男", "_rowNum", "7"));
        assertThat(maps.get(2)).isEqualTo(of("memberName", "李红", "sex", "男", "_rowNum", "8"));
        assertThat(maps.get(3)).isEqualTo(of("memberName", "张晓", "sex", "女", "_rowNum", "9"));
    }
}
