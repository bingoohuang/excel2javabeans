package com.github.bingoohuang.excel2beans;

import lombok.val;
import org.junit.Test;

import java.util.List;

import static com.github.bingoohuang.excel2beans.ExcelToBeansUtils.getClassPathInputStream;
import static com.google.common.truth.Truth.assertThat;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/15.
 */
public class MemberImportBeanWithBlankHeadRowsAndColsTest {

    @Test public void testWithBlankHeadRowsAndCols() {
        val workbookInputStream = getClassPathInputStream("member-blankheadrowsandcols.xlsx");
        val excelToBeans = new ExcelToBeans(MemberImportBeanWithBlankHeadRowsAndCols.class);
        List<MemberImportBeanWithBlankHeadRowsAndCols> beans = excelToBeans.convert(workbookInputStream);
        assertThat(beans).hasSize(2);

        assertThat(beans.get(0).getRowNum()).isEqualTo(5);
        assertThat(beans.get(1).getRowNum()).isEqualTo(6);

        assertThat(beans.get(0)).isEqualTo(MemberImportBeanWithBlankHeadRowsAndCols.builder().memberName("张小凡").sex("女").mobile("18795952311").cardName("示例次卡（100次次卡）").totalTimes("100").availableTimes("90").build());
        assertThat(beans.get(1)).isEqualTo(MemberImportBeanWithBlankHeadRowsAndCols.builder().memberName("李红").sex("男").mobile("18676952432").cardName("示例年卡（一周3次年卡）").totalTimes(null).availableTimes(null).build());
    }
}
