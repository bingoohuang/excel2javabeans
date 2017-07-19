package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.asmvalidator.annotations.*;
import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import lombok.*;
import org.apache.commons.lang3.StringUtils;
import org.junit.Test;

import java.util.List;

import static com.google.common.truth.Truth.assertThat;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/15.
 */
public class MemberImportBeanWithBlankHeadRowsAndColsTest {
    @SneakyThrows
    @Test public void testWithBlankHeadRowsAndCols() {
        @Cleanup val workbook = ExcelToBeansUtils.getClassPathWorkbook("member-blankheadrowsandcols.xlsx");
        val excelToBeans = new ExcelSheetToBeans(workbook, MemberImportBeanWithBlankHeadRowsAndCols.class);
        List<MemberImportBeanWithBlankHeadRowsAndCols> beans = excelToBeans.convert();
        assertThat(beans).hasSize(2);

        assertThat(beans.get(0).getRowNum()).isEqualTo(5);
        assertThat(beans.get(1).getRowNum()).isEqualTo(6);

        assertThat(beans.get(0)).isEqualTo(MemberImportBeanWithBlankHeadRowsAndCols.builder().memberName("张小凡").sex("女").mobile("18795952311").cardName("示例次卡（100次次卡）").totalTimes("100").availableTimes("90").build());
        assertThat(beans.get(1)).isEqualTo(MemberImportBeanWithBlankHeadRowsAndCols.builder().memberName("李红").sex("男").mobile("18676952432").cardName("示例年卡（一周3次年卡）").totalTimes(null).availableTimes(null).build());
    }

    /**
     * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
     */
    @Data @Builder
    public static class MemberImportBeanWithBlankHeadRowsAndCols extends ExcelRowRef implements ExcelRowIgnorable {
        @ExcelColTitle("会员姓名") @AsmMaxSize(12) @AsmMessage("请填写会员姓名") String memberName; // 不超过12字
        @ExcelColTitle("性别") @AsmRange("男,女") @AsmMessage("性别请填男或女") String sex;
        @ExcelColTitle("手机号") @AsmMobile @AsmMessage("请填写正确的手机号码") String mobile;
        @ExcelColTitle("卡名称") @MemberCardName String cardName;
        @ExcelColTitle("总次数") @AsmBlankable @AsmDigits String totalTimes;
        @ExcelColTitle("剩余次数") @AsmBlankable @AsmDigits String availableTimes;

        @Override public boolean ignoreRow() {
            return StringUtils.startsWith(memberName, "示例-");
        }
    }
}
