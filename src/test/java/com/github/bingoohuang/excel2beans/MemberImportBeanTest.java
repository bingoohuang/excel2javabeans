package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.asmvalidator.annotations.*;
import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import lombok.*;
import org.apache.commons.lang3.StringUtils;
import org.junit.Test;

import static com.google.common.truth.Truth.assertThat;

public class MemberImportBeanTest {
    @SneakyThrows
    @Test public void testWithoutBlankHeadRowsAndCols() {
        @Cleanup val workbook = PoiUtil.getClassPathWorkbook("member.xlsx");
        val excelToBeans = new ExcelToBeans(workbook);
        val beans = excelToBeans.convert(MemberImportBean.class);
        assertThat(beans).hasSize(4);

        assertThat(beans.get(0).getRowNum()).isEqualTo(6);
        assertThat(beans.get(1).getRowNum()).isEqualTo(7);
        assertThat(beans.get(2).getRowNum()).isEqualTo(8);
        assertThat(beans.get(3).getRowNum()).isEqualTo(9);

        beans.get(0).setError("error 000");
        beans.get(1).setError("error 000");

        excelToBeans.writeError(MemberImportBean.class, beans);
        excelToBeans.removeOkRows(MemberImportBean.class, beans);

        PoiUtil.writeExcel(excelToBeans.getWorkbook(), "member-error.xlsx");
    }

    @Data @Builder
    public static class MemberImportBean extends ExcelRowRef implements ExcelRowIgnorable {
        @ExcelColTitle("会员姓名") @AsmMaxSize(12) @AsmMessage("请填写会员姓名") String memberName; // 不超过12字
        @ExcelColTitle @AsmRange("男,女") @AsmMessage("性别请填男或女") String sex;
        @ExcelColTitle("手机号") @AsmMobile @AsmMessage("请填写正确的手机号码") String mobile;
        @ExcelColTitle("生日") @AsmBlankable @AsmMessage("请填写正确的生日") @MemberCardBirthDay String birthday;
        @ExcelColTitle("卡名称") @MemberCardName String cardName;
        @ExcelColTitle("办卡价格") @AsmBlankable @AsmDigits String cardPrice;
        @ExcelColTitle("消费上限") @AsmBlankable @AsmDigits String upperTimes;
        @ExcelColTitle("总次数") @AsmBlankable @AsmDigits String totalTimes;
        @ExcelColTitle("剩余次数") @AsmBlankable @AsmDigits String availableTimes;
        @ExcelColTitle("有效期开始日") @MemberCardEffDay String effectiveTime;
        @ExcelColTitle("有效期截止日") @MemberCardEffDay String expiredTime;

        @Override public boolean ignoreRow() {
            return StringUtils.startsWith(memberName, "示例-");
        }


        private String error;

        @Override public String error() {
            return error;
        }
    }
}
