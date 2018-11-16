package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.asmvalidator.annotations.*;
import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.google.common.collect.Lists;
import lombok.*;
import org.apache.commons.lang3.StringUtils;
import org.junit.Test;

import java.util.List;
import java.util.regex.Pattern;

public class ImportMemberExcelRowTest {
    @Test @SneakyThrows
    public void test() {
        @Cleanup val workbook = PoiUtil.getClassPathWorkbook("amita.xlsx");
        val excelToBeans = new ExcelToBeans(workbook);
        excelToBeans.convert(ImportMemberExcelRow.class);

// ImportMemberExcelRowTest.ImportMemberExcelRow(memberName=朱A, sex=女, mobile=13338090200, birthday=null, cardName=会员课两年卡, cardNo=5208810, price=9800, times=无限次, totalTimes=null, availTimes=null, totalMoney=null, availMoney=null, effective=2017-09-11, expired=2019-09-10, latestActivateDay=null, brief=null, advisor=null, memberPinyin=null, memberInitials=null, errorMessage=, checkedPass=false, sameMember=false, userId=null, mbrCardId=0, memberId=null, cardId=null, cardTypeId=null, expiredValue=0, expiredUnit=null, cycleValue=0, cycleUnit=null, state=null),
// ImportMemberExcelRowTest.ImportMemberExcelRow(memberName=朱B, sex=女, mobile=13338090200, birthday=null, cardName=会员课年卡, cardNo=5208041, price=6380, times=无限次, totalTimes=null, availTimes=null, totalMoney=null, availMoney=null, effective=2017-03-16, expired=2018-05-15, latestActivateDay=null, brief=null, advisor=null, memberPinyin=null, memberInitials=null, errorMessage=, checkedPass=false, sameMember=false, userId=null, mbrCardId=0, memberId=null, cardId=null, cardTypeId=null, expiredValue=0, expiredUnit=null, cycleValue=0, cycleUnit=null, state=null),
// ImportMemberExcelRowTest.ImportMemberExcelRow(memberName=朱C, sex=女, mobile=13338090200, birthday=null, cardName=会员课两年卡, cardNo=17271688, price=9800, times=无限次, totalTimes=null, availTimes=null, totalMoney=null, availMoney=null, effective=2016-09-10, expired=2018-09-10, latestActivateDay=null, brief=null, advisor=null, memberPinyin=null, memberInitials=null, errorMessage=, checkedPass=false, sameMember=false, userId=null, mbrCardId=0, memberId=null, cardId=null, cardTypeId=null, expiredValue=0, expiredUnit=null, cycleValue=0, cycleUnit=null, state=null),
// ImportMemberExcelRowTest.ImportMemberExcelRow(memberName=郑D, sex=女, mobile=13338090200, birthday=null, cardName=会员课10次卡, cardNo=17051688, price=1280, times=null, totalTimes=10, availTimes=10, totalMoney=null, availMoney=null, effective=null, expired=null, latestActivateDay=null, brief=null, advisor=null, memberPinyin=null, memberInitials=null, errorMessage=, checkedPass=false, sameMember=false, userId=null, mbrCardId=0, memberId=null, cardId=null, cardTypeId=null, expiredValue=0, expiredUnit=null, cycleValue=0, cycleUnit=null, state=null)]

    }

    @Data
    @ToString(exclude = "sameMembers")
    public static class ImportMemberExcelRow extends ExcelRowRef implements ExcelRowIgnorable {
        @AsmMaxSize(12) @AsmMessage("会员姓名未填写或超过12字;") @ExcelColTitle("会员姓名")
        String memberName;
        @AsmRange("男,女") @AsmMessage("性别未正确填写;") @ExcelColTitle("性别") String sex;
        @AsmMessage("手机号码未正确填写;") @ExcelColTitle("手机号") String mobile;
        @AsmBlankable
        @ExcelColTitle("生日") String birthday;
        @AsmBlankable @ExcelColTitle("卡名称") String cardName;
        @AsmBlankable @AsmMaxSize(20) @AsmMessage("卡号格式不正确")
        @ExcelColTitle(value = "卡号", required = false) String cardNo;
        @AsmBlankable @ExcelColTitle("办卡价格") String price;
        @AsmBlankable @ExcelColTitle("消费上限") String times;
        @AsmBlankable @ExcelColTitle("总次数") String totalTimes;
        @AsmBlankable @ExcelColTitle("剩余次数") String availTimes;
        @AsmBlankable @ExcelColTitle("总金额") String totalMoney; // 储值卡总金额
        @AsmBlankable @ExcelColTitle("剩余金额") String availMoney; // 储值卡剩余金额
        @AsmBlankable @ExcelColTitle("有效期开始日") String effective;
        @AsmBlankable @ExcelColTitle("有效期截止日") String expired;
        @AsmBlankable @ExcelColTitle(value = "最迟激活日", required = false)
        String latestActivateDay;
        @AsmMaxSize(100) @AsmMessage("会员备注不得超过100个字;") @AsmBlankable
        @ExcelColTitle(value = "会员备注", required = false) String brief;
        @AsmMaxSize(12) @AsmBlankable
        @ExcelColTitle(value = "会籍", required = false) String advisor;

        @Override public boolean ignoreRow() {
            return StringUtils.startsWith(memberName, "示例-");
        }

        private @AsmIgnore
        String memberPinyin;
        private @AsmIgnore String memberInitials;
        private @AsmIgnore StringBuilder errorMessage = new StringBuilder();
        private @AsmIgnore boolean checkedPass;

        @Override
        public String error() {
            return errorMessage.length() > 0 ? errorMessage.toString() : null;
        }

        public StringBuilder appendError(String error) {
            return errorMessage.append(error);
        }

        public boolean hasError() {
            return errorMessage.length() > 0;
        }


        private @AsmIgnore boolean sameMember = false;
        private @AsmIgnore
        List<ImportMemberExcelRow> sameMembers = Lists.newArrayList(this);

        private @AsmIgnore String userId;
        private @AsmIgnore long mbrCardId;
        private @AsmIgnore String memberId;
        private @AsmIgnore String cardId;
        private @AsmIgnore String cardTypeId;
        private @AsmIgnore int expiredValue;
        private @AsmIgnore String expiredUnit;
        private @AsmIgnore int cycleValue;
        private @AsmIgnore String cycleUnit;
        private @AsmIgnore String state;

        private static String REG = "^(0|([1-9]\\d{0,9}(\\.\\d{1,2})?))$";
        public static final String LIMITLESS = "-1";
        private static final Pattern DIGITS = Pattern.compile("(\\d{2,4})[^\\d]+(\\d{1,2})[^\\d]+(\\d{1,2})");

    }
}
