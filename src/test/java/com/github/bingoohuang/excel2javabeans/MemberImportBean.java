package com.github.bingoohuang.excel2javabeans;

import com.github.bingoohuang.asmvalidator.annotations.*;
import com.github.bingoohuang.excel2javabeans.annotations.ExcelColTitle;
import lombok.Data;
import org.apache.commons.lang3.StringUtils;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
@Data
public class MemberImportBean extends ExcelRowRef implements ExcelRowIgnorable {
    @ExcelColTitle("会员姓名") @AsmMaxSize(12) @AsmMessage("请填写会员姓名") String memberName; // 不超过12字
    @ExcelColTitle("性别") @AsmRange("男,女") @AsmMessage("性别请填男或女") String sex;
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
}
