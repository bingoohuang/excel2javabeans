package com.github.bingoohuang.excel2javabeans;

import com.github.bingoohuang.asmvalidator.annotations.*;
import com.github.bingoohuang.excel2javabeans.annotations.ExcelColumnTitleContains;
import lombok.Data;
import org.apache.commons.lang3.StringUtils;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
@Data
public class MemberImportBean extends ExcelRowReference implements ExcelRowIgnore {
    @ExcelColumnTitleContains("会员姓名") @AsmMaxSize(12) String memberName; // 不超过12字
    @ExcelColumnTitleContains("性别") @AsmRange("男,女") String sex;
    @ExcelColumnTitleContains("手机号") @AsmMobile String mobile;
    @ExcelColumnTitleContains("生日") @AsmBlankable String birthday;
    @ExcelColumnTitleContains("卡名称") String cardName;
    @ExcelColumnTitleContains("办卡价格") @AsmBlankable @AsmDigits String cardPrice;
    @ExcelColumnTitleContains("消费上限") @AsmBlankable @AsmDigits String upperTimes;
    @ExcelColumnTitleContains("总次数") @AsmBlankable @AsmDigits String totalTimes;
    @ExcelColumnTitleContains("剩余次数") @AsmBlankable @AsmDigits String availableTimes;
    @ExcelColumnTitleContains("有效期开始日") String effectiveTime;
    @ExcelColumnTitleContains("有效期截止日") String expiredTime;

    @Override public boolean ignoreRow() {
        return StringUtils.startsWith(memberName, "示例-");
    }
}
