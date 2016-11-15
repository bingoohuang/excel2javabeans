package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.asmvalidator.annotations.*;
import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import lombok.Builder;
import lombok.Data;
import org.apache.commons.lang3.StringUtils;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
@Data @Builder
public class MemberImportBeanWithBlankHeadRowsAndCols extends ExcelRowRef implements ExcelRowIgnorable {
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
