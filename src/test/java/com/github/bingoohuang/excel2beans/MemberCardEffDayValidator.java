package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.asmvalidator.AsmValidateResult;
import com.github.bingoohuang.asmvalidator.MsaValidator;
import com.github.bingoohuang.asmvalidator.ValidateError;
import org.joda.time.format.DateTimeFormat;
import org.joda.time.format.DateTimeFormatter;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/11.
 */
public class MemberCardEffDayValidator implements MsaValidator<MemberCardEffDay, String> {
    DateTimeFormatter dateTimeFormatter = DateTimeFormat.forPattern("yyyy-MM-dd");

    @Override
    public void validate(MemberCardEffDay memberCardEffDay,
                         AsmValidateResult result, String effectiveDay) {
        if (effectiveDay.equals("无限期")) return;

        try {
            dateTimeFormatter.parseDateTime(effectiveDay);
        } catch (IllegalArgumentException e) {
            result.addError(new ValidateError("有效期", effectiveDay,
                    "有效期日期格式无法解析成yyyy-MM-dd"));
        }
    }
}
