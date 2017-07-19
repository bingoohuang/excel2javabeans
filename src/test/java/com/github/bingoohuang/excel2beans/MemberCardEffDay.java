package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.asmvalidator.AsmValidateResult;
import com.github.bingoohuang.asmvalidator.MsaValidator;
import com.github.bingoohuang.asmvalidator.ValidateError;
import com.github.bingoohuang.asmvalidator.annotations.AsmConstraint;
import org.joda.time.format.DateTimeFormat;
import org.joda.time.format.DateTimeFormatter;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;


@AsmConstraint(validateBy = MemberCardEffDay.MemberCardEffDayValidator.class)
@Target(FIELD)
@Retention(RUNTIME)
public @interface MemberCardEffDay {
    class MemberCardEffDayValidator implements MsaValidator<MemberCardEffDay, String> {
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
}
