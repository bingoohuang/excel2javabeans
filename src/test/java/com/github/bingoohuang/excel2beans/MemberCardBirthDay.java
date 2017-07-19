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

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/11.
 */
@AsmConstraint(validateBy = MemberCardBirthDay.MemberCardDayValidator.class)
@Target(FIELD)
@Retention(RUNTIME)
public @interface MemberCardBirthDay {
    /**
     * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/11.
     */
    class MemberCardDayValidator implements MsaValidator<MemberCardBirthDay, String> {
        DateTimeFormatter dateTimeFormatter = DateTimeFormat.forPattern("yyyy-MM-dd");

        @Override
        public void validate(MemberCardBirthDay memberCardBirthDay, AsmValidateResult result, String birthday) {
            try {
                dateTimeFormatter.parseDateTime(birthday);
            } catch (IllegalArgumentException e) {
                result.addError(new ValidateError("birthday", birthday,
                        "日期格式无法解析成yyyy-MM-dd"));
            }
        }
    }
}
