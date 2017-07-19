package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.asmvalidator.AsmValidateResult;
import com.github.bingoohuang.asmvalidator.MsaValidator;
import com.github.bingoohuang.asmvalidator.annotations.AsmConstraint;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/11.
 */
@AsmConstraint(validateBy = MemberCardName.CardNameValidator.class)
@Target(FIELD)
@Retention(RUNTIME)
public @interface MemberCardName {
    /**
     * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/11.
     */
    class CardNameValidator implements MsaValidator<MemberCardName, String> {
        @Override
        public void validate(MemberCardName memberCardNameAnn, AsmValidateResult result, String cardName) {
            System.out.println(cardName);
        }
    }
}
