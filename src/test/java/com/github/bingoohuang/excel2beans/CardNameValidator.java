package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.asmvalidator.AsmValidateResult;
import com.github.bingoohuang.asmvalidator.MsaValidator;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/11.
 */
public class CardNameValidator implements MsaValidator<MemberCardName, String> {
    @Override
    public void validate(MemberCardName memberCardNameAnn, AsmValidateResult result, String cardName) {
        System.out.println(cardName);
    }
}
