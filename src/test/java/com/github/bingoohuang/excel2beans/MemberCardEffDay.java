package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.asmvalidator.annotations.AsmConstraint;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/11.
 */
@AsmConstraint(validateBy = MemberCardEffDayValidator.class)
@Target(FIELD)
@Retention(RUNTIME)
public @interface MemberCardEffDay {
}
