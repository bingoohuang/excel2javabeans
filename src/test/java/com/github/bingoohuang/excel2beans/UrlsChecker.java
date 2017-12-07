package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.asmvalidator.AsmValidateResult;
import com.github.bingoohuang.asmvalidator.MsaValidator;
import com.github.bingoohuang.asmvalidator.ValidateError;
import com.github.bingoohuang.asmvalidator.annotations.AsmConstraint;
import lombok.val;
import org.apache.commons.lang3.StringUtils;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;
import java.util.List;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

@AsmConstraint(supportedClasses = List.class, validateBy = UrlsChecker.AsmUrlsValidator.class)
@Target(FIELD)
@Retention(RUNTIME)
public @interface UrlsChecker {
    class AsmUrlsValidator implements MsaValidator<UrlsChecker, List<String>> {
        @Override public void validate(UrlsChecker annotation, AsmValidateResult result, List<String> urls) {
            // 手工校验其它字段
            for (int i = 0, ii = urls.size(); i < ii; ++i) {
                val url = urls.get(i);
                if (StringUtils.isEmpty(url)) {
                    result.addError(new ValidateError("playUrls_" + i, url, "URL不能为空"));
                } else if (url.length() > 2) {
                    result.addError(new ValidateError("playUrls_" + i, url, "URL长度不能超过2"));
                }
            }

        }
    }
}
