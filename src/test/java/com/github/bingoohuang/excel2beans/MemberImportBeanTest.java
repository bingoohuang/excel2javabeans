package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.asmvalidator.AsmValidateResult;
import com.github.bingoohuang.asmvalidator.AsmValidatorFactory;
import lombok.val;
import org.junit.Test;

import java.util.List;

import static com.github.bingoohuang.excel2beans.ExcelToBeansUtils.getClassPathWorkbook;
import static com.google.common.truth.Truth.assertThat;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
public class MemberImportBeanTest {
    @Test public void testWithoutBlankHeadRowsAndCols() {
        val workbook = getClassPathWorkbook("member.xlsx");
        val excelToBeans = new ExcelToBeans(MemberImportBean.class);
        List<MemberImportBean> beans = excelToBeans.convert(workbook);
        assertThat(beans).hasSize(4);

        assertThat(beans.get(0).getRowNum()).isEqualTo(6);
        assertThat(beans.get(1).getRowNum()).isEqualTo(7);
        assertThat(beans.get(2).getRowNum()).isEqualTo(8);
        assertThat(beans.get(3).getRowNum()).isEqualTo(9);

        AsmValidateResult result = new AsmValidateResult();
        AsmValidatorFactory.validateAll(beans, result);
        System.out.println(result);
    }
}
