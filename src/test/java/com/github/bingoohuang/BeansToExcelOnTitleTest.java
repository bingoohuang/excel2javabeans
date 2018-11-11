package com.github.bingoohuang;

import com.github.bingoohuang.beans2excel.CepingRecord;
import com.github.bingoohuang.excel2beans.BeansToExcelOnTitle;
import com.github.bingoohuang.excel2beans.ExcelToBeansUtils;
import com.github.bingoohuang.excel2beans.PoiUtil;
import com.google.common.collect.Lists;
import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.val;
import org.junit.Test;

import java.util.HashMap;
import java.util.List;

public class BeansToExcelOnTitleTest {
    @Test @SneakyThrows
    public void test1() {
        List<CepingRecord> records = Lists.newArrayList();
        @Cleanup val wb = ExcelToBeansUtils.getClassPathWorkbook("ceping.xlsx");
        val beansToExcel = new BeansToExcelOnTitle(wb.getSheet("批量导出"));

        records.add(CepingRecord.builder()
                .name("张无忌")
                .itemName("武术大师招聘")
                .source("外部")
                .details(new HashMap<String, String>() {{
                    put("积极乐观", "98");
                    put("偏执多疑", "8");
                    put("1.测试完成时间", "2018-11-02 11:03:25");
                }}).build());

        records.add(CepingRecord.builder()
                .name("牛顿")
                .itemName("物理大师招聘")
                .source("外部")
                .details(new HashMap<String, String>() {{
                    put("施展才华", "99");
                    put("数学能力", "100");
                    put("空间能力", "200");
                    put("2.测试完成时间", "2018-11-02 11:03:37");
                }}).build());

        @Cleanup val newWb = beansToExcel.create(records);

        PoiUtil.writeExcel(newWb, "批量导出-result.xlsx");
    }
}
