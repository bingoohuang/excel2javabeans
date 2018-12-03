package com.github.bingoohuang.beans2excel;

import com.github.bingoohuang.excel2beans.BeansToExcelOnTitle;
import com.github.bingoohuang.excel2beans.PoiUtil;
import com.google.common.collect.Lists;
import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.val;
import org.junit.Test;

import java.io.File;

public class BeansToExcelOnTitleTest {
    @Test @SneakyThrows
    public void test1() {
        val records = Lists.<CepingRecord>newArrayList(
                CepingRecord.builder()
                        .name("张无忌")
                        .itemName("武术大师招聘")
                        .source("外部")
                        .detail("积极乐观", "98")
                        .detail("偏执多疑", "8")
                        .detail("1.测试完成时间", "2018-11-02 11:03:25")
                        .build(),

                CepingRecord.builder()
                        .name("牛顿")
                        .itemName("物理大师招聘")
                        .source("外部")
                        .detail("施展才华", "99")
                        .detail("数学能力", "100")
                        .detail("空间能力", "200")
                        .detail("2.测试完成时间", "2018-11-02 11:03:37")
                        .build()
        );

        val beansToExcel = new BeansToExcelOnTitle("ceping.xlsx", CepingRecord.class);
        @Cleanup val newWb = beansToExcel.create(records);

        val file = new File("批量导出-result.xlsx");
        PoiUtil.writeExcel(newWb, file);
//        file.deleteOnExit();
    }
}
