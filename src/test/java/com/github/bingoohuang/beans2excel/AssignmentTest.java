package com.github.bingoohuang.beans2excel;

import com.github.bingoohuang.excel2beans.BeansToExcel;
import com.github.bingoohuang.excel2beans.ExcelToBeansUtils;
import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.google.common.collect.Lists;
import lombok.Builder;
import lombok.Data;
import lombok.val;
import org.junit.Test;

import java.io.File;
import java.util.List;


public class AssignmentTest {
    @Test
    public void test() {
        val workbook = ExcelToBeansUtils.getClassPathWorkbook("assignment.xlsx");
        val beansToExcel = new BeansToExcel(workbook);
        val name = "test-assignment.xlsx";

        createExcel(beansToExcel, name);
    }

    private void createExcel(BeansToExcel beansToExcel, String name) {
        List<ExportFollowUserExcelRow> members = Lists.newArrayList();
        members.add(ExportFollowUserExcelRow.builder().seq(1).name("aqwet").grade("B级访客").gender("女").mobile("123333333311").createTime("2017-10-11 09:16:45").sources("暂未填写").followTotalNum("0").advisorName("-").currentFollowName("-").currentFollowTime("-").build());
        members.add(ExportFollowUserExcelRow.builder().seq(2).name("wang3").grade("B级访客").gender("女").mobile("111111111114").createTime("2017-10-10 17:21:23").sources("暂未填写").followTotalNum("0").advisorName("-").currentFollowName("-").currentFollowTime("-").build());
        members.add(ExportFollowUserExcelRow.builder().seq(3).name("wang2").grade("B级访客").gender("女").mobile("111111111113").createTime("2017-10-10 17:14:14").sources("暂未填写").followTotalNum("1").advisorName("肖维维").currentFollowName("馆主").currentFollowTime("2017-10-10 17:17:06").build());
        members.add(ExportFollowUserExcelRow.builder().seq(4).name("wang1").grade("B级访客").gender("女").mobile("111111111112").createTime("2017-10-10 17:14:14").sources("暂未填写").followTotalNum("1").advisorName("Parkz").currentFollowName("馆主").currentFollowTime("2017-10-10 17:17:06").build());

        val workbook = beansToExcel.create(members);

        new File(name).delete();
        ExcelToBeansUtils.writeExcel(workbook, name);
    }

    @Data @Builder
    public static class ExportFollowUserExcelRow {
        @ExcelColTitle("序号") private int seq;
        @ExcelColTitle("客户姓名") private String name;
        @ExcelColTitle("客户类型") private String grade;
        @ExcelColTitle("性别") private String gender;
        @ExcelColTitle("手机号码") private String mobile;
        @ExcelColTitle("建档时间") private String createTime;
        @ExcelColTitle("来源渠道") private String sources;
        @ExcelColTitle("跟进总数") private String followTotalNum;
        @ExcelColTitle("当前所属会籍") private String advisorName;
        @ExcelColTitle("最近跟进人") private String currentFollowName;
        @ExcelColTitle("最近跟进时间") private String currentFollowTime;
    }
}
