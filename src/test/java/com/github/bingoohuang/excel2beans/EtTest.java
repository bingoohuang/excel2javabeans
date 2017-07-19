package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.github.bingoohuang.excel2beans.annotations.ExcelSheet;
import lombok.*;
import org.junit.Test;

import java.io.File;

import static com.google.common.truth.Truth.assertThat;

public class EtTest {
    @Test @SneakyThrows
    public void member() {
        @Cleanup val workbook = ExcelToBeansUtils.getClassPathWorkbook("et.xlsx");
        val excelToBeans = new ExcelToBeans(workbook);
        val etMembers = excelToBeans.convert(EtMember.class);
        val member1 = EtMember.builder().memberName("黄同学").teachers("林老师")
                .memberCard("零基础素描课程").totalTimes(25)
                .mobile(CellData.builder().comment("手机号码不能重复哦").commentAuthor("Microsoft Office 用户")
                        .row(2).col(3).value("12345678910").build()).build();
        val member2 = EtMember.builder().memberName("陈学").teachers("刘老师")
                .mobile(CellData.builder().comment(null).commentAuthor(null)
                        .row(3).col(3).value("12345678918").build())
                .memberCard("零基础素描课程").totalTimes(25).build();
        val member3 = EtMember.builder().memberName("齐学").teachers("刘老师/华老师")
                .mobile(CellData.builder().comment(null).commentAuthor(null)
                        .row(4).col(3).value("12345678919").build())
                .memberCard("零基础油画系统课程").totalTimes(40).build();
        val member4 = EtMember.builder().memberName("徐同学").teachers("刘老师")
                .mobile(CellData.builder().comment(null).commentAuthor(null)
                        .row(5).col(3).value("12345678920").build())
                .memberCard("零基础素描课程").totalTimes(25).build();
        val member5 = EtMember.builder().memberName("宋同学").teachers("刘老师")
                .mobile(CellData.builder().comment(null).commentAuthor(null)
                        .row(6).col(3).value("12345678921").build())
                .memberCard("零基础素描课程").totalTimes(25).build();
        assertThat(etMembers).containsExactly(member1, member2, member3, member4, member5);

        val mobile = member1.getMobile();
        mobile.setComment("号码重复");
        mobile.setCommentAuthor("et-server");

        ExcelToBeansUtils.writeRedComments(workbook, mobile);
        ExcelToBeansUtils.writeExcel(workbook, "test-et-out.xlsx");
        new File("test-et-out.xlsx").delete();

        val etCards = excelToBeans.convert(EtCard.class);
        val card1 = EtCard.builder().cardName("零基础素描课程").times(25).expiredValue(12).expiredUnit("月").salePrice(2580).courses("25节素描课程").build();
        val card2 = EtCard.builder().cardName("零基础油画课程").times(20).expiredValue(12).expiredUnit("月").salePrice(3380).courses("20节油画课程").build();
        val card3 = EtCard.builder().cardName("零基础精品油画速成班").times(15).expiredValue(12).expiredUnit("月").salePrice(2980).courses("5节素描课程+10节油画课程").build();
        val card4 = EtCard.builder().cardName("零基础油画系统课程").times(40).expiredValue(24).expiredUnit("月").salePrice(4980).courses("20节素描课程+20节油画课程").build();
        val card5 = EtCard.builder().cardName("零基础彩铅课程").times(20).expiredValue(12).expiredUnit("月").salePrice(2680).courses("10节素描课程+10节彩铅课程").build();
        assertThat(etCards).containsExactly(card1, card2, card3, card5, card4);
    }

    @Data @Builder @ExcelSheet(name = "学员")
    public static class EtMember {
        @ExcelColTitle("学员姓名") private String memberName;
        @ExcelColTitle("责任教师") private String teachers;
        @ExcelColTitle("手机号") private CellData mobile;
        @ExcelColTitle("课程卡") private String memberCard;
        @ExcelColTitle("总次数") private int totalTimes;
    }

    @Data @Builder @ExcelSheet(name = "卡模板")
    public static class EtCard {
        @ExcelColTitle("卡模板名称") private String cardName;
        @ExcelColTitle("次数") private int times;
        @ExcelColTitle("有效期值") private int expiredValue;
        @ExcelColTitle("有效期单位") private String expiredUnit;
        @ExcelColTitle("标准售价") private int salePrice;
        @ExcelColTitle("所含课程") private String courses;
    }
}
