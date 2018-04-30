package com.github.bingoohuang.beans2excel;

import com.github.bingoohuang.excel2beans.BeansToExcel;
import com.github.bingoohuang.excel2beans.ExcelToBeans;
import com.github.bingoohuang.excel2beans.ExcelToBeansUtils;
import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.google.common.collect.Lists;
import lombok.*;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;

import static com.google.common.truth.Truth.assertThat;

public class EmojiTest {
    @Test @SneakyThrows
    public void testWriteEmoji() {
        val wxNick = new WxNick("🦄女侠🌈💄💓", "🎈Nancy🍬");
        val wxNicks = Lists.newArrayList(wxNick);
        @Cleanup val workbook = new BeansToExcel().create(wxNicks);

        String fileName = "test-emoji-out.xlsx";
        ExcelToBeansUtils.writeExcel(workbook, fileName);

        @Cleanup val fis = new FileInputStream(fileName);
        val wb = WorkbookFactory.create(fis);
        val beans = new ExcelToBeans(workbook).convert(WxNick.class);
        assertThat(beans).containsExactly(wxNick);

        new File(fileName).delete();
    }

    @Test @SneakyThrows
    public void testReadEmoji() {
        @Cleanup
        val workbook = ExcelToBeansUtils.getClassPathWorkbook("emoji.xlsx");
        val excelToBeans = new ExcelToBeans(workbook);
        val beans = excelToBeans.convert(WxNick.class);

        assertThat(beans).containsExactly(
                new WxNick("春秋小鱼", "自然疯 ❤")
                ,new WxNick("🌹禾🚼🌹", "天天老师")
                ,new WxNick("大(^o^)丹丹", "yuanyuanji")
                ,new WxNick("蔚蓝的天空", "💎小卟点")
                ,new WxNick("🎈Nancy🍬", "🎈Nancy🍬")
                ,new WxNick("🦄女侠🌈💄💓", "🎈Nancy🍬")
                ,new WxNick("🍭卢小贯", "杨洋老师")
                ,new WxNick("金娃娃👧", "何老师")
                ,new WxNick("🐻维尼熊之笨笨🐳", "杨洋老师")
                );
    }

    @Data @AllArgsConstructor
    public static class WxNick {
        @ExcelColTitle("购买人微信昵称")
        private String referrerWxNick;
        @ExcelColTitle("推荐人微信昵称")
        private String buyerWxNick;
    }
}
