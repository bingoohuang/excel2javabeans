package com.github.bingoohuang.beans2excel;

import com.github.bingoohuang.beans2excel.beans.Member;
import com.github.bingoohuang.beans2excel.beans.Schedule;
import com.github.bingoohuang.beans2excel.beans.Subscribe;
import com.github.bingoohuang.excel2beans.BeansToExcel;
import com.github.bingoohuang.excel2beans.ExcelToBeansUtils;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.ss.usermodel.Workbook;
import org.joda.time.DateTime;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.sql.Timestamp;
import java.util.List;
import java.util.Map;

/**
 * Created by bingoohuang on 2017/3/20.
 */
public class SimpleTest {
    @Test
    public void testHead() {
        val beansToExcel = new BeansToExcel();
        String name = "test-workbook-head.xlsx";

        List<Member> members = Lists.newArrayList();
        members.add(new Member(1000, 100, 80));

        List<Schedule> schedules = Lists.newArrayList();
        schedules.add(new Schedule(Timestamp.valueOf("2007-11-11 12:13:14"), 100, 200, 90, 10));
        schedules.add(new Schedule(Timestamp.valueOf("2007-01-11 12:13:14"), 101, 201, 91, 11));


        Map<String, Object> props = Maps.newHashMap();
        // 增加头行信息
        props.put("memberHead", "会员信息" + DateTime.now().toString("yyyy-MM-dd"));
        val workbook = beansToExcel.create(props, members, schedules);

        writeExcel(name, workbook);
    }

    @SneakyThrows
    private void writeExcel(String name, Workbook workbook) {
        @Cleanup val fileOut = new FileOutputStream(name);
        workbook.write(fileOut);
    }

    @Test
    public void test() {
        val beansToExcel = new BeansToExcel();
        String name = "test-workbook.xlsx";

        createExcel(beansToExcel, name);
    }

    @Test
    public void testTemplate() {
        val workbook = ExcelToBeansUtils.getClassPathWorkbook("template.xlsx");
        val beansToExcel = new BeansToExcel(workbook);
        String name = "test-workbook-templ.xlsx";

        createExcel(beansToExcel, name);
    }


    private void createExcel(BeansToExcel beansToExcel, String name) {
        List<Member> members = Lists.newArrayList();
        members.add(new Member(1000, 100, 80));

        List<Schedule> schedules = Lists.newArrayList();
        schedules.add(new Schedule(Timestamp.valueOf("2007-11-11 12:13:14"), 100, 200, 90, 10));
        schedules.add(new Schedule(Timestamp.valueOf("2007-01-11 12:13:14"), 101, 201, 91, 11));

        List<Subscribe> subscribes = Lists.newArrayList();
        subscribes.add(new Subscribe(Timestamp.valueOf("2007-01-11 12:13:14"), 100, 10));
        subscribes.add(new Subscribe(Timestamp.valueOf("2007-02-11 12:13:14"), 101, 11));
        subscribes.add(new Subscribe(Timestamp.valueOf("2007-03-11 12:13:14"), 102, 12));

        val workbook = beansToExcel.create(members, schedules, subscribes);

        writeExcel(name, workbook);
    }

    public static InputStream convert(ByteArrayOutputStream out) {
        return new ByteArrayInputStream(out.toByteArray());
    }
}
