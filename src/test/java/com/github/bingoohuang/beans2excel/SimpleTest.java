package com.github.bingoohuang.beans2excel;

import com.github.bingoohuang.beans2excel.beans.Member;
import com.github.bingoohuang.beans2excel.beans.Schedule;
import com.github.bingoohuang.beans2excel.beans.Subscribe;
import com.github.bingoohuang.excel2beans.BeansToExcel;
import com.google.common.collect.Lists;
import lombok.Cleanup;
import lombok.SneakyThrows;
import lombok.val;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.FileOutputStream;
import java.sql.Timestamp;
import java.util.List;

/**
 * Created by bingoohuang on 2017/3/20.
 */
public class SimpleTest {
    @Test
    @SneakyThrows
    public void test() {
        BeansToExcel beansToExcel = new BeansToExcel();
        List<Member> members = Lists.newArrayList();
        members.add(new Member(1000, 100, 80));

        List<Schedule> schedules = Lists.newArrayList();
        schedules.add(new Schedule(Timestamp.valueOf("2007-11-11 12:13:14"), 100, 200, 90, 10));
        schedules.add(new Schedule(Timestamp.valueOf("2007-01-11 12:13:14"), 101, 201, 91, 11));

        List<Subscribe> subscribes = Lists.newArrayList();
        subscribes.add(new Subscribe(Timestamp.valueOf("2007-01-11 12:13:14"), 100, 10));
        subscribes.add(new Subscribe(Timestamp.valueOf("2007-02-11 12:13:14"), 101, 11));
        subscribes.add(new Subscribe(Timestamp.valueOf("2007-03-11 12:13:14"), 102, 12));

        Workbook workbook = beansToExcel.create(members, schedules, subscribes);

        @Cleanup val fileOut = new FileOutputStream("test-workbook.xlsx");
        workbook.write(fileOut);
    }
}
