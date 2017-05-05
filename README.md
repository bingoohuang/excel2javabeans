# excel2javabeans
convert excel rows to javabeans
<br/>
[![Build Status](https://travis-ci.org/bingoohuang/excel2javabeans.svg?branch=master)](https://travis-ci.org/bingoohuang/excel2javabeans)
[![Coverage Status](https://coveralls.io/repos/github/bingoohuang/excel2javabeans/badge.svg?branch=master)](https://coveralls.io/github/bingoohuang/excel2javabeans?branch=master)
[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.github.bingoohuang/excel2javabeans/badge.svg?style=flat-square)](https://maven-badges.herokuapp.com/maven-central/com.github.bingoohuang/excel2javabeans/)
[![License](http://img.shields.io/:license-apache-brightgreen.svg)](http://www.apache.org/licenses/LICENSE-2.0.html)


## Convert Excel to Javabeans

```java
// ... 
ExcelToBeans excelToBeans = new ExcelToBeans(BeanWithTitle.class);
List<BeanWithTitle> beans = excelToBeans.convert(workbook);
// ...
```

```java
public class BeanWithTitle extends ExcelRowRef implements ExcelRowIgnorable {
    @ExcelColTitle("会员姓名") String memberName;
    @ExcelColTitle("卡名称") String cardName;
    @ExcelColTitle("办卡价格") String cardPrice;
    @ExcelColTitle("性别") String sex;

    @Override public boolean ignoreRow() {
        return StringUtils.startsWith(memberName, "示例-");
    }
    
    // getters and setters ignored
}
```

## Convert Javabeans to Excel
```java
BeansToExcel beansToExcel = new BeansToExcel();
List<Member> members = // create members
List<Schedule> schedules = // create schdules
List<Subscribe> subscribes = // create subcribes
Workbook workbook = beansToExcel.create(members, schedules, subscribes);
```

```java
@ExcelSheet(name = "会员")
public class Member {
    @ExcelColTitle("会员总数")
    private int total;
    @ExcelColTitle("其中：新增")
    private int fresh;
    @ExcelColTitle("其中：有效")
    private int effective;
    // getters and setters ignored
}

@ExcelSheet(name = "排期")
public class Schedule {
    @ExcelColTitle("日期")
    private Timestamp time;
    @ExcelColTitle("排期数")
    private int schedules;
    @ExcelColTitle("定课数")
    private int subscribes;
    @ExcelColTitle("其中：小班课")
    private int publics;
    @ExcelColTitle("其中：私教课")
    private int privates;
    // getters and setters ignored
}

@ExcelSheet(name = "订课情况")
public class Subscribe {
    @ExcelColTitle("订单日期")
    private Timestamp day;
    @ExcelColTitle("人次")
    private int times;
    @ExcelColTitle("人数")
    private int heads;
    // getters and setters ignored
}

```


# Sonarqube
```
New token "excel2javabeans" has been created. Make sure you copy it now, you won’t be able to see it again!
a7fe683637d6e1f54e194817cc36e78936d4fe61

travis encrypt a7fe683637d6e1f54e194817cc36e78936d4fe61

mvn clean install sonar:sonar -Dsonar.organization=bingoohuang-github -Dsonar.host.url=https://sonarqube.com -Dsonar.login=a7fe683637d6e1f54e194817cc36e78936d4fe61
```