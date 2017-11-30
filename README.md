# excel2javabeans
convert excel rows to javabeans and vice visa.
<br/>
[![Build Status](https://travis-ci.org/bingoohuang/excel2javabeans.svg?branch=master)](https://travis-ci.org/bingoohuang/excel2javabeans)
[![Quality Gate](https://sonarqube.com/api/badges/gate?key=com.github.bingoohuang%3Aexcel2javabeans)](https://sonarqube.com/dashboard/index/com.github.bingoohuang%3Aexcel2javabeans)
[![Coverage Status](https://coveralls.io/repos/github/bingoohuang/excel2javabeans/badge.svg?branch=master)](https://coveralls.io/github/bingoohuang/excel2javabeans?branch=master)
[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.github.bingoohuang/excel2javabeans/badge.svg?style=flat-square)](https://maven-badges.herokuapp.com/maven-central/com.github.bingoohuang/excel2javabeans/)
[![License](http://img.shields.io/:license-apache-brightgreen.svg)](http://www.apache.org/licenses/LICENSE-2.0.html)


## Convert Excel to Javabeans
![image](https://user-images.githubusercontent.com/1940588/33408746-0213ccf6-d5b3-11e7-9f89-1c6cf08322bf.png)

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
Map<String, Object> props = Maps.newHashMap();
// 增加头行信息(如果有的话)
props.put("memberHead", "会员信息" + DateTime.now().toString("yyyy-MM-dd"));
Workbook workbook = beansToExcel.create(props, members, schedules, subscribes);
```

```java
@ExcelSheet(name = "会员", headKey = "memberHead")
public class Member {
    @ExcelColTitle("会员总数")
    private int total;
    @ExcelColTitle("其中：新增")
    @ExcelColStyle(align = CENTER)
    private int fresh;
    @ExcelColTitle("其中：有效")
    @ExcelColStyle(align = CENTER)
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

# Cell Image Support
Now the image in excel can be bound to bean field of type ImageData.
The image's axis will be computed to match the related cell. 
![image](https://user-images.githubusercontent.com/1940588/33408499-c5489ba4-d5b1-11e7-86ee-10913dd1eaef.png)


```java
@Data
public static class ImageBean {
    @ExcelColTitle("图片")
    private ImageData imageData;
    @ExcelColTitle("名字")
    private String name;
}

public void testImage() {
    @Cleanup val workbook = ExcelToBeansUtils.getClassPathWorkbook("images.xls");
    val excelToBeans = new ExcelToBeans(workbook);
    val beans = excelToBeans.convert(ImageBean.class);
}
```


# Excel Download Utility
```java
// HttpServletResponse response = ...
// Workbook workbook = ...
ExcelToBeansUtils.download(response,  workbook, "fileName.xlsx");
```

# Sonarqube
```bash
travis encrypt a7fe683637d6e1f54e194817cc36e78936d4fe61

mvn clean install sonar:sonar -Dsonar.organization=bingoohuang-github -Dsonar.host.url=https://sonarqube.com -Dsonar.login=a7fe683637d6e1f54e194817cc36e78936d4fe61
```

# Problems
## Autosize column does not work on CentOS.
Maybe there is not relative fonts installed. Methods: 
1. Create fonts folder:`mkdir ~/.fonts` 
2. Copy fonts to the fold:`scp /System/Library/Fonts/STHeiti\ Light.ttc yogaapp@test.ino01:./.fonts/`
3. Install the fonts:`fc-cache -f -v`

For all users available, just copy the fonts file to the `/usr/share/fonts` directory and then `fc-cache -f -v`.

# gpg
```bash
GPG_TTY=$(tty)
export GPG_TTY
```

```fish
set -gx GPG_TTY (tty)
```

```bash
mvn clean install -DskipTests  -Dgpg.passphrase=slgsdmxl
```
