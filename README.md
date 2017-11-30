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
@Data @Builder
public class ExportFollowUserExcelRow {
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

val styleTemplate = ExcelToBeansUtils.getClassPathWorkbook("assignment.xlsx");
val beansToExcel = new BeansToExcel(styleTemplate);
List<ExportFollowUserExcelRow> members = Lists.newArrayList();
members.add(...);
members.add(...);
members.add(...);
members.add(...);

val workbook = beansToExcel.create(template);

ExcelToBeansUtils.writeExcel(workbook, name);
```

![image](https://user-images.githubusercontent.com/1940588/33408898-d26086ce-d5b3-11e7-9431-c48ccf6799aa.png)

# Cell Image Support
Now the image in excel can be bound to bean field of type ImageData.
The image's axis will be computed to match the related cell. 
![image](https://user-images.githubusercontent.com/1940588/33408499-c5489ba4-d5b1-11e7-86ee-10913dd1eaef.png)


```java
@Data
public class ImageBean {
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
mvn clean install -DskipTests -Dgpg.passphrase=slgsdmxl
mvn clean install -Dgpg.skip -DskipTests
```
