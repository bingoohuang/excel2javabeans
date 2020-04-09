# excel2javabeans
convert excel rows to javabeans and vice visa.
<br/>
[![Build Status](https://travis-ci.org/bingoohuang/excel2javabeans.svg?branch=master)](https://travis-ci.org/bingoohuang/excel2javabeans)
[![Quality Gate](https://sonarcloud.io/api/project_badges/measure?project=com.github.bingoohuang%3Aexcel2javabeans&metric=alert_status)](https://sonarcloud.io/dashboard/index/com.github.bingoohuang%3Aexcel2javabeans)
[![Coverage Status](https://coveralls.io/repos/github/bingoohuang/excel2javabeans/badge.svg?branch=master)](https://coveralls.io/github/bingoohuang/excel2javabeans?branch=master)
[![Maven Central](https://maven-badges.herokuapp.com/maven-central/com.github.bingoohuang/excel2javabeans/badge.svg?style=flat-square)](https://maven-badges.herokuapp.com/maven-central/com.github.bingoohuang/excel2javabeans/)
[![License](http://img.shields.io/:license-apache-brightgreen.svg)](http://www.apache.org/licenses/LICENSE-2.0.html)


## Convert Excel to Javabeans
![image](https://user-images.githubusercontent.com/1940588/33408746-0213ccf6-d5b3-11e7-9f89-1c6cf08322bf.png)

```java
// ... 
Workbook workbook = getClassPathWorkbook("member.xlsx");
ExcelToBeans excelToBeans = new ExcelToBeans(workbook);
List<BeanWithTitle> beans = excelToBeans.convert(BeanWithTitle.class);
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

Workbook templateWorkbook = ExcelToBeansUtils.getClassPathWorkbook("assignment.xlsx");
BeansToExcel beansToExcel = new BeansToExcel(templateWorkbook);
List<ExportFollowUserExcelRow> members = Lists.newArrayList();
members.add(...);
members.add(...);
members.add(...);
members.add(...);

Workbook workbook = beansToExcel.create(members);
ExcelToBeansUtils.writeExcel(workbook, name);
```

![image](https://user-images.githubusercontent.com/1940588/33408898-d26086ce-d5b3-11e7-9431-c48ccf6799aa.png)

# Cell Image Support

Now the image in excel can be bound to bean field of type ImageData.
The image's axis will be computed to match the related cell. 
![image](https://user-images.githubusercontent.com/1940588/33585908-ab2809aa-d9a1-11e7-962e-ce7c142faf99.png)


```java
@Data
public class ImageBean {
    @ExcelColTitle("图片")
    private ImageData imageData;
    @ExcelColTitle("名字")
    private String name;
}

public void testImage() {
    Workbook val workbook = ExcelToBeansUtils.getClassPathWorkbook("images.xls");
    ExcelToBeans excelToBeans = new ExcelToBeans(workbook);
    List<ImageBean> beans = excelToBeans.convert(ImageBean.class);
}
```

# List<String/Integer> bean fields support

![image](https://user-images.githubusercontent.com/1940588/33585728-afbdced8-d9a0-11e7-8903-e172fafbf577.png)

```java
@Data
public static class MultipleColumnsBeanWithTitle {
    @ExcelColTitle("会员姓名") String memberName; // for the first row, the value will be "张小凡"
    @ExcelColTitle("手机号") List<String> mobiles; // for the first row，the values will be: null, "18795952311", "18795952311", "18795952311"
    @ExcelColTitle("归属地") List<String> homeareas; // for the first row, the values will be: "南京", "北京", "上海", "广东"
}
```

# Excel SpringMVC upload and download demo
```java

/**
 * 从EXCEL中批量导入会员。
 */
@RequestMapping("/ImportMembers") @RestController
public class ImportMembersController {
    /**
     * 下载失败条目的EXCEL。
     *
     * @return RestResp
     */
    @RequestMapping("/downloadError") @SneakyThrows
    public RestResp downloadError(HttpServletResponse response) {
        byte[] workbook = ImportMembersHelper.redisExcel4ImportMemberError();
        if (workbook == null) {
            return RestResp.ok("当前没有失败条目");
        }

        ExcelDownloads.download(response, workbook, "导入错误" + WestId.next() + ".xlsx");
        return RestResp.ok("失败条目下载成功");
    }

    /**
     * 使用EXCEL 批量导入学员。
     *
     * @param file EXCEL文件
     * @return RestResp
     */
    @RequestMapping("/importMembers") @SneakyThrows
    public RestResp importMembers(@RequestParam("file") MultipartFile file) {
        @Cleanup val excelToBeans = new ExcelToBeans(file.getInputStream());
        val importedMembers = excelToBeans.convert(ImportedMember.class);
        // ...
    }

}
```

# BeansToExcelOnTemplate
![image](https://user-images.githubusercontent.com/1940588/48628405-003aa180-e9f2-11e8-83d7-429df9157a38.png)


```java
@Data @Builder
public class CepingResult {
    @ExcelCell(sheetName = true)
    private String sheetName;       // 表单名称

    @ExcelCell(value = "A2", replace = "XX")
    private String interviewCode;   // 面试编号

    @ExcelCell
    private String name;           // 身份证姓名
    @ExcelCell
    private String gender;         // 性别
    @ExcelCell
    private String age;            // 年龄

    @ExcelCell("B4")
    private String position;       // 应聘职位
    @ExcelCell("E4")
    private String level;          // 推荐职级
    @ExcelCell("G4")
    private String annualSalary;   // 期望年薪

    @ExcelCell("C5")
    private double matchScore;     // 岗位匹配度
    @ExcelCell(value = "C6", maxLineLen = 40)
    private String matchComment;   // 岗位匹配度评语
}

@Cleanup val wb = ExcelToBeansUtils.getClassPathWorkbook("template.xlsx");
val beansToExcel = new BeansToExcelOnTemplate(wb.getSheet("templateName"));

@Cleanup val newWb = beansToExcel.create(bean);
PoiUtil.protectWorkbook(newWb, "123456");
PoiUtil.writeExcel(newWb, "exported.xlsx");
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
4. 查看字体: `fc-list|grep SimSun`

```bash
[betaoper@beta-hetong ~]$ fc-list|grep SimSun
/usr/share/fonts/winfonts/simsun.ttc: 宋体,SimSun:style=常规,Regular
/usr/share/fonts/winfonts/simsun.ttc: 新宋体,NSimSun:style=常规,Regular
/usr/share/fonts/winfonts/simsunb.ttf: SimSun\-ExtB:style=Regular,obyčejné
```



For all users available, just copy the fonts file to the `/usr/share/fonts` directory and then `fc-cache -f -v`.

## Emoji output error
When writting emoji like 🦄女侠🌈💄💓 , the output excel content will show like ?女侠???, try to fix this with following dependency.
```xml
<dependency>
    <groupId>org.apache.xmlbeans</groupId>
    <artifactId>xmlbeans</artifactId>
    <version>3.0.0</version>
</dependency>
```

## How to manually test excel downloads in Chrome console?
```javascript
// execute the following javascript code in the console to download excel for testing
var fileName = "abc.xlsx"
var url = 'http://localhost:8090/yoga-system/MemberExportController/memberExport'
var request = new XMLHttpRequest()
request.open('POST', url, true)
request.setRequestHeader('Content-Type', 'application/json; charset=utf-8')
request.responseType = 'blob'
request.onload = function(e) {
    if (this.status === 200) {
        var blob = this.response;
        if(window.navigator.msSaveOrOpenBlob) {
            window.navigator.msSaveBlob(blob, fileName)
        } else {
            var downloadLink = window.document.createElement('a')
            var contentTypeHeader = request.getResponseHeader("Content-Type")
            downloadLink.href = window.URL.createObjectURL(new Blob([blob], {type: contentTypeHeader}))
            downloadLink.download = fileName
            document.body.appendChild(downloadLink)
            downloadLink.click()
            document.body.removeChild(downloadLink)
       }
   }
}
request.send(JSON.stringify({firstBlood:false,export:true}));
```
## Ajax js
```javascript
export const downloadBlobFile = (response) => {
  const blob = new Blob([response.data], {type: response.data.type})
  const contentDisposition = response.headers['content-disposition']
  let fileName = 'unknown'
  if (contentDisposition) {
    const fileNameMatch = contentDisposition.match(/filename="(.+)"/)
    if (fileNameMatch.length === 2) {
      fileName = decodeURIComponent(fileNameMatch[1])
    }
  }

  if (window.navigator.msSaveOrOpenBlob) {
    navigator.msSaveBlob(blob, fileName)
  } else {
    const link = document.createElement('a')
    link.href = window.URL.createObjectURL(blob)
    link.download = fileName
    link.click()
    window.URL.revokeObjectURL(link.href)
  }
}

import {downloadBlobFile} from 'utils'
this.$http.post('/BasicParametersController/exportPositions/', {}, {
  responseType: 'blob'
}).then((response) => {
  downloadBlobFile(response)
})

```

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

# TODO
1. Support SXSSF (Streaming Usermodel API) for very large spreadsheets have to be produced.

# Warning

1. Avoid to upgrade  poi-ooxml to 4.0.0. that will cause shift rows go failing. See [Bug 62711 New: Calling shiftRows corrupts file in POI 4.0](http://apache-poi.1045710.n5.nabble.com/Bug-62711-New-Calling-shiftRows-corrupts-file-in-POI-4-0-td5731583.html)
