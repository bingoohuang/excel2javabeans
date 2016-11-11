package com.github.bingoohuang.excel2javabeans;

import com.github.bingoohuang.excel2javabeans.annotations.ExcelColTitle;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.commons.lang3.StringUtils;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
@Data @NoArgsConstructor @AllArgsConstructor
public class BeanWithTitle extends ExcelRowRef implements ExcelRowIgnorable {
    @ExcelColTitle("会员姓名") String memberName;
    @ExcelColTitle("性别") String sex;
    @ExcelColTitle("卡名称") String cardName;
    @ExcelColTitle("办卡价格") String cardPrice;

    @Override public boolean ignoreRow() {
        return StringUtils.startsWith(memberName, "示例-");
    }
}
