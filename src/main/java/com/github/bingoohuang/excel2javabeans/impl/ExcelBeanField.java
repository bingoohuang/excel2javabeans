package com.github.bingoohuang.excel2javabeans.impl;

import com.esotericsoftware.reflectasm.FieldAccess;
import com.esotericsoftware.reflectasm.MethodAccess;
import lombok.Data;

import static org.apache.commons.lang3.StringUtils.isNotEmpty;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
@Data
public class ExcelBeanField {
    private String name;
    private String setter;
    private String title;
    private int columnIndex;

    public <T> void setFieldValue(
            FieldAccess fieldAccess,
            MethodAccess methodAccess,
            T o,
            Object cellValue) {

        try {
            methodAccess.invoke(o, setter, cellValue);
            return;
        } catch (Exception e) {
            e.printStackTrace();
        }

        try {
            fieldAccess.set(o, name, cellValue);
            return;
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public boolean hasTitle() {
        return isNotEmpty(title);
    }

    public boolean containTitle(String cellValue) {
        return cellValue != null && cellValue.toUpperCase().contains(title);
    }

    public void setTitle(String title) {
        this.title = title.toUpperCase();
    }
}
