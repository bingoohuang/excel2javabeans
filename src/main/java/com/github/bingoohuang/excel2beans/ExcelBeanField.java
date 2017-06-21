package com.github.bingoohuang.excel2beans;

import com.esotericsoftware.reflectasm.FieldAccess;
import com.esotericsoftware.reflectasm.MethodAccess;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;

import java.lang.reflect.Field;

import static org.apache.commons.lang3.StringUtils.isNotEmpty;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
@Data @Slf4j
public class ExcelBeanField {
    private String name;
    private String setter;
    private String getter;
    private boolean titleColumnFound;
    private String title;
    private int columnIndex;
    private CellStyle cellStyle;
    private Field field;
    private boolean cellDataType;

    public <T> void setFieldValue(
            FieldAccess fieldAccess,
            MethodAccess methodAccess,
            T o,
            Object cellValue) {

        try {
            methodAccess.invoke(o, setter, cellValue);
            return;
        } catch (Exception e) {
            log.warn("call setter {} failed", setter, e);
        }

        try {
            fieldAccess.set(o, name, cellValue);
            return;
        } catch (Exception e) {
            log.warn("field set {} failed", name, e);
        }
    }


    public Object getFieldValue(
            FieldAccess fieldAccess,
            MethodAccess methodAccess,
            Object o) {

        try {
            return methodAccess.invoke(o, getter);
        } catch (Exception e) {
            log.warn("call getter {} failed", getter, e);
        }

        try {
            return fieldAccess.get(o, name);
        } catch (Exception e) {
            log.warn("field get {} failed", getter, e);
        }

        return "";
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
