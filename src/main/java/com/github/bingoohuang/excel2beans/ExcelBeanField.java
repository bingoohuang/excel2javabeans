package com.github.bingoohuang.excel2beans;

import com.esotericsoftware.reflectasm.FieldAccess;
import com.esotericsoftware.reflectasm.MethodAccess;
import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.google.common.collect.Lists;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import lombok.val;
import org.apache.poi.ss.usermodel.CellStyle;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.util.List;

import static org.apache.commons.lang3.StringUtils.capitalize;
import static org.apache.commons.lang3.StringUtils.isNotEmpty;

@Slf4j
public class ExcelBeanField {
    @Getter private final String fieldName;
    private final String setter;
    private final String getter;

    private final boolean titleRequired;
    private final Class elementType;
    private final Method valueOfMethod;

    @Setter @Getter private boolean titleColumnFound;
    @Setter @Getter private int columnIndex;
    @Setter @Getter private CellStyle cellStyle;

    @Getter private final String title;
    @Getter private final boolean cellDataType;
    @Getter private final boolean multipleColumns;
    @Getter private final List<Integer> multipleColumnIndexes = Lists.newArrayList();

    public ExcelBeanField(Field field, int columnIndex) {
        this.columnIndex = columnIndex;
        this.fieldName = field.getName();
        this.setter = "set" + capitalize(fieldName);
        this.getter = "get" + capitalize(fieldName);

        val colTitle = field.getAnnotation(ExcelColTitle.class);
        if (colTitle != null) {
            this.titleRequired = colTitle.required();
            this.title = colTitle.value().toUpperCase();
        } else {
            this.titleRequired = false;
            this.title = null;
        }

        val genericType = field.getGenericType();
        val isCollectionGeneric = genericType instanceof ParameterizedType
                && List.class.isAssignableFrom(field.getType());
        if (isCollectionGeneric &&
                ((ParameterizedType) genericType).getActualTypeArguments().length == 1) {
            this.multipleColumns = true;
            this.elementType = (Class) ((ParameterizedType) genericType).getActualTypeArguments()[0];
        } else {
            this.multipleColumns = false;
            this.elementType = field.getType();
        }

        this.cellDataType = this.elementType == CellData.class;

        this.valueOfMethod = elementType != String.class ? ValueOfs.getValueOfMethodFrom(elementType) : null;
    }

    public void setFieldValue(FieldAccess fieldAccess, MethodAccess methodAccess, Object o, Object cellValue) {
        try {
            methodAccess.invoke(o, setter, cellValue);
            return;
        } catch (Exception e) {
            log.warn("call setter {} failed", setter, e);
        }

        try {
            fieldAccess.set(o, fieldName, cellValue);
            return;
        } catch (Exception e) {
            log.warn("field set {} failed", fieldName, e);
        }
    }


    public Object getFieldValue(FieldAccess fieldAccess, MethodAccess methodAccess, Object o) {
        try {
            return methodAccess.invoke(o, getter);
        } catch (Exception e) {
            log.warn("call getter {} failed", getter, e);
        }

        try {
            return fieldAccess.get(o, fieldName);
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

    public boolean isImageDataField() {
        return elementType == ImageData.class;
    }

    public void addMultipleColumnIndex(int columnIndex) {
        multipleColumnIndexes.add(columnIndex);
    }


    public Object convert(String cellValue) {
        if (valueOfMethod == null) return cellValue;

        return ValueOfs.invokeValueOf(elementType, cellValue);
    }

    public boolean isTitleNotMatched() {
        return hasTitle() && titleRequired && !titleColumnFound;
    }

}
