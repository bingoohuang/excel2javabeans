package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.google.common.collect.Lists;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import lombok.val;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellStyle;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.ParameterizedType;
import java.util.List;

@Slf4j
public class ExcelBeanField {
    @Getter private final String fieldName;
    private final String setter;
    private final String getter;

    private final boolean titleRequired;
    @Getter private final Class elementType;
    private final Method valueOfMethod;
    private final Class<?> beanClass;

    @Setter @Getter private boolean titleColumnFound;
    @Setter @Getter private int columnIndex;
    @Setter @Getter private CellStyle cellStyle;

    @Getter private final String title;
    @Getter private final boolean cellDataType;
    @Getter private final boolean multipleColumns;
    @Getter private final List<Integer> multipleColumnIndexes = Lists.newArrayList();

    public ExcelBeanField(Class<?> beanClass, Field field, int columnIndex) {
        this.beanClass = beanClass;
        this.columnIndex = columnIndex;
        this.fieldName = field.getName();
        this.setter = "set" + StringUtils.capitalize(fieldName);
        this.getter = "get" + StringUtils.capitalize(fieldName);

        val colTitle = field.getAnnotation(ExcelColTitle.class);
        if (colTitle != null) {
            this.titleRequired = colTitle.required();
            this.title = colTitle.value().toUpperCase();
        } else {
            this.titleRequired = false;
            this.title = null;
        }

        val genericType = field.getGenericType();
        val isParameterizedType = genericType instanceof ParameterizedType;

        if (isParameterizedType && List.class.isAssignableFrom(field.getType())) {
            val parameterizedType = (ParameterizedType) genericType;
            val actualTypeArgs = parameterizedType.getActualTypeArguments();

            this.multipleColumns = true;
            this.elementType = (Class) actualTypeArgs[0];
        } else {
            this.multipleColumns = false;
            this.elementType = field.getType();
        }

        this.cellDataType = this.elementType == CellData.class;

        this.valueOfMethod = elementType != String.class
                ? ValueOfs.getValueOfMethodFrom(elementType) : null;
    }

    public void setFieldValue(Object target, Object cellValue) {
        try {
            val methodAccess = ReflectAsms.getMethodAccess(beanClass);
            methodAccess.invoke(target, setter, cellValue);
            return;
        } catch (Exception e) {
            log.warn("call setter {} failed", setter, e);
        }

        try {
            val fieldAccess = ReflectAsms.getFieldAccess(beanClass);
            fieldAccess.set(target, fieldName, cellValue);
            return;
        } catch (Exception e) {
            log.warn("field set {} failed", fieldName, e);
        }
    }

    public Object getFieldValue(Object target) {
        try {
            val methodAccess = ReflectAsms.getMethodAccess(beanClass);
            return methodAccess.invoke(target, getter);
        } catch (Exception e) {
            log.warn("call getter {} failed", getter, e);
        }

        try {
            val fieldAccess = ReflectAsms.getFieldAccess(beanClass);
            return fieldAccess.get(target, fieldName);
        } catch (Exception e) {
            log.warn("field get {} failed", getter, e);
        }

        return "";
    }

    public boolean hasTitle() {
        return StringUtils.isNotEmpty(title);
    }

    public boolean containTitle(String cellValue) {
        return cellValue != null && cellValue.toUpperCase().contains(title);
    }

    public boolean isImageDataField() {
        return elementType == ImageData.class;
    }

    private boolean isStringField() {
        return elementType == String.class;
    }

    public void addMultipleColumnIndex(int columnIndex) {
        multipleColumnIndexes.add(columnIndex);
    }

    public Object convert(String cellValue) {
        return valueOfMethod == null
                ? cellValue
                : ValueOfs.invokeValueOf(elementType, cellValue);
    }

    public boolean isTitleNotMatched() {
        return hasTitle() && titleRequired && !titleColumnFound;
    }

    public boolean isElementTypeSupported() {
        return isImageDataField() || isStringField() || cellDataType || valueOfMethod != null;
    }


}
