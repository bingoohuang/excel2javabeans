package com.github.bingoohuang.excel2beans;

import com.github.bingoohuang.excel2beans.annotations.ExcelColTitle;
import com.github.bingoohuang.util.GenericTypeUtil;
import com.google.common.collect.Lists;
import lombok.Getter;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import lombok.val;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellStyle;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;

import static org.apache.commons.lang3.StringUtils.defaultIfEmpty;

@Slf4j
public class ExcelBeanField {
    @Getter private final Field field;

    private final boolean titleRequired;
    @Getter private final Class elementType;
    private final Method valueOfMethod;
    private final ReflectAsmCache reflectAsmCache;

    @Setter @Getter private boolean titleColumnFound;
    @Setter @Getter private int columnIndex;
    @Setter @Getter private CellStyle cellStyle;

    @Getter private final String title;
    @Getter private final boolean cellDataType;
    @Getter private final boolean multipleColumns;
    @Getter private final List<Integer> multipleColumnIndexes = Lists.newArrayList();

    public ExcelBeanField(Field f, int columnIndex, ReflectAsmCache reflectAsmCache) {
        this.field = f;
        this.columnIndex = columnIndex;

        val ct = f.getAnnotation(ExcelColTitle.class);
        this.titleRequired = ct != null && ct.required();
        this.title = ct != null ? defaultIfEmpty(ct.value(), f.getName()).toUpperCase() : null;

        val gtu = new GenericTypeUtil(f.getGenericType());
        this.multipleColumns = gtu.isParameterized() && List.class.isAssignableFrom(f.getType());
        this.elementType = this.multipleColumns ? gtu.getActualTypeArg(0) : f.getType();
        this.cellDataType = this.elementType == CellData.class;
        this.valueOfMethod = elementType == String.class ? null : ValueOfs.getValueOfMethodFrom(elementType);
        this.reflectAsmCache = reflectAsmCache;
    }

    public void setFieldValue(Object target, Object cellValue) {
        reflectAsmCache.setFieldValue(field, target, cellValue);
    }

    public Object getFieldValue(Object target) {
        return reflectAsmCache.getFieldValue(field, target);
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
        return valueOfMethod == null ? cellValue
                : ValueOfs.invokeValueOf(elementType, cellValue);
    }

    public boolean isTitleNotMatched() {
        return hasTitle() && titleRequired && !titleColumnFound;
    }

    public boolean isElementTypeSupported() {
        return isImageDataField() || isStringField() || cellDataType || valueOfMethod != null;
    }

}
