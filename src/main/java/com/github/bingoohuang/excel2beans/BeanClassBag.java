package com.github.bingoohuang.excel2beans;

import lombok.Data;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;

@Data @RequiredArgsConstructor
public class BeanClassBag {
    private final Class<?> beanClass;

    private Sheet sheet;
    private List<ExcelBeanField> beanFields;
    private boolean firstRowCreated;

    public ExcelBeanField getBeanField(int index) {
        return beanFields.get(index);
    }
}
