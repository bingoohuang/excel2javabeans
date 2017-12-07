package com.github.bingoohuang.excel2beans;

import lombok.Data;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.List;

@Data
public class BeanClassBag {
    private Class<?> beanClass;
    private Sheet sheet;
    private List<ExcelBeanField> beanFields;
    private boolean firstRowCreated;

    public BeanClassBag(Class<?> beanClass) {
        this.beanClass = beanClass;
    }

    public ExcelBeanField getBeanField(int index) {
        return beanFields.get(index);
    }
}
