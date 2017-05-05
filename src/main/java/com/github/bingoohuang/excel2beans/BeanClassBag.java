package com.github.bingoohuang.excel2beans;

import com.esotericsoftware.reflectasm.FieldAccess;
import com.esotericsoftware.reflectasm.MethodAccess;
import lombok.Data;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Created by bingoohuang on 2017/3/20.
 */
@Data
public class BeanClassBag {
    private Class<?> beanClass;
    private Sheet sheet;
    ExcelBeanField[] beanFields;
    private FieldAccess fieldAccess;
    private MethodAccess methodAccess;
    private boolean firstRowCreated;

    public BeanClassBag(Class<?> beanClass) {
        this.beanClass = beanClass;
        this.fieldAccess = FieldAccess.get(beanClass);
        this.methodAccess = MethodAccess.get(beanClass);
    }
}
