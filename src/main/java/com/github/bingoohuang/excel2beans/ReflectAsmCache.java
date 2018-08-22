package com.github.bingoohuang.excel2beans;

import com.esotericsoftware.reflectasm.FieldAccess;
import com.esotericsoftware.reflectasm.MethodAccess;
import com.google.common.cache.CacheBuilder;
import com.google.common.cache.CacheLoader;
import com.google.common.cache.LoadingCache;
import lombok.extern.slf4j.Slf4j;
import lombok.val;
import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.Field;

@Slf4j
public class ReflectAsmCache {
    private LoadingCache<Class, MethodAccess> methodAccessCache =
            CacheBuilder.newBuilder().build(new CacheLoader<Class, MethodAccess>() {
                @Override public MethodAccess load(Class beanClass) {
                    return MethodAccess.get(beanClass);
                }
            });

    private LoadingCache<Class, FieldAccess> fieldAccessCache =
            CacheBuilder.newBuilder().build(new CacheLoader<Class, FieldAccess>() {
                @Override public FieldAccess load(Class beanClass) {
                    return FieldAccess.get(beanClass);
                }
            });

    public void setFieldValue(Field field, Object target, Object cellValue) {
        val setter = "set" + StringUtils.capitalize(field.getName());
        try {
            methodAccessCache.getUnchecked(field.getDeclaringClass())
                    .invoke(target, setter, cellValue);
            return;
        } catch (Exception e) {
            log.warn("call setter {} failed", setter, e);
        }

        try {
            fieldAccessCache.getUnchecked(field.getDeclaringClass())
                    .set(target, field.getName(), cellValue);
        } catch (Exception e) {
            log.warn("field set {} failed", field, e);
        }
    }

    public Object getFieldValue(Field field, Object target) {
        val getter = "get" + StringUtils.capitalize(field.getName());
        try {
            return methodAccessCache.getUnchecked(field.getDeclaringClass())
                    .invoke(target, getter);
        } catch (Exception e) {
            log.warn("call getter {} failed", getter, e);
        }

        try {
            return fieldAccessCache.getUnchecked(field.getDeclaringClass())
                    .get(target, field.getName());
        } catch (Exception e) {
            log.warn("field get {} failed", getter, e);
        }

        return "";
    }

}
