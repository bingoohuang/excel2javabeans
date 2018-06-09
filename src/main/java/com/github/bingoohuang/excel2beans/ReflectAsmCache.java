package com.github.bingoohuang.excel2beans;

import com.esotericsoftware.reflectasm.FieldAccess;
import com.esotericsoftware.reflectasm.MethodAccess;
import com.google.common.cache.CacheBuilder;
import com.google.common.cache.CacheLoader;
import com.google.common.cache.LoadingCache;

public class ReflectAsmCache {
    private LoadingCache<Class, MethodAccess> methodAccessCache =
            CacheBuilder.newBuilder().build(new CacheLoader<Class, MethodAccess>() {
                @Override public MethodAccess load(Class beanClass) {
                    return MethodAccess.get(beanClass);
                }
            });

    private static LoadingCache<Class, FieldAccess> fieldAccessCache =
            CacheBuilder.newBuilder().build(new CacheLoader<Class, FieldAccess>() {
                @Override public FieldAccess load(Class beanClass) {
                    return FieldAccess.get(beanClass);
                }
            });

    public MethodAccess getMethodAccess(Class<?> beanClass) {
        return methodAccessCache.getUnchecked(beanClass);
    }

    public FieldAccess getFieldAccess(Class<?> beanClass) {
        return fieldAccessCache.getUnchecked(beanClass);
    }
}
