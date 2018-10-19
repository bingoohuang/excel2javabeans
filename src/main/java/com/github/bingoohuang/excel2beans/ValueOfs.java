package com.github.bingoohuang.excel2beans;

import com.google.common.collect.Maps;
import com.google.common.primitives.Primitives;
import lombok.val;

import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.util.Map;
import java.util.Optional;

public class ValueOfs {
    public static final Map<Class, Optional<Method>> valueOfMethodCache = Maps.newConcurrentMap();

    @SuppressWarnings("unchecked")
    public static Method getValueOfMethodFrom(Class returnClass) {
        val existsMethod = valueOfMethodCache.get(returnClass);
        if (existsMethod != null) return existsMethod.orElse(null);

        val clazz = Primitives.wrap(returnClass);
        try {
            val valueOfMethod = clazz.getMethod("valueOf", String.class);
            if (isPublicStatic(clazz, valueOfMethod)) {
                valueOfMethodCache.put(returnClass, Optional.of(valueOfMethod));
                return valueOfMethod;
            }
        } catch (Exception e) {
            valueOfMethodCache.put(returnClass, Optional.empty());
        }

        return null;
    }

    public static Object invokeValueOf(Class clazz, String value) {
        val valueOfMethod = valueOfMethodCache.get(clazz);
        if (valueOfMethod == null || !valueOfMethod.isPresent()) return null;

        try {
            return valueOfMethod.get().invoke(null, value);
        } catch (Exception e) {
            // ignore
        }

        return null;
    }

    public static boolean isPublicStatic(Class returnType, Method valueOfMethod) {
        return Modifier.isStatic(valueOfMethod.getModifiers())
                && Modifier.isPublic(valueOfMethod.getModifiers())
                && valueOfMethod.getReturnType().isAssignableFrom(returnType);
    }
}
