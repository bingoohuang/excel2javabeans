package com.github.bingoohuang.util;

import lombok.Getter;

import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;

@Getter
public class GenericTypeUtil {
    private final Type genericType;
    private final boolean isParameterized;
    private final ParameterizedType parameterized;

    public GenericTypeUtil(Type genericType) {
        this.genericType = genericType;
        this.isParameterized = genericType instanceof ParameterizedType;
        this.parameterized = isParameterized ? ((ParameterizedType) genericType) : null;
    }

    public Class<?> getActualTypeArg(int index) {
        return isParameterized ? (Class<?>) parameterized.getActualTypeArguments()[index] : null;
    }
}
