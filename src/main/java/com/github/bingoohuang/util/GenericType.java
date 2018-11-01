package com.github.bingoohuang.util;

import lombok.Getter;

import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;

@Getter
public class GenericType {
    private final Type genericType;
    private final boolean isParameterized;
    private final ParameterizedType parameterized;

    public GenericType(Type genericType) {
        this.genericType = genericType;
        this.isParameterized = genericType instanceof ParameterizedType;
        this.parameterized = isParameterized ? ((ParameterizedType) genericType) : null;
    }

    public static GenericType of(Type genericType) {
        return new GenericType(genericType);
    }

    public Class<?> getActualTypeArg(int index) {
        return isParameterized ? (Class<?>) parameterized.getActualTypeArguments()[index] : null;
    }

    public boolean isRawType(Type type) {
        return isParameterized && parameterized.getRawType() == type;
    }
}
