package com.github.bingoohuang.util.instantiator;

import lombok.SneakyThrows;

import java.lang.reflect.Constructor;

public class ConstructorBeanInstantiator<T> implements BeanInstantiator<T> {
    private final Constructor<T> constructor;

    public ConstructorBeanInstantiator(Constructor<T> constructor) {
        this.constructor = constructor;
    }

    @SneakyThrows @Override public <T> T newInstance() {
        return (T) constructor.newInstance();
    }
}
