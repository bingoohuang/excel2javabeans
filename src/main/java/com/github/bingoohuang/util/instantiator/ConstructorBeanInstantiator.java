package com.github.bingoohuang.util.instantiator;

import lombok.AllArgsConstructor;
import lombok.SneakyThrows;

import java.lang.reflect.Constructor;

@AllArgsConstructor
public class ConstructorBeanInstantiator<T> implements BeanInstantiator<T> {
    private final Constructor<T> constructor;

    @SneakyThrows @Override public <T> T newInstance() {
        return (T) constructor.newInstance();
    }
}
