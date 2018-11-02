package com.github.bingoohuang.util.instantiator;

import lombok.AllArgsConstructor;
import lombok.SneakyThrows;

import java.lang.reflect.Constructor;

@AllArgsConstructor
public class ConstructorBeanInstantiator<T> implements BeanInstantiator<T> {
    private final Constructor<T> constructor;

    @SuppressWarnings("unchecked")
    @SneakyThrows @Override public T newInstance() {
        return constructor.newInstance();
    }
}
