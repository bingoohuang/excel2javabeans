package com.github.bingoohuang.util.instantiator;

public interface BeanInstantiator<T> {
    <T> T newInstance();
}
