package com.github.bingoohuang.instantiator;

import org.objenesis.ObjenesisStd;
import org.objenesis.instantiator.ObjectInstantiator;

public class ObjenesisBeanInstantiator<T> implements BeanInstantiator<T> {
    private final ObjectInstantiator<T> instantiator;

    public ObjenesisBeanInstantiator(Class<T> beanClass) {
        this.instantiator = new ObjenesisStd().getInstantiatorOf(beanClass);
    }

    @Override public T newInstance() {
        return instantiator.newInstance();
    }
}
