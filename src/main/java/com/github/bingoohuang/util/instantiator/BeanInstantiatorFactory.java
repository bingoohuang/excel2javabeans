package com.github.bingoohuang.util.instantiator;

import lombok.val;

import java.lang.reflect.Constructor;

public class BeanInstantiatorFactory {
    @SuppressWarnings("unchecked")
    public static <T> BeanInstantiator<T> newBeanInstantiator(Class<T> beanClass) {
        val ctors = (Constructor<T>[]) beanClass.getConstructors();
        for (val ctor : ctors) {
            if (ctor.getParameterTypes().length == 0) {
                return new ConstructorBeanInstantiator<T>(ctor);
            }
        }

        return new ObjenesisBeanInstantiator<T>(beanClass);
    }
}
