package com.github.bingoohuang.util.instantiator;

import lombok.val;

import java.lang.reflect.Constructor;

public class BeanInstantiatorFactory {
    public static <T> BeanInstantiator<T> newBeanInstantiator(Class<T> beanClass) {
        Constructor<T>[] constructors = (Constructor<T>[]) beanClass.getConstructors();
        for (val constructor : constructors) {
            if (constructor.getParameterTypes().length == 0) {
                return new ConstructorBeanInstantiator<T>(constructor);
            }
        }

        return new ObjenesisBeanInstantiator<T>(beanClass);
    }
}
