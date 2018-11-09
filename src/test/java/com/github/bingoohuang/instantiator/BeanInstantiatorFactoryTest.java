package com.github.bingoohuang.instantiator;

import lombok.Data;
import org.junit.Test;

import static com.google.common.truth.Truth.assertThat;


public class BeanInstantiatorFactoryTest {
    @Data
    public static class NullaryConstructorBean {
        private String name = "bingoo";
    }

    @Data
    public static class NonNullaryConstructorBean {
        private String name = "bingoo";

        public NonNullaryConstructorBean(String name) {
        }
    }


    @Test
    public void test1() {
        BeanInstantiator<NullaryConstructorBean> instantiator = BeanInstantiatorFactory.newBeanInstantiator(NullaryConstructorBean.class);
        NullaryConstructorBean bean = instantiator.newInstance();
        assertThat(bean.getName()).isEqualTo("bingoo");
    }

    @Test
    public void test2() {
        BeanInstantiator<NonNullaryConstructorBean> instantiator = BeanInstantiatorFactory.newBeanInstantiator(NonNullaryConstructorBean.class);
        NonNullaryConstructorBean bean = instantiator.newInstance();
        assertThat(bean.getName()).isNull();
    }
}