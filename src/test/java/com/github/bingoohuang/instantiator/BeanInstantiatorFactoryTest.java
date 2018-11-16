package com.github.bingoohuang.instantiator;

import lombok.Data;
import lombok.val;
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
        val instantiator = BeanInstantiatorFactory.newBeanInstantiator(NullaryConstructorBean.class);
        val bean = instantiator.newInstance();
        assertThat(bean.getName()).isEqualTo("bingoo");
    }

    @Test
    public void test2() {
        val instantiator = BeanInstantiatorFactory.newBeanInstantiator(NonNullaryConstructorBean.class);
        val bean = instantiator.newInstance();
        assertThat(bean.getName()).isNull();
    }
}