package com.github.bingoohuang.excel2beans;

import lombok.Builder;
import lombok.Data;

/**
 * @author bingoohuang [bingoohuang@gmail.com] Created on 2016/11/10.
 */
@Data @Builder public class SimpleBean {
    private String name;
    private String addr;
}
