package com.github.bingoohuang.excel2beans;

import lombok.Data;
import lombok.experimental.Accessors;

@Data
@Accessors(fluent = true)
public class TitleColumn {
  /** 标题 */
  String title;
  /** 示例数据 */
  String example;
}
