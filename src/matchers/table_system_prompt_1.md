# 任务
- 从表格中提取key-value关系
- 提取的key-value不要包含表头
- 计算value所在单元格行索引和列索引不能忽略表头
- 提取的key-value要包含value为空的情况

# 输入
placeholder_table_content

# 输出（直接给结果）
[
  {
    "key": "key",
    "value": "value"，
    "valuePos": "(value行索引，value列索引)"
  },
  ...
]
