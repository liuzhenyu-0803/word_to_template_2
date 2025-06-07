# 任务
- 为key-value数组中每个old_key， 从key描述文件中寻找语义完全匹配的new_key
- 语义完全匹配指new_key必须和old_key表达的完整语义完全一致（比如old_key表达的是某个温度，而new_key表达的是另一个温度，虽然都有温度，但不是完全匹配）

# 输入
## key-value数组：
placeholder_key_value_array

## key描述文件：
placeholder_key_description

# 输出(直接给结果)
[
  {
    "old_key": "old_key",
    "value": "value"
    "new_key": "key描述文件中的key或空字符串"
  },
  ...
]
