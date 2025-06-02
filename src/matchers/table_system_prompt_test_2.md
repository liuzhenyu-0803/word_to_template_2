# 任务
## 从key描述文件中，为key-value数组中的每个old_key寻找new_key
## new_key的含义必须和old_key完全一致

# 输入
## key-value数组：
[
    {
    "old_key": "Measurements",
    "value": ""
    },
    {
    "old_key": "Sp1",
    "value": "29.3℃"
    },
    {
    "old_key": "Bx1",
    "value": ""
    },
    {
    "old_key": "Max",
    "value": "35.8℃"
    },
    {
    "old_key": "Avg",
    "value": "28.1℃"
    },
    {
    "old_key": "Min",
    "value": "23.2℃"
    },
    {
    "old_key": "Dt1(公式)",
    "value": ""
    },
    {
    "old_key": "Bx1.Max-Sp1DeltaValue",
    "value": "6.5℃"
    }
]

## key描述文件：
key_Sp1_temperature：Sp1测量点的温度
key_max_temperature：bx1测量区域的最高温度
key_avg_temperature：bx1测量区域的平均温度
key_min_temperature：bx1测量区域的最低温度
key_relative_humidity：相对湿度

# 输出
[
  {
    "old_key": "old_key",
    "value": "value或空字符串"
    "new_key": "new_key或空字符串"
  },
  ...
]

