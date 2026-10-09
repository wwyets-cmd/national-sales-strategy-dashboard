# National Sales Strategy Dashboard

把全国销售发/退货明细一键生成自包含、离线可打开的 HTML 销售数据看板，适用于销售会议、经营复盘和区域/客户/品类/型号分析。

## 功能

- 顶部固定筛选器：年度、月份、省份、业务员、客户，全模块联动。
- KPI：总完成金额、同比增量、同比增长率、年度目标完成率。
- 月度同比、城市目标进度、品类饼图与柱状图。
- 基础五金看板：铰链、滑轨、反弹器。
- 功能五金看板：除基础五金和辅料外的品类。
- 辅料物料明细与客户辅料排名。
- 客户排名展开：月度同比、品类同比、TOP10 型号。
- 业务员业绩、品类对比、客户维度联动。
- 产品型号统计：品类筛选、型号实时搜索、同型号多尺寸归一。
- 月度分析：金额、同比、环比、自动判断。
- 智能分析建议：RFM、ABC、BCG、B2B销售漏斗、安索夫矩阵、风险预警和执行动作。

## 安装依赖

```bash
pip install pandas openpyxl
```

WPS/Excel 导出的 `.htm/.html` 大表使用 Python 标准库流式解析，不需要额外 HTML 解析依赖。

## 快速开始

```bash
cd skills/national-sales-strategy-dashboard

python3 scripts/build_dashboard.py \
  --file26 2026销售发退货明细.htm \
  --file25 2025销售发退货明细.xlsx \
  --targets assets/targets_example.json \
  --model-overrides assets/model_overrides_example.json \
  --out 全国销售数据看板.html
```

没有上年数据时：

```bash
python3 scripts/build_dashboard.py \
  --file26 2026销售发退货明细.htm \
  --skip-25 \
  --out 全国销售数据看板.html
```

## 输入字段

必需字段：

- 发/退货日期、客户名称、客户简称、省份、地级市
- 销售员、销货单号
- 物料描述、数量、含税本位币金额
- 大类名称、小类名称

支持 `.xlsx`、`.xls`、`.csv`、`.htm`、`.html`。脚本会自动扫描前 10 行定位表头。

## 目标配置

工具不内置任何全国、区域或业务员目标。使用者上传自己的目标 JSON，未配置的位置显示“待配置”，不会按 0 完成率计算。

简单城市目标：

```json
{
  "济南市": 600,
  "杭州市": {"province": "浙江省", "target_wan": 200}
}
```

完整区域目标：

```json
{
  "total_wan": 5000,
  "salesman_target_wan": {
    "张三": 1800,
    "李四": 1200
  },
  "cities": {
    "济南市": 600,
    "杭州市": {"province": "浙江省", "target_wan": 200}
  }
}
```

## 业务口径

- 基础五金 = 铰链 + 滑轨 + 反弹器。
- 反弹器严格取 `大类名称=OEM` 且 `小类名称=开门器`。
- 功能五金 = 除基础五金、辅料宣传外的所有品类。
- 铰链、滑轨保留原始物料描述及尺寸；其他品类自动提取关键型号。
- 型号例外可通过 `assets/model_overrides_example.json` 复制后维护。
- 新/老客户和新/老产品按上年同期是否有销售自动判定。

## 作为 Skill 使用

将 `skills/national-sales-strategy-dashboard` 文件夹复制到你的 Agent Skill 目录（通常为 `~/.doubao/agent_mode/workspace/.user_skills/`），之后上传销售明细并要求“生成全国销售数据看板”即可触发。

## 许可证

MIT
