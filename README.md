# 航班异常应对

赛题是「重点机场航班异常应对智能助手」。机场、城市、航班数字和规则都是赛事方编的，不是真实运行。

目录里两块是分开的：

- `baseline_diagnosis.py` 看月度基线。同样叫「异常」，换一种算法，命中的月份会变。天气和完成率去掉季节之后几乎不相关。
- `assistant.py` 对着模拟事件写建议。运行时不调用大模型，也不连接真实航班，更不能替人取消、放行或发旅客通知。

## 怎么跑

在这个目录里：

```
pip install -r requirements.txt
python baseline_diagnosis.py
python assistant.py --list
python assistant.py --event EVT001
python assistant.py --all
python assistant.py --tests
python assistant.py --question "航班是否一定会延误？"
python assistant.py --custom-json data/custom_event_example.json
python assistant.py --feedback EVT001 --decision 接受 --reason "先去核对保障能力"
```

这台机器上命令是 `python3` 的话，把上面的 `python` 换成 `python3`。

`--tests` 会把工作簿 `08_测试问题` 里的 15 道题送进助手，结果写到 `results/test_results.csv` 和 `docs/事件测试报告.md`。没通过时进程会以非 0 退出。

再分析一条事件时，`results/assistant_outputs.csv` 按事件编号更新那一行，不会把别的事件抹掉。反馈写在 `results/feedback_log.csv`，只追加，不删原来的建议。

## 交上去的材料

- `docs/客户汇报.md`
- `docs/数据与规则说明.md`
- `docs/事件测试报告.md`
- `docs/风险卡.md`
- `docs/实施路线图.md`
- `results/` 里的诊断、测试表、反馈、基线报告
- `figures/fig1_baseline_definitions.png`

这个文件夹连同 `requirements.txt` 就是复现包。

## Grok 用在哪

跑 `assistant.py` 的时候不会访问 Grok，也不会访问别的大模型。某次事件标成低、中、高还是无法判断，是脚本按关键词和表算的。

Grok 用在做这份作业的过程里。在 Cursor 里让 Grok 读 `data/Track 3 Data.xlsx`，把当时的原型和表里的东西对了一遍：`05_运营规则` 的 12 条、`06_智能助手输出` 要有的栏、`08_测试问题` 的 15 道题。对出来的洞主要是这些：基线算过但结论里没有数字，来源没写成 SIM-BASE、SIM-WEATHER、SIM-SCALE，RULE003、RULE007、RULE011、RULE012 在五件事件上走不到，另外有几道测试是看到题干里的词就自己标通过。后面改助手、改测试、补 docs，是按这个单子做的。

Grok 不给某一次运行打风险等级。改完之后用 `python assistant.py --tests` 重跑，题目要真正进 `analyze`。

## 还不行的地方

关键词会对「登机口频繁调整」这种句子，换种说法就可能漏。月度相关系数不能证明某一班是因为天气延误的。姓名和手机号、身份证号会从输出里拿掉，真上线还得另做脱敏和审批。没有接实时航班或气象，也没有运营的人在旁边盲测过。
