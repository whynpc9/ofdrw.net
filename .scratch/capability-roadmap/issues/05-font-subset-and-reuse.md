# 字体子集化与复用（issue #3）

**What to build:** 嵌入字体按实际用字子集化；相同字体文件按内容身份复用，禁止无差别整份嵌入。中英多样例下 OFD 体积明显下降；缺字有回退；许可处理有说明。

**Blocked by:** None (can start immediately)

**Priority:** P1

**Status:** ready-for-agent

- [ ] 按用字子集化嵌入；相同字节字体只保留一份
- [ ] 中英样例体积相对整份嵌入明显下降；缺字回退有行为、许可有说明
- [ ] 更新功能对照

## Parent

[GitHub issue #3](https://github.com/whynpc9/ofdrw.net/issues/3)；[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P1-02

## What to build

issue #3 的另一半：OFD 太大，是因为整份字体反复打进包。这张票让生成/转换路径只嵌入用到的字形，并按字体内容身份复用（现有注册预算可以继续用）。缺字时走已有回退族，而不是静默空白。文档说明子集化与字体许可的关系：库不替调用方完成授权。

## Acceptance criteria

- [ ] 生成含中英、有限字符集的 OFD 时，嵌入字体是子集而不是完整 TTF/OTF；相同内容的字体在包内只有一份
- [ ] 用同一中英多样例对比「整份嵌入」与「子集+复用」，输出体积明显下降，并记下数量级
- [ ] 缺字走回退策略（配置族或明确失败），不出现空白框冒充成功
- [ ] 文档说明许可：子集化不改变调用方是否有权嵌入该字体
- [ ] 相关转换/布局测试通过；更新功能对照；不要关闭 issue #3

## Blocked by

- None (can start immediately)
