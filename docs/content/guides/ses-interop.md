---
title: "SES 测试互操作扩展"
---

`Ofdrw.Net.Signatures.SesInterop` 是独立测试扩展，需要调用方显式注册。它覆盖具名 SES 测试容器、固定测试证书、SM2 自签和限定交叉检查。

核心签名 API 不自带通用 SES / SM2 verifier，默认 CLI 也不会因为测试自签成功而给出生产 FullyValid。测试容器成功不证明厂商阅读器互认、可信时间戳、证书信任或法律效力。

扩展包仅按已核对的容器版本和字段结构使用。具体依赖、许可、样本和结果见[SES 证据](https://github.com/whynpc9/ofdrw.net/blob/main/docs/evidence/ses-interop/README.md)，核心模型见[签名指南](https://whynpc9.github.io/ofdrw.net/guides/signatures/)。
