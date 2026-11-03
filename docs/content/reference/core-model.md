---
title: "OFD 模型与读写边界"
---

Ofdrw.Net 使用强类型页面、图层、文字游程、路径、图片、字体、模板和附件模型。`OfdReader` 读入容器，`OfdPackageWriter` 写出 OFD 包，FlowDocument 和 OfdDocumentBuilder 提供不同层次的生成入口。

OFD 是 ZIP 容器，OFD.xml 指向文档，文档再引用页面与资源。SDK 对未知 XML / 包条目采取保留策略，但保留不等于能够渲染、编辑或安全重映射。

单位通常为毫米，页面和 API 页索引从零开始。多 DocBody 的默认选择和显式枚举见[多文档指南](https://whynpc9.github.io/ofdrw.net/guides/multi-docbody/)。

公开能力以[能力矩阵](https://whynpc9.github.io/ofdrw.net/reference/feature-parity/)和[分出口约定](https://whynpc9.github.io/ofdrw.net/reference/capability-contract-v12/)为准。医疗、归档和密码规范的完整合规不由默认 DocType 或局部结构推定。
