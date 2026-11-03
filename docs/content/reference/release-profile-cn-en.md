---
title: "中英文受限发布配置"
---

`0.1.0-preview.10` 的 DOCX 页面验收配置为 **`cn-en-physical-cjk-400-700-v1`**。发布状态、产品源码/控制器、公开包及签名、独立消费与实际页面哈希见[最终发布记录](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/validation/preview10-publication-2026-11-02.md)。本文保存配置与承诺边界；安装及完整代码见[字体教程](https://whynpc9.github.io/ofdrw.net/guides/fonts/)。

## 固定字体配置

固定输入为未修改的 F01 `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`，SHA-256 `17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`。独立运行显式 `OfdMode=Native` 与不设置 `OfdMode` 的默认 Native，两模式各两页；查看链路为 DOCX → OFD → PDF → macOS Preview。

API 使用显式 `FontDirectories`、`NativeEmbeddedBoldBehavior=Throw`、`UseInstalledMicrosoftOfficeFonts=false`；清空 fallback 后依次指定 `Noto Sans CJK SC`、`Noto Sans`。Office 开关不禁用全部系统字体解析，`Throw` 只约束选中的 Native 嵌入粗体；不能把配置称为全局字体隔离。已发布的 preview.10 CLI 没有同等严格粗体开关；后续源码新增的 `--native-embedded-bold-behavior throw` 及其限定范围见[字体教程](https://whynpc9.github.io/ofdrw.net/guides/fonts/)，不改变本次发布记录。

| 字体 / face index | 源字节 / SHA-256 | 限定用途 |
| --- | --- | --- |
| `Noto Sans CJK SC` Regular / 0，静态 weight 400 | 21,659,544 / `3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa` | 中文与已覆盖的混排常规面 |
| `Noto Sans CJK SC` Bold / 0，静态 weight 700 | 21,638,668 / `647a6e375c5a0d8958111173d7b7cec5237e0784cc4490663e43834dfff491be` | 同族真实粗体；不得替换成 Regular 加样式标志或描边 |
| `Noto Sans` Regular / 0 | 569,208 / `b85c38ecea8a7cfb39c24e395a4007474fa5a4fc864f6ee33309eb4948d232d5` | 比例英文常规面 |

CJK 来源 revision `f8d157532fbfaeda587e826d4cd5b21a49186f7c`、源 VF SHA-256 `990c807e79c25662a5a9ecf7f971baeb2bf2eab9a559e5ecf15cdfdb8561d21f`；fontTools `4.59.2` 离线实例化 400/700，运行时不是变量字体。来源、固定参数和 OFL 许可证详见教程。字体不加入 NuGet，不增加字体载荷。实际输出的嵌入子集身份另验，不能要求其哈希等于源字体。

此配置没有 Noto Sans Bold 或物理 Italic，不承诺任意家族/face、缺字 fallback 或合成字重的视觉质量。其他真实样式须提供对应合法字体并另验；既有斜体可模拟，不能称为物理 Italic。

## 验收结果与限制

公开包独立消费的两模式均保留 189 个非空白源字符，PDF 无重复文本；四页 PNG 与实际 Preview 查看通过。通过仅覆盖上述配置、固定输入与四页，不保证任意 Word、任意字体或跨机像素一致；AI reviewer 不称人工签收。

- 原 Regular-only 默认 TextStroke 的表头“量”字字腔视觉 **FAIL 未修复**；真实 Bold 通过不覆盖旧失败。
- macOS Preview 完整页末日期/部分整词 Find 异常根因未定，公开消费阶段未重测。
- 原始 Unicode 剪贴板 Copy 未完成验收；这不是已经证实的 Copy 缺陷，其他阅读器 Copy/Find 也未验证。

历史失败见[本地独立验证](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/validation/cn-en-release-local-validation.md)及[系统对照](https://github.com/whynpc9/ofdrw.net/blob/2489966c47833cb137fe33cdc4c3d2df60daec3a/docs/validation/cn-en-preview-control-2026-11-02.md)。后续产品候选仍须按[发布门禁](https://github.com/whynpc9/ofdrw.net/blob/main/docs/engineering/release-gates.md)重新绑定源码、包与真实页面；本配置不自动放行新版本。
