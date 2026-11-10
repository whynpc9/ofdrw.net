# Ofdrw.Net

Ofdrw.Net 是 .NET OFD SDK 与 CLI，提供文档生成、读取、编辑、转换和可扩展的签名完整性检查。

**[阅读文档](https://whynpc9.github.io/ofdrw.net/)** · [安装与包选择](https://whynpc9.github.io/ofdrw.net/getting-started/installation/) · [快速开始](https://whynpc9.github.io/ofdrw.net/getting-started/quickstart/) · [能力矩阵](https://whynpc9.github.io/ofdrw.net/reference/feature-parity/)

## 安装

```sh
dotnet add package Ofdrw.Net.Converter --version 0.3.1
dotnet tool install --global Ofdrw.Net.Cli --version 0.3.1
```

普通 SDK 为 netstandard2.0 / netstandard2.1；CLI 需要 .NET 10。DOCX 默认 Native，直接保留 OpenXML 原文；PDF 导入默认双层。Skia、PDF vector、SES 和 Password 为独立可选扩展。

0.3.1 的变更与限制见[发布说明](https://whynpc9.github.io/ofdrw.net/releases/031/)。历史版本见[版本差异](https://whynpc9.github.io/ofdrw.net/troubleshooting/known-issues/)。

## 开发与证据

安装 Git LFS 后获取测试所需载荷：

```sh
GIT_LFS_SKIP_SMUDGE=1 git clone --depth 1 https://github.com/whynpc9/ofdrw.net.git
cd ofdrw.net
git lfs pull --include="e2e/**,tests/**" --exclude=""
```

历史证据默认保留为小指针，按需下载具名目录。构建产物放在 ignored artifacts/。提交前暂存变更并运行 `python3 scripts/check-repository-storage.py`。

- [工程与文档维护](https://github.com/whynpc9/ofdrw.net/blob/main/docs/engineering/documentation.md)
- [仓库存储规则](https://github.com/whynpc9/ofdrw.net/blob/main/docs/engineering/repository-storage.md)
- [发布门禁](https://github.com/whynpc9/ofdrw.net/blob/main/docs/engineering/release-gates.md)
- [路线图](https://github.com/whynpc9/ofdrw.net/blob/main/docs/engineering/roadmap.md)
- [验收资料](https://github.com/whynpc9/ofdrw.net/tree/main/docs/validation/)与[必要证据](https://github.com/whynpc9/ofdrw.net/tree/main/docs/evidence/)

## 许可与范围

项目采用 MIT，第三方依赖和字体遵循各自许可，见 [THIRD-PARTY-NOTICES](THIRD-PARTY-NOTICES.md)。本项目参考 [ofdrw/ofdrw](https://github.com/ofdrw/ofdrw) 的模块划分及目标，为独立 .NET 实现。

完整国家/行业标准符合性、厂商互认、签章法律效力和任意复杂 Word 保真均超出本版承诺范围。使用前按实际输入、字体、部署方式和阅读器验证结果。
