# 10 数字签名结构导读

上一课：[模板与附件](09-templates-attachments-tags.md) · 下一课：[OFD-H 是什么](11-ofd-h-overview.md) · 返回：[目录](README.md)

## 这一课结束时

能在包内找到签名列表和 `Signature.xml`，分清“保护文件摘要通过”和“电子印章密码学验证通过”，并且知道 SES/SM2 不在本课实现范围内。

发票、政务公文、部分病案归档会带签章。生成第一页文字不需要它；打开别人的文件时经常第一个碰到。

## 规范位置

第 18 章数字签名：签名列表、签名文件、摘要、签名范围、外观、签名值。电子印章的密码算法还涉及 GB/T 38540 等，本课只讲 OFD 包结构。

源项目入门：[ofdrw-sign quickstart](https://github.com/ofdrw/ofdrw/blob/master/ofdrw-sign/doc/quickstart/README.md)。

## 包内通常有什么

`OFD.xml` 的 `DocBody` 带可选节点：

```xml
<ofd:Signatures>Doc_0/Signs/Signatures.xml</ofd:Signatures>
```

列表文件再指向每个签名目录：

```text
Doc_0/Signs/
├── Signatures.xml
└── Sign_0/
    ├── Signature.xml      ← 签名描述：算法、保护哪些条目、外观
    ├── SignedValue.dat    ← 签名值（厂商/SES 结构）
    └── Seal.esl           ← 可选，电子印章
```

目录名因生成器而异（`Signs` / `Sign_0` 只是常见写法）。以 XML 里的 `ST_Loc` 为准。

`Signature.xml` 里最重要的几块：

1. **References**：被保护的包内文件列表及摘要（常见 SM3）
2. **签名范围**：哪些内容纳入保护
3. **外观**：印章图画在哪一页（往往是带 `StampAnnot` 的图元）
4. **签名值位置**：指向 `SignedValue.dat`

改了被保护条目的任何一个字节，引用摘要就会失败。这是阅读器“文件已被修改”的第一层。

## 两层验证

| 层次 | 检查什么 | 通过意味着什么 |
| --- | --- | --- |
| 引用完整性 | 每个保护文件的摘要是否匹配 | 声明范围内的文件没被改 |
| 签名值 | `SignedValue.dat` 是否由对应私钥/印章签出 | 签署者身份与不可抵赖（需证书链、算法实现） |

本仓库 `OfdSignatureVerifier`：

- 内置 SM3、SHA-1、SHA-256 的引用摘要比对
- `ReferenceIntegrityValid` 只表示第一层
- `FullyValid` 还要求为该签名方法注册了 `IOfdSignedValueVerifier` 且签名值校验成功

没有注册 SES/SM2 验证器时，CLI `ofdrw verify-signatures` 在引用完整的情况下退出码为 `2`，不是 `0`。不要把“摘要通过”说成“验章通过”。

源项目 `ofdrw-sign` + `ofdrw-gm` 覆盖更完整的国密签章数据结构；密码应用包 `ofdrw-crypto` 对应 GM/T 0099，本仓库明确未实现。

## 外观不是签名值

印章图片或矢量出现在页面上，只是 **外观**。删除外观图元不等于去掉签名；反之，贴一张章图片也不等于已签名。验证必须走 `Signature.xml` + `SignedValue.dat`。

本仓库重写已签名包且字节变化时，会去掉失效的签名声明并报告 `SignaturesInvalidated`。需要签章时应对 **最终字节** 重新签。

## 动手

对一份你有权使用的已签名 OFD：

```bash
unzip -l signed.ofd | grep -i sign
unzip -p signed.ofd OFD.xml
dotnet run --project src/Ofdrw.Net.Cli -- verify-signatures --input signed.ofd
```

对照输出区分引用完整性与完整有效。然后故意改一个 `TextCode` 再验，确认摘要失败。

创建签名需要实现 `IOfdSignatureProvider`：接收序列化后的 `Signature.xml` 字节，返回厂商签名值。编排由 `OfdSignatureService` 写入列表和保护引用。算法细节不在本教程展开。

## 常见坑

- **验了摘要就对外宣称合法电子签章**。缺证书链、时间戳、印章外观校验时只能报告完整性。
- **签名后再次转换格式**。PDF↔OFD 会改变字节，原签名作废。
- **保护列表漏了页面文件**。只保护 `OFD.xml` 改不了内容也能通过摘要。
- **多个签名的覆盖范围重叠又只验了一个**。应对列表中每一项分别验证。

## 本系列到这里为止

前 8 课覆盖打开包、找到页、写出文字/线/图。第 9–10 课覆盖真实文件里的底板、内嵌数据和签章入口。电子病历轮廓（`OFD-H`）对“至少一签、保护范围排除注释列表、阅读器只读”的额外要求见 [第 11–14 课](11-ofd-h-overview.md)。渐变、注释、动作、加密包见 [目录里的后置清单](README.md#刻意后置的内容)。实现能力边界以 [功能对照](../feature-parity.md) 为准，不要把教程里的“规范有这个对象”理解成本仓库已经完整实现。
