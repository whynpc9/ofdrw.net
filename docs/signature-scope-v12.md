# 签名保护范围、关系与扩展参数（R13）

`OfdSignatureService.SignWithResultAsync` 返回本次签名的包条目清单：`ProtectedFiles` 的原始字节进入 `References`，`UnprotectedFiles` 未进入。新生成的 `Signature.xml` 和 `SignedValue.dat` 必然列为本次未保护；它们由签名值本身的校验关系覆盖。旧 `SignAsync` 仍可用，默认保护所有签署前存在的条目，仅排除可变的 `Signatures.xml`。显式 `ProtectSignatureList=true` 后，后续追加签名会被拒绝。

`OfdSignatureOptions.ProtectFile` 接收以 `/` 开头的包内绝对路径，返回是否保护。它只能缩小默认引用集；不能排除根索引、未声明为旧签名或批注的正文目录内容、文档实际依赖的 XML 或共享资源。文档引用的 `PublicRes`、`DocumentRes`、`Attachments`、`CustomTags` 和 `Extensions` 描述符会按 XML 强制解析，不依赖 `.xml` 后缀或字节前缀；`BaseLoc` 与 `MultiMedia/@MediaFile`、`Font/FontFile` 会组合成真实包路径，再检查保护范围。调用者可显式排除文档声明的批注索引及其中的 `PageAnnot` 文件，让后续批注更新不破坏这份签名；由此批注本身不受这份签名保护。目录名本身不赋予例外：即使页面、字体或图片位于 `Annots/` 或 `Signs/` 下，正文与资源引用仍优先受保护。旧签名条目与正文依赖指向同一包路径时拒绝签署，包括默认排除可变签名清单的模式。无法证明依赖闭包时签署失败，且在调用 provider、写入输出流之前失败。签署结果逐条列出选择，不能从签名值有效推断任意未保护条目也完整。

Java OFDRW 标作 OFD 2.0 的 `RelativeSignatureId` 写在 `Signatures.xml` 的 `<Signature Relative="...">` 记录上，必须指向同一清单中的一个既有 ID。验签报告的 `RelativeIntegrityValid` 单独检查该关系及基签名的摘要和已注册容器验签结果；缺失、歧义或无效的基签名使整体 `FullyValid=false`。`Parameters` 写在新 `Signature.xml` 的 `SignedInfo/Parameters/Parameter`，每项有唯一非空 `Name`、可选 `Type` 和文本值。签名 provider 接收完整原始 `Signature.xml` 字节，因此参数变化应由签名值校验检出。当前 Writer 输出 OFD 1.0；使用这两个扩展字段时，签署器只接受明确声明 `Version="2.0"` 且使用 2016 命名空间的输入包，不会悄悄升级文档版本。

[GB/T 33190-2016 官方页面](https://openstd.samr.gov.cn/bzgk/std/newGbInfo?hcno=3AF6682D939116B6F5EED53D01A9DB5D)目前仍将 2016 版列为现行。上述 Java OFDRW 扩展模型和本库的 2.0 输入门不等于已证明符合现行国标或独立阅读器接受 `Version="2.0"`。本轮验证只覆盖准确列出的 XML、摘要、签名容器和包消费样本；标准符合性和桌面阅读器互操作须另行核准。

`ReferenceIntegrityValid` 只说明**列出的**引用摘要有效；`FullyValid` 另要求注册的签名值 verifier 成功，以及已声明的 `Relative` 基签名关系有效。它们都不表示所有业务内容被保护，更不表示自动完成证书信任、吊销、TSA 或法律效力验证。调用者应读取 `SignWithResultAsync` 的保护范围，并对自己的文档策略作独立判断。

`Relative` 位于可变的 `Signatures.xml` 清单里。默认续签模式不把清单纳入引用，否则无法追加签名；因此 `RelativeIntegrityValid` 检查的是**当前清单声明**与当前基签名的关系，不证明该属性从未被改写。需要冻结关系时，应在最终签名显式保护清单，并接受此后不能再追加的行为。不得把这项结构检查单独称为已认证的签名链。

XML 位置与 OFD 2.0 标注按固定版 Java OFDRW 的 [`Signature` 清单模型](https://github.com/ofdrw/ofdrw/blob/84e9eb260ba6fd351ecccf08a01a40305d4af635/ofdrw-core/src/main/java/org/ofdrw/core/signatures/Signature.java)、[`SignedInfo`](https://github.com/ofdrw/ofdrw/blob/84e9eb260ba6fd351ecccf08a01a40305d4af635/ofdrw-core/src/main/java/org/ofdrw/core/signatures/sig/SignedInfo.java)、[`Parameters`](https://github.com/ofdrw/ofdrw/blob/84e9eb260ba6fd351ecccf08a01a40305d4af635/ofdrw-core/src/main/java/org/ofdrw/core/signatures/sig/Parameters.java) 和 [`OFDSigner`](https://github.com/ofdrw/ofdrw/blob/84e9eb260ba6fd351ecccf08a01a40305d4af635/ofdrw-sign/src/main/java/org/ofdrw/sign/OFDSigner.java) 核对。上游过滤器会吞掉自身抛出的异常；本库的选择器异常会直接中止签署，以免意外扩大保护范围。

[`java-verify`](../e2e/Ofdrw.Net.Signatures.Scope.E2E/java-verify/pom.xml) 固定使用发布版 Java OFDRW `2.4.0`，通过 `OFDReader` 和其签名 XML 模型实际读取包消费样本，确认两条记录、`Relative=1` 和 `purpose` 参数。它是字段位置与可读性的独立实现证据，不执行 Java 端签名值验签。运行命令：`mvn -f e2e/Ofdrw.Net.Signatures.Scope.E2E/java-verify/pom.xml compile exec:java -Dexec.args=<生成的 r13-signed.ofd 路径>`；本轮在 `maven:3.9.9-eclipse-temurin-21` 容器中通过。
