# 上游 Java OFDRW 的 DocType 与 OFD-A 分层对照

核对日期：2026-11-17。上游 `ofdrw/ofdrw` 默认分支为 `master`，核对提交 [`026a53419eeae2a5c0ff5dc467f8b6db19a93d16`][upstream]；本库对照基线为 `826d9cef0c346881466406a5b7d45a6e69a2cbfd`，参照[归档路线图](plans/ofd-archival-roadmap.md)和 [D0 验收摘要](../validation/standard-doctype-d0-20261117.md)。上游默认分支及提交由 Git 远端 HEAD 和浅克隆确认；POM 在该提交标为 `2.5.1`，本文仅对该源码快照负责。

**结论：上游与本库 D0/A1 的方向一致。** 普通生成和格式转换默认 `OFD`；基础模型能携带 `OFD-A` 声明；归档检查和去技术化处理在显式调用的独立模块。Java 上游提供了检查器和处理器的分层样例，不过所查普通阅读和编辑路径没有按 `DocType` 实施保护，归档转换没有最终字节复检，静态调用链还暴露了转换缓存未落盘的风险（未运行复现）。因此不宜照搬它的标签时序、输出提交或有损处理策略。

## 普通 OFD、读入声明与归档标记

| 路径 | Java 上游当前行为 | 与本库的关系 |
| --- | --- | --- |
| Core 新建 | `OFD()` 写 `Version=1.2` 和 `DocType=OFD`，接收 `DocBody` 的新建构造器复用该默认值。[`OFD.java` 构造器][ofd-new] | D0 将默认 DocType 统一为 `OFD`，无需因本次对照改变命名空间或版本约定。 |
| Layout、PDF、图片和文本导入 | [`OFDDoc` 新建][layout-new]使用 `new OFD()`；[`PDFConverter`][pdf]、[`PDFGraphicsConverter`][pdf-graphics]通过 [`OFDGraphicsDocument`][graphics-new] 新建普通 OFD；[`ImageConverter`][image]、[`TextConverter`][text]使用 `OFDDoc`。这些入口没有自动调用归档层。 | 本库 Layout、DOCX 默认/Native/DualLayer 与 PDF 导入仍输出普通 OFD；Java 路径用于对照通用默认策略，不替代本库 DOCX 验收。 |
| 读入及低层显式赋值 | [`OFDDir.getOfd()`][container-root]用 XML 元素代理构造 `OFD`，不会补写声明；[`getDocType()`][ofd-type]读实际值，缺失为 `null`；`setDocType` 只拒绝 `null` 或空串，既不限定为 `OFD`/`OFD-A`，也不检查文件是否符合归档要求。 | 本库 [Reader][net-reader] 对缺失属性提供 `OFD` 回退；归档 [AC-001][net-coverage-type] 读取原始 XML，仍能单独指出缺失或异常声明。 |
| 编辑既有文件 | [`OFDDoc(OFDReader, …)` 的初始化][layout-edit]复用 reader 的容器、设置修改日期并加载正文/资源，没有按 `DocType` 拒绝编辑，也没有把旧声明改回 `OFD`。 | A2 的原件保护和普通 OFD 派生副本是本库拟增加的策略。只读检查或导出不应因声明而被阻断；声明也不证明读取软件符合 GB/T 42133-2022 第 7 章。 |

在此提交全仓跟踪的生产 Java 源码中搜索 `DocType`、`DOC_TYPE`、`OFD-A` 和 `OFD-H`：实际访问或修改 `DocType` 的逻辑集中于基础 `OFD` 模型和归档规则/处理器，没有找到另外的普通 Reader、Editor、Converter 类型分流；生产 Java 中未找到 `OFD-H`。这是该快照的静态搜索范围，不是对上游所有历史版本或下游应用的断言。

## 独立模块与实际注册项

`org.ofdrw:ofdrw-archive` 是[根 Maven reactor 中单独的子模块][root-pom]；[自身 POM][archive-pom]直接依赖 Reader、Crypto 和 Commons IO，Layout 仅用于测试。核对 [`ofdrw-full`][full-pom]、[`ofdrw-converter`][converter-pom] 的 POM 及其同仓正式模块依赖闭包，未发现它们引入 Archive。根 reactor 构建 Archive 与应用消费普通模块时自动安装 Archive 是不同的事。这个依赖方向支持本库 `Archival → Packaging → Core` 的选择；Java 模块直接带 Reader/Crypto 并不意味着本库只读首版也需要这些依赖。

| 计数口径 | 检查规则 | 转换处理器 |
| --- | ---: | ---: |
| [README 架构图][archive-readme-count] | 28 | 20 |
| [README 处理顺序表][archive-readme-order] | — | 23（编号 0–22） |
| 默认构造器实际加载的列表 | [23][rules] | [23][handlers]（含首项解密） |

这是列表注册数，不是标准条款覆盖数、已验证效果或 `ofdrw.net` 条款账本的对照计数。[README][archive-readme-scope]也写明实现不完整。举例：[`FontSubsetRule`][font-rule]检查 `FontFile` 及路径是否存在，不解析字体字节、实际字形、子集表或豁免；[`ResourcePlacementHandler`][placement-handler]虽已注册，移动资源的方法体只遍历页，未实现跨文件移动和引用更新。不能将 README 对应条款的勾选当作功能验收。

## Checker：性质、结果与异常

[`OFDArchiveChecker`][checker]既可接收已打开的 `OFDReader`，也可从路径临时打开；按注册顺序调用规则，合并 [`ArchiveViolation`][violation]（规则名、严重程度 `ERROR/WARN/INFO`、描述、位置、实际值、期望值），按规则名排序。无参构造加载默认规则，自定义列表构造直接替换。规则返回 `null` 等同没有新增违规；单条规则抛 `Exception` 时，以规则类名记录 `ERROR` 后继续。路径入口打开文档失败不会转成某条规则结果。

[`DocTypeRule`][doctype-rule]直接检查根声明，精确等于 `OFD-A` 才返回空列表；缺失、非 `OFD-A` 和根 XML 读取异常都会给出 `DOC_TYPE` 错误。普通 `OFD` 因此不满足归档标识规则，仍可作为普通 OFD；其他规则照常执行。空列表仅说明本次执行的规则没有返回违规；API 没有声明状态、输入支持范围、条款适用性、`Unknown`、运行完整性或整标准结论等独立维度。作为对照，本库已有 [AC-001][net-coverage-type] 和[覆盖报告][net-coverage]，整体保持 `Incomplete` / `NotAssessed`、`StandardOutcome` 为 `Fail` 或 `Unknown`；A1 可复用现有原始声明检查和这套报告边界。

## Converter：实际执行顺序、落盘与输出

[`OFDArchiveConverter.convert(src, dst)`][convert] 校验输入/输出参数，解压到临时 `src/`，运行固定队列；随后从磁盘 `OFD.xml` 开始收集可达文件并复制到临时 `output/`，直接打包到 `dst`，最后尽力清理临时目录。转换器支持用自定义列表替换整条队列，[`addDecryptor`][convert-custom] 需队列中有 `EncryptionHandler`。链路没有先调用 Checker 来决定执行哪些处理器，处理器逸出的异常会中止后续步骤；个别处理器或遍历方法也有自行吞掉异常的分支。

1. **打标在前。** [默认顺序][handlers]为 `EncryptionHandler` → `SingleDocHandler` → `DocTypeHandler` → 其余处理器。第一步发现根 `Encryptions.xml` 却没有解密器会抛错；配置了解密器时，代码尝试打包、调用 Crypto 解密并解压回工作目录。[`EncryptionHandler`][encryption] 仅说明入口和失败分支，未验证任何加密样例。第三步 [`DocTypeHandler`][doctype-handler]调用 `setDocType("OFD-A")` 和 `ofdDir.setOfd(ofd)`；这是改类型属性的意图，独立于后续处理成功和符合性判断。
2. **此处改属性与最终文件落盘之间有断点。** [`setOfd`][container-root]调用 [`VirtualContainer.putObj`][container-cache]，只更新实例中的对象缓存，[`flush()`][container-flush]才写磁盘。转换器给 handler 的 `OFDDir` 和 [reader 自己的 `OFDDir`][reader-dir] 是两个实例；默认队列在 `DocTypeHandler` 之后没有刷新源 `OFDDir`，`OFDReader(..., false)` 的 [`close()`][reader-close]也不会刷新它。最后的 [`jar(dst)`][container-jar]刷新的是新建的输出容器，复制所用的仍是源目录磁盘文件。根据这条静态调用链，不能把 handler 的赋值视为输出 `OFD.xml` 已改为 `OFD-A`。现有 [`DocTypeRuleTest`][doctype-test] 使用的是[手工修改声明、对同一个容器实例 `jar` 的 fixture][doctype-fixture]，未走这一转换链。
3. **可达遍历并未按 XML 引用语义实现。** [`dfsCollect`][reachable]读取 UTF-8 文本，以正则挑选形似路径的双引号字符串，并相对于包根查找；它不识别 XML 引用字段。普通 Layout 写出的 [`DocRoot`][docroot]是元素文本，例如 `<ofd:DocRoot>Doc_0/Document.xml</ofd:DocRoot>`，其本身不会被此正则匹配。源码可见漏收合法引用的条件，因此不能用该机制证明来源标签、正文或资源完整，也不宜复制到 A1b/A3。
4. **没有最终产物交付门禁。** 默认转换没有重开最终 ZIP 检查，也没有基于 Checker 结果阻止发布。`convert` 没有拒绝源目标同路径；[`OFDDir.jar(Path)`][container-jar]会删除已有目标，然后直接打包。因此即使源内容在临时目录中处理，也不能推导源目标同路径安全或失败时可回滚。无异常返回不能证明输出完整、可读或符合归档要求。

## 丢失内容与签名的边界

- **多文档。** README 顺序表将 `SingleDocHandler` 描述为合并；[`SingleDocHandler`][single-doc]实际尝试删除第 2 个及其后的文档目录，在容器中只保留第一个 `DocBody`，没有合并页面或生成损失报告。上面的缓存问题也影响根节点变更。A5 仍需单独定义保留、合并、拒绝和来源记录。
- **签名/签章。** [`SignatureRule`][signature-rule]枚举默认文档签名清单，分别报告 `Sign` 或 `Seal` 需要处理；它不做签名校验。默认 [`SignatureHandler`][signature-handler]对读到的非空签名清单尝试删除签名目录、清单文件及根中的引用，没有实现验签、签名人/证书/时间/摘要提取、`SignRef` 处置或印章外观转注释。不能以它返回成功证明原签名仍有效或归档证据已保存。[README 详细表][signature-readme]也将部分签名信息保留标为框架/自定义项。
- **附件。** [`AttachmentHandler`][attachment-handler]按声明格式保留 TXT/XML，尝试删除其他附件文件并移除条目；没有生成摘要或将附件转换为 OFD 追加页面的逻辑。处理范围和损失必须有独立决策，不能由处理器名称推断。

## 对路线图的建议

| 工作包 | 建议 |
| --- | --- |
| D0 | 维持已集成的普通 `OFD` 默认策略。回读原声明及缺失回退各自按本库契约处理，继续区分 main 集成与 NuGet 发布。 |
| A1 | 借鉴独立模块及规则聚合；继续以固定输入 profile 做只读评估。分别报告原始声明、适用范围、单条规则结果和执行状态。已知 CustomTags 按 schema 遍历；字体以实际包内载荷、字形与豁免证据为准。 |
| A2 | 明确产品层哪些修改入口保护 OFD-A 原件，验证派生副本的普通 OFD 声明、来源和原件不变；只读路径继续可用。 |
| A3 | 延用计划中的独立候选 → 写出候选 `OFD-A` → 重开最终字节 → 检查证据和交付资格 → 发布。应从最终 ZIP 断言声明与必要内容同时存在。本库现有 [`ArchivalCopyConverter`][net-copy] 已对有限条款做最终字节复检、拒绝签名输入，而且只修改 `Document.xml` 的相应声明，不写 `OFD-A`。它可作为这种输出边界的本地参考。 |
| A4/A5 | 上游的处理器可用作场景清单。对改变外观、签名、附件和多文档语义的处理，仍按计划逐项指定输入范围、信息损失、拒绝条件及最终证据。 |

本次只核对固定快照的 Java 源码、Maven POM、README 和已有测试代码。宿主没有可用 Java Runtime 或 Maven；未执行转换、测试、产物渲染或阅读器验证。上述落盘和引用结论来自静态追踪；本文不评定整部标准、第三方软件行为、Maven 实际发包或完整安全性。

[upstream]: https://github.com/ofdrw/ofdrw/commit/026a53419eeae2a5c0ff5dc467f8b6db19a93d16
[ofd-new]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-core/src/main/java/org/ofdrw/core/basicStructure/ofd/OFD.java#L23-L66
[ofd-type]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-core/src/main/java/org/ofdrw/core/basicStructure/ofd/OFD.java#L88-L123
[layout-new]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-layout/src/main/java/org/ofdrw/layout/OFDDoc.java#L278-L305
[layout-edit]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-layout/src/main/java/org/ofdrw/layout/OFDDoc.java#L313-L331
[pdf]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-converter/src/main/java/org/ofdrw/converter/ofdconverter/PDFConverter.java#L125-L142
[pdf-graphics]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-converter/src/main/java/org/ofdrw/converter/ofdconverter/PDFGraphicsConverter.java#L124
[graphics-new]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-graphics2d/src/main/java/org/ofdrw/graphics2d/OFDGraphicsDocument.java#L135-L147
[image]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-converter/src/main/java/org/ofdrw/converter/ofdconverter/ImageConverter.java#L64-L77
[text]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-converter/src/main/java/org/ofdrw/converter/ofdconverter/TextConverter.java#L63-L76
[reader-dir]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-reader/src/main/java/org/ofdrw/reader/OFDReader.java#L153-L175
[reader-close]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-reader/src/main/java/org/ofdrw/reader/OFDReader.java#L955-L966
[container-root]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-pkg/src/main/java/org/ofdrw/pkg/container/OFDDir.java#L108-L122
[container-cache]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-pkg/src/main/java/org/ofdrw/pkg/container/VirtualContainer.java#L284-L297
[container-flush]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-pkg/src/main/java/org/ofdrw/pkg/container/VirtualContainer.java#L694-L711
[container-jar]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-pkg/src/main/java/org/ofdrw/pkg/container/OFDDir.java#L408-L439
[docroot]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-core/src/main/java/org/ofdrw/core/basicStructure/ofd/DocBody.java#L58-L74
[root-pom]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/pom.xml#L12-L36
[archive-pom]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/pom.xml#L12-L45
[full-pom]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-full/pom.xml#L16-L77
[converter-pom]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-converter/pom.xml#L20-L95
[archive-readme-count]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/README.md#L46-L62
[archive-readme-order]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/README.md#L64-L92
[archive-readme-scope]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/README.md#L1-L7
[rules]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/src/main/java/org/ofdrw/archive/check/OFDArchiveChecker.java#L114-L155
[handlers]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/src/main/java/org/ofdrw/archive/convert/OFDArchiveConverter.java#L297-L345
[checker]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/src/main/java/org/ofdrw/archive/check/OFDArchiveChecker.java#L45-L104
[violation]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/src/main/java/org/ofdrw/archive/check/ArchiveViolation.java#L20-L45
[doctype-rule]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/src/main/java/org/ofdrw/archive/check/rule/DocTypeRule.java#L26-L49
[doctype-handler]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/src/main/java/org/ofdrw/archive/convert/handler/DocTypeHandler.java#L23-L33
[doctype-test]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/src/test/java/org/ofdrw/archive/DocTypeRuleTest.java#L19-L55
[doctype-fixture]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/src/test/java/org/ofdrw/archive/TestOFDGenerator.java#L153-L170
[font-rule]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/src/main/java/org/ofdrw/archive/check/rule/FontSubsetRule.java#L17-L65
[placement-handler]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/src/main/java/org/ofdrw/archive/convert/handler/ResourcePlacementHandler.java#L38-L70
[convert]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/src/main/java/org/ofdrw/archive/convert/OFDArchiveConverter.java#L116-L161
[convert-custom]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/src/main/java/org/ofdrw/archive/convert/OFDArchiveConverter.java#L67-L104
[reachable]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/src/main/java/org/ofdrw/archive/convert/OFDArchiveConverter.java#L173-L258
[encryption]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/src/main/java/org/ofdrw/archive/convert/handler/EncryptionHandler.java#L47-L109
[single-doc]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/src/main/java/org/ofdrw/archive/convert/handler/SingleDocHandler.java#L29-L70
[signature-rule]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/src/main/java/org/ofdrw/archive/check/rule/SignatureRule.java#L26-L62
[signature-handler]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/src/main/java/org/ofdrw/archive/convert/handler/SignatureHandler.java#L40-L118
[signature-readme]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/README.md#L407-L432
[attachment-handler]: https://github.com/ofdrw/ofdrw/blob/026a53419eeae2a5c0ff5dc467f8b6db19a93d16/ofdrw-archive/src/main/java/org/ofdrw/archive/convert/handler/AttachmentHandler.java#L23-L55
[net-reader]: https://github.com/whynpc9/ofdrw.net/blob/826d9cef0c346881466406a5b7d45a6e69a2cbfd/src/Ofdrw.Net.Reader/Readers/OfdReader.cs#L155-L165
[net-coverage-type]: https://github.com/whynpc9/ofdrw.net/blob/826d9cef0c346881466406a5b7d45a6e69a2cbfd/src/Ofdrw.Net.Packaging/Preflight/ArchivalCoverageService.cs#L152-L162
[net-coverage]: https://github.com/whynpc9/ofdrw.net/blob/826d9cef0c346881466406a5b7d45a6e69a2cbfd/src/Ofdrw.Net.Packaging/Preflight/ArchivalCoverageService.cs#L86-L114
[net-copy]: https://github.com/whynpc9/ofdrw.net/blob/826d9cef0c346881466406a5b7d45a6e69a2cbfd/src/Ofdrw.Net.Packaging/Preflight/ArchivalCopyConverter.cs#L71-L141
