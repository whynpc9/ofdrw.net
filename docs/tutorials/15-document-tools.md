# 15. 文档工具：水印、Split、Mix 与清签

这些 API 面向已有 OFD；几何单位为毫米，API 页位置零基，CLI 页号一基。水印和注释外观是普通页面图元；它们不赋予密码学有效性。

## 水印

```csharp
var package = await new OfdReader().ReadAsync(input);
OfdWatermark.AddText(package, new[] { 0 }, "DRAFT 草稿",
    new OfdWatermarkOptions {
        LayerId = "draft", LayerType = "Foreground",
        XMillimeters = 50, YMillimeters = 250,
        WidthMillimeters = 110, HeightMillimeters = 16
    }, fontSizeMillimeters: 7);
OfdWatermark.AddImage(package, new[] { 0 }, pngBytes, "image/png",
    new OfdWatermarkOptions { LayerId = "draft", LayerType = "Foreground" });
await new OfdPackageWriter().WriteAsync(package, output);
```

已有 `LayerId` 追加在该图层最后一个对象之后；类型必须匹配。新 Background 图层放在正文对象之前，Body/Foreground 放在正文对象之后。模板仍按背景模板→正文→前景模板绘制，所以前景模板可覆盖正文水印。页列表必须非空、无重复、全部有效。验证或取消失败不会部分修改原模型。新图层自动生成唯一标识。

文字可抽取，图片作为普通资源写入；合并和 PDF/SVG 导出保留。默认最多 10,000 页、累计 1,000,000 个水印文字字符、32 MiB 编码图片、40,000,000 解码像素；支持 PNG/JPEG；每页图片对象有独立可变字节副本，累计生成图片默认预算 512 MiB。字体名按既有字体选择规则解析，调用方可在包中加入可分发的嵌入字体。

```bash
ofdrw watermark input.ofd text-mark.ofd --pages 1,2 --text 'DRAFT 草稿' --layer draft --layer-type Foreground --x 50 --y 250 --width 110 --height 16 --font-size 7
ofdrw watermark input.ofd image-mark.ofd --pages 1 --image mark.png --layer draft --layer-type Foreground --x 175 --y 215 --width 15 --height 15 --alpha 160
```

不传 `--pages` 时处理全部页。所有尺寸必须有限，宽高及字号为正；Alpha 为 0–255。

## Split

```csharp
var selection = OfdDocumentSplitter.Split(package, new[] { 2, 0 });
var cleanup = await new OfdPackageWriter().WriteWithResultAsync(selection, output);
```

```bash
ofdrw split input.ofd selected.ofd --pages 3,1
```

输出是按列表顺序排列的一个自包含包。原模型不变。保留选中页面原路径、模板、对应注释、附件和未知结构，保存时清未选页私有字体、图片、注释及不再使用的模板，并移除失效签名声明。共享资源或扩展仍引用的载荷保留；这属于引用闭包，不是泄漏清理失败。对于尚未保存的纯类型模型，Split 展开模板/注释并只复制实际使用的字体；带保留扩展且无原始包基线时须先保存再拆分。无法解析的 XML 使闭包无法证明，Split 拒绝处理。页号越界、重复及空列表同样拒绝。

## Mix

```csharp
var mixed = OfdDocumentMixer.Mix(new[] {
    new OfdMixSource(firstDocument, 0),
    new OfdMixSource(secondDocument, 1)
});
await new OfdPackageWriter().WriteAsync(mixed, output);
```

```bash
ofdrw mix mixed.ofd first.ofd 1 second.ofd 2
```

每个输入页按背景模板→正文图层→前景模板→可见注释外观展开，然后叠放下一个输入页。输出保留首输入页完整 PhysicalBox（包括原点），坐标不缩放、不居中，超出首页面框的内容按导出/阅读器的页面框裁切。输出图层统一为 Body，以独立图层 ID 保持输入顺序，避免阅读器按原图层类型重排跨来源内容。

字体以显式资源 ID 为优先；无 ID 时按字体名及 Weight/Italic 选择嵌入风格，按内容及风格身份重映射，图片按载荷复用，对象及嵌套 ID 重新分配。附件一并复制；原有签名不复制。注释外观的 Boundary、有限六元 CTM 和精确裁剪多边形合成为页面坐标；裁剪保存为独立 Clips 交集，PDF/SVG 也保留，PDF 字形随 CTM 缩放/旋转。支持文字、路径、图片及只有 ID 的嵌套 PageBlock（深度最多 32、每外观最多 100,000 图元）。复合/未知绘图或不支持的裁剪以整外观保留，Mix 拒绝、PDF/SVG 明确失败以免部分成功；未知外观属性在普通往返中保持原包，导出仍绘制命名空间及几何有效的标准子图元，同时保留阻止 Mix 的标记；未知绘图子树（含 vendor TextCode/Clips）不能被解释成正文。缺失/不可解析的注释元数据保持原包，Mix 仍拒绝。未建模资源引用、页动作、文档扩展同样拒绝。此限制避免产生丢内容的输出；不等于全部复杂 OFD 都可混合。

默认最多 1,000 个输入页、100,000 个对象、512 MiB 累计输入载荷（包括保留条目、字体、附件及正文/模板/注释图片；保守重复计数）。API 可调预算并传取消令牌。CLI 固定相同默认限额。

## 清签

```csharp
var cleanup = await OfdSignatureCleaner.CleanAsync(input, output,
    new OfdPackageLoadOptions(), cancellationToken);
```

```bash
ofdrw clean-signatures signed.ofd clean.ofd
ofdrw verify-signatures clean.ofd
# Signature verification: no signature declarations.
```

清签直接编辑包级条目，遍历全部 DocBody，保留正文条目字节及其他文档。删除标准签名声明；只有能证明属于 `Signs` 目录、无剩余引用的清单、描述、签名值及 Seal 外观载荷才删除。支持 Seal 的属性与子元素 BaseLoc。保护引用 `References/FileRef` 不是删除候选。指向正文、包根或签名目录外的 SignedValue 只产生保留诊断，不能授予正文删除权限。未知 XML 或共享扩展仍引用的载荷保留；路径逃逸导致操作失败。

验签区别“无声明”“引用摘要匹配”“密码学有效”。缺失/空/损坏的已声明清单不能报告无声明。清签不会生成生产签章。

## 输出、取消与样例

CLI 使用相邻临时文件，成功后原子替换，支持输入输出同路径；验证失败或取消保留旧输出。API 的模型操作先预检、返回新包或批量提交水印；任意 Stream 输出无法回滚，调用方应暂存后原子发布。文本抽取包含已支持的可见注释文字。注释索引与每份注释 XML 只解析一次，重复页/文件记录去重并保留列表顺序。清签拒绝非 OFD ZIP，去重签名描述并用队列扫描引用闭包，各循环传取消。Reader/Cleaner 接收既有压缩包、展开量、条目数和页数预算。

页面编辑限定单 DocBody；清签支持全部 DocBody。无隐私 API 样例与命令见 [DocumentTools E2E](../../e2e/Ofdrw.Net.DocumentTools.E2E/README.md)，实际验证范围见 [持久证据](../evidence/document-tools/README.md)。

Writer 自动生成的 name-only 假斜体矩阵带版本化命名空间提示，记录原始完整因子。PDF/SVG 保留该因子的文字锚点，但移除其字形剪切，语义斜体只绘制一次；外层变换和裁剪保留。未标记的用户 CTM 按原语义完整组合，不通过矩阵外观猜测其来源。旧版未标记的合成矩阵与相同用户矩阵无法可靠区分，本版本不启用自动猜测。资源/注释/模板改写仅作用于当前文档声明或经验证的隐式资源；未声明、vendor 同名 XML 保留。

Clips 写回位于继承的图元属性区（Actions 之后、TextCode/AbbreviatedData 等对象专有子节点之前），不将裁剪追加到内容之后。

注释绘图仅接受已建模的标准父子结构；TextCode 和 AbbreviatedData 必须是叶节点。未知属性保留原 XML 和已知绘图，但阻止 Mix；未知子元素或未支持的绘图资源引用使整组 Appearance 明确不支持导出。普通对象的 Mix 同样拒绝未知结构或属性，避免复制扩展引用却丢失其载荷。Merge 保留既有的无命名空间元数据兼容性；其未知子元素及外部命名空间属性仍明确拒绝。

标准 `Name`、`Visible`、裁剪 Path 的图形属性和 Area/Start 会保留；Visible=false 对象在 PDF/SVG 和可见文本提取中不绘制，但原 XML 仍可往返。图片 ImageMask 与 Substitution 均因尚未实现引用重映射而明确拒绝。正文/模板 TextCode 和 AbbreviatedData 只读取直接文本/CDATA，不把嵌套扩展文本拼入正文或路径，也不回退到整个对象的 Value。

已声明的注释图片 ID、载荷或页文件无法加载时，整组绘图会保留为不支持标记，PDF/SVG 和 Mix 明确失败。带 PageID 的未知索引记录只影响目标页；无归属记录聚合为共享元数据字符串，避免按页乘以记录数生成标记。

Mix 的 Template 引用包装器只接受已支持的 TemplateID/ZOrder 属性；额外属性、子元素或正文会保留在原包但阻止展平。split 的模板存活检查按 XML 内容识别未知扩展，不依赖后缀，并跟踪 ID、叶值及常规/Res BaseLoc 文件路径；无法安全解析的 XML 继续保守保留。未匹配已加载页面的注释 PageID 转入共享全局元数据，不能静默丢弃。

Mix 保留已建模的平面 EMR CustomTags，键区分大小写；唯一键复制，相同键值去重，相同键的不同值在预检时报出键及来源序号。Native 的 machine-readable/source 标签可保留；Native 与 DualLayer 的模式值冲突会明确失败。未知标签 XML、其他 schema、矛盾的重复记录仍拒绝；指向源包载荷或附件 ID 的潜在引用尚无重映射规则，也明确拒绝，不猜测替换字符串。标签值作为标量原样复制，不能据此声明任意附件路径/ID 已被重映射。vendor Template 引用保持原 XML，并且不会进入标准模板绘制列表；标准模板旁的 vendor 同名节点不导致重复绘制。

Mix 展平的是可见批注的静态 Appearance；标准 Annot 的 Type/Creator/LastModDate/Subtype、Remark 和 Parameters 等描述信息不复制成目标批注。未知 Annot 属性或子节点、Annotations/PageAnnot 根元数据会保留原包并阻止 Mix；明确 FileLoc 的扩展属性不会使已知批注漏画。隐藏批注中的扩展也参与预检。

Mix 还检查所选页面及其模板的原始 Page/Area/Content/Layer 包装器；仅支持已建模的 PhysicalBox 和图层属性，额外区域框、DrawParam、嵌套模板及扩展结构明确拒绝。模板展平仅支持 Background/Foreground；Body 尚未实现，拒绝展平。普通读取中空白或非法引用层级继承有效声明默认值，否则回退 Background；Mix 对这些原始未支持的层级值明确拒绝，避免丢失原 token。
