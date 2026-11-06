using System;
using System.Collections.Generic;
using System.Linq;

namespace Ofdrw.Net.Packaging.Preflight;

/// <summary>A ledger obligation, not an assertion of compliance or applicability.</summary>
public sealed class ArchivalClauseDefinition
{
    internal ArchivalClauseDefinition(string clause, string condition, string gap)
    { Clause = clause; Condition = condition; RemainingEvidence = gap; }
    /// <summary>Fixed primary standard and edition.</summary>
    public string StandardId => "GB/T 42133-2022";
    /// <summary>Location of the previously inspected standard copy; not fetched by this API.</summary>
    public string StandardSource => ArchivalCoverageRuleCatalog.Rules[0].Source;
    /// <summary>SHA-256 of that previously inspected standard copy.</summary>
    public string StandardSourceSha256 => ArchivalCoverageRuleCatalog.Rules[0].SourceSha256;
    /// <summary>Exact clause identifier from ArchivalClauseCatalog.</summary>
    public string Clause { get; }
    /// <summary>Ledger condition or treatment boundary, not the full normative text.</summary>
    public string Condition { get; }
    /// <summary>Known coverage boundaries and material requirements from the ledger revision.</summary>
    public string RemainingEvidence { get; }
    /// <summary>Mapped subset rules only; their Pass never closes the whole obligation.</summary>
    public IReadOnlyList<string> RuleIds => Array.AsReadOnly(ArchivalCoverageRuleCatalog.Rules
        .Where(x => x.Clause == Clause).Select(x => x.RuleId)
        .Concat(ArchivalRuleCatalog.Rules.Where(x => x.Clause == Clause).Select(x => x.RuleId)).ToArray());
    /// <summary>The only existing archival copy handlers. Other clauses have no implicit handler.</summary>
    public string? Handler => Clause == "6.2.2(a)" || Clause == "6.2.2(b)" ? "ArchivalCopyConverter" : null;
}

/// <summary>Revision 1 inventory of chapters 5-7 and Annex A; independent of local APF policy.</summary>
public static class ArchivalClauseCatalog
{
    /// <summary>Version identifier for this obligation inventory.</summary>
    public const string Revision = "ofdrw-archival-ledger/1";
    /// <summary>Every ledger obligation, including unimplemented and externally assessed items.</summary>
    public static IReadOnlyList<ArchivalClauseDefinition> Clauses { get; } = Array.AsReadOnly(new[]
    {
        new ArchivalClauseDefinition("5.1(a)", "归档文本/图形场景", "外部证据：预期用途和档案策略。"),
        new ArchivalClauseDefinition("5.1(b)", "选定的照片档案", "外部证据：收集范围及排除的专业图像。"),
        new ArchivalClauseDefinition("5.1(c)", "管理统计/记录", "外部证据：记录分类。"),
        new ArchivalClauseDefinition("5.2(a)", "长期格式特性", "外部证据：DA/T 47 以及独立的可移植性/显示证据。"),
        new ArchivalClauseDefinition("5.2(b)", "归档单元转换/组织", "外部证据：GB/T 18894 和源记录来源。"),
        new ArchivalClauseDefinition("6.1（第一段）", "所有显示/打印/交互数据及相关资源位于包内", "部分覆盖：APF-002/003/005 检查已知边；未知 XML、完整资源闭包和外部语义仍未覆盖。"),
        new ArchivalClauseDefinition("6.1（第二段）", "可选元数据、签名、标签和附件", "未覆盖：可选性及其各自条款需要单独评估。"),
        new ArchivalClauseDefinition("6.2.1(a)", "根 `DocType` 标识 OFD-A", "**精确：**`AC-001`，仅检查可识别根节点和无限定属性。"),
        new ArchivalClauseDefinition("6.2.1(b)", "自定义元数据名称及优先独立存储", "外部证据：DA/T 46/54 映射和 GB/T 33190 的许可尚未评估。"),
        new ArchivalClauseDefinition("6.2.1(c)", "不使用多文档机制；必要时拆分或合并来源", "**精确（仅最终包数量）：**`AC-002`。拆分/合并来源和追加文件列表仍未覆盖。"),
        new ArchivalClauseDefinition("6.2.1(d)", "保留现有 `Versions`", "未覆盖：需要转换前后来源证据，不能只检查一个最终包。"),
        new ArchivalClauseDefinition("6.2.1(e)", "移除与根节点及传递引用内容无关的文件", "部分覆盖：APF 检查已知引用。R18 可选物理附件移除能证明请求载荷的所有权和引用闭包，但不能证明整包归档可达性，也不会删除所有无关文件。"),
        new ArchivalClauseDefinition("6.2.2(a)", "移除文档 `Permissions`", "**精确：**R17 `AR-001`；仅对可识别 XML 提供可选的审计副本处理器。"),
        new ArchivalClauseDefinition("6.2.2(b)", "移除文档 `VPreferences`", "**精确：**R17 `AR-002`；副本边界相同。"),
        new ArchivalClauseDefinition("6.2.2(c)", "移除文档中除文档内 Goto 外的动作", "**精确的只读类型检查：**`AC-003` 检查一个可识别的 `Document.xml`。没有移除处理器；不支持或不透明的形式为 Unknown。书签/页面目标有效性和其他动作位置仍分别未覆盖。"),
        new ArchivalClauseDefinition("6.2.2(d)", "汇总未保留的附件", "未覆盖：保留决策、摘要和描述需要来源证据。"),
        new ArchivalClauseDefinition("6.2.2(e)", "移除 `Extensions`，同时保留所有视觉效果", "外部证据：需要感知来源的外观比较和页面转换。仅检查节点不存在不能使此条款通过。"),
        new ArchivalClauseDefinition("6.2.3(a)", "将最常用的页面区域提升为文档默认值", "未覆盖：统计页面区域、保持外观并验证输出。"),
        new ArchivalClauseDefinition("6.2.3(b)", "优先省略冗余页面区域", "未覆盖：可选优化，需要比较页面与默认值。"),
        new ArchivalClauseDefinition("6.2.3(c)", "移除非 Goto 页面动作", "**精确的只读类型检查：**`AC-004` 遍历可识别的页面和模板描述；无法解析的页面条目和不透明动作形式为 Unknown。Goto 目标有效性另行检查。没有移除处理器。"),
        new ArchivalClauseDefinition("6.2.3(d)", "区分同名图层", "**精确的只读名称检查：**`AC-011` 检测每个可识别页面/模板内容内重复的无限定 Layer 名称；无法解析的页面或外来图层名称为 Unknown。没有重命名处理器或来源命名证据。"),
        new ArchivalClauseDefinition("6.2.3(e)", "移除 PageBlock 嵌套，或限制不可避免的嵌套", "未覆盖：需要完整页面/模板遍历和渲染等价性。"),
        new ArchivalClauseDefinition("6.2.4(a)", "优先保留原图层顺序及正文/背景组织", "外部证据：生成意图和视觉比较。"),
        new ArchivalClauseDefinition("6.2.4(b)", "将扫描图像和识别文本放入适当图层", "外部证据：扫描/OCR 来源和页面检查。"),
        new ArchivalClauseDefinition("6.2.4(c)", "优先采用语义内容顺序", "外部证据：语义阅读顺序和可访问性评审。"),
        new ArchivalClauseDefinition("6.2.5(a)", "移除非 Goto 大纲动作", "**精确的只读类型检查：**`AC-005` 检查嵌套的可识别 `OutlineElem` 动作。不透明的大纲/动作作用域为 Unknown。没有移除处理器。"),
        new ArchivalClauseDefinition("6.2.5(b)", "移除指向不存在书签/页面的大纲跳转", "**精确的只读目标检查：**`AC-006` 根据所选文档解析可识别的 Goto 书签名和页面 ID。有歧义的清单为 Unknown。没有移除处理器。"),
        new ArchivalClauseDefinition("6.2.6(a)", "将资源类型放入公共/文档资源表", "未覆盖：验证每个资源表和引用作用域。"),
        new ArchivalClauseDefinition("6.2.6(b)(1)", "嵌入或子集化填充/描边 TextObject 引用的字体", "部分覆盖：APF-003 只检查文件是否存在及已知字体 ID；真实字体有效性、字形使用和 GB/T 33190 豁免仍未覆盖。"),
        new ArchivalClauseDefinition("6.2.6(b)(2)", "包含裁剪文字使用的字体", "未覆盖：裁剪遍历和实际字形依赖。"),
        new ArchivalClauseDefinition("6.2.6(c)(1)", "子集化时保留字形轮廓表", "未覆盖：检查实际嵌入的 OpenType 表和轮廓正确性。"),
        new ArchivalClauseDefinition("6.2.6(c)(2)", "保留必需的度量/程序/name 表", "未覆盖：同样需要字体字节证据。"),
        new ArchivalClauseDefinition("6.2.6(c)(3)", "保留忠实绘制所需的其他表", "外部证据：字形渲染和平台比较。"),
        new ArchivalClauseDefinition("6.2.6(d)", "优先使用 GB 18030 字符集和 UCS 映射", "外部证据：GB 18030、GB/T 13000 以及文本/字体映射测试。"),
        new ArchivalClauseDefinition("6.2.6(e)", "限制栅格资源格式，或在归档前转换", "未覆盖：解码经认证的媒体载荷，不能只看文件名。"),
        new ArchivalClauseDefinition("6.2.6(f)", "栅格资源不含自定义扩展数据", "未覆盖：解码每种受支持图像格式及其扩展。"),
        new ArchivalClauseDefinition("6.2.6(g)", "移除音频/视频，并在其他位置保留摘要", "未覆盖：需要感知来源的转换和摘要证据。"),
        new ArchivalClauseDefinition("6.3.1(a)", "可定义默认颜色空间", "未覆盖：这是许可而非强制存在规则；只有在完整作用域检查后，对象不存在才能记为 NotApplicable。"),
        new ArchivalClauseDefinition("6.3.1(b)", "限制颜色空间类型", "未覆盖：遍历资源作用域和引用的颜色空间。"),
        new ArchivalClauseDefinition("6.3.1(c)", "优先使用颜色配置文件", "未覆盖：可选指导和颜色管理证据。"),
        new ArchivalClauseDefinition("6.3.2(a)", "移除完全包含对象边界的裁剪", "未覆盖：实际变换后的裁剪几何和视觉等价性。"),
        new ArchivalClauseDefinition("6.3.2(b)", "移除零面积裁剪并隐藏对象", "未覆盖：几何、变换和可见性比较。"),
        new ArchivalClauseDefinition("6.3.3(a)", "用适当紧致矩形限制可见文字", "外部证据：字形几何和外观检查。"),
        new ArchivalClauseDefinition("6.3.3(b)", "移除对象动作，迁移文档内链接", "**部分：**`AC-010` 会使页面、模板或批注外观中保留的可识别对象动作失败。没有来源动作清单以及链接批注迁移、视觉和目标证据时，未发现动作仍为 Unknown。"),
        new ArchivalClauseDefinition("6.3.3(c)", "优先在 Stroke=false 时移除描边属性", "未覆盖：需要处理样式继承的可选清理。"),
        new ArchivalClauseDefinition("6.3.3(d)", "优先移除冗余组合绘制属性", "未覆盖：继承的资源语义和外观。"),
        new ArchivalClauseDefinition("6.3.3(e)", "优先移除图像/视频填充属性", "未覆盖：对象类型和渲染语义。"),
        new ArchivalClauseDefinition("6.4(a)", "优先为重复路径共享颜色 DrawParam", "未覆盖：需要来源等价性的可选去重。"),
        new ArchivalClauseDefinition("6.4(b)", "优先为重复路径共享描边 DrawParam", "未覆盖：同上，包括虚线、端点和连接样式。"),
        new ArchivalClauseDefinition("6.5(a)", "优先采用适合图像共享范围的资源表", "未覆盖：跨页/页内使用清单。"),
        new ArchivalClauseDefinition("6.5(b)", "使用 `ImageObject Interpolate` 时，其值为 false", "**精确的只读属性检查：**`AC-007` 扫描可识别页面/模板和保留批注外观对象；true 失败，false 通过，完整作用域内属性不存在为 NotApplicable，不透明图形作用域为 Unknown。"),
        new ArchivalClauseDefinition("6.6(a)", "优先让每个匹配字体的行使用一个文本对象", "外部证据：来源行和字体游程语义。"),
        new ArchivalClauseDefinition("6.6(b)", "拆分内置字体/字符豁免之外的文本", "未覆盖：GB/T 33190 豁免和字形游程验证。"),
        new ArchivalClauseDefinition("6.6(c)", "用 `Size` 标识文字大小", "**精确的只读属性检查：**`AC-008` 要求可识别页面/模板和保留批注外观内容中的文本对象具有正的无限定 `Size`。不透明图形作用域为 Unknown；实际字形几何仍需外部证据。"),
        new ArchivalClauseDefinition("6.6(d)", "仅水平变换时优先使用 `HScale`", "未覆盖：精确的变换分解和外观。"),
        new ArchivalClauseDefinition("6.6(e)", "优先为重复文本颜色共享 DrawParam", "未覆盖：可选资源去重。"),
        new ArchivalClauseDefinition("6.7", "将视频视为带视频动作的图像", "未覆盖：媒体/动作处理和渲染。"),
        new ArchivalClauseDefinition("6.8", "优先采用适合复用范围的矢量资源表", "未覆盖：跨页和页内组合对象清单。"),
        new ArchivalClauseDefinition("6.9(a)", "对象/大纲/页面/文档不含外部资源动作", "**部分：**`AC-013` 会使保留且可识别的 Action/URI 声明失败，包括显式绝对 HTTP(S)、FTP 或 file 目标，或相对于显式绝对外部 `Base` 解析的相对目标；覆盖文档、页面/模板、大纲和可识别图形对象位置，包括保留批注外观。无法证明为外部的 URI 方案或资源支持动作、不透明位置以及完整动作/资源闭包均为 Unknown。没有移除处理器。"),
        new ArchivalClauseDefinition("6.9(b)", "优先保留已移除动作的摘要", "外部证据：转换前后的动作来源。"),
        new ArchivalClauseDefinition("6.10.1(a)(1)", "冻结影响外观的批注参数", "外部证据：原始/最终外观比较。"),
        new ArchivalClauseDefinition("6.10.1(a)(2)", "移除批注外观 PageBlock 嵌套", "**精确的只读结构检查：**`AC-012` 跟踪保留批注外观，并使嵌套的可识别 PageBlock 失败；不透明外观为 Unknown。没有扁平化处理器。"),
        new ArchivalClauseDefinition("6.10.1(a)(3)", "设置保留批注的 ReadOnly/NoZoom/NoRotate", "**精确的只读属性检查：**`AC-009` 跟踪可识别批注索引，要求每个保留 `Annot` 的三个标志均为 true；无法解析或外来作用域为 Unknown。其他批注处理子条款仍未覆盖。"),
        new ArchivalClauseDefinition("6.10.1(b)", "替代方案：将批注外观扁平化到页面", "外部证据：来源所有权、视觉等价性和信息损失记录。"),
        new ArchivalClauseDefinition("6.10.2", "单独存储归档后管理批注", "外部证据：档案生命周期和独立文件证据。"),
        new ArchivalClauseDefinition("6.11", "优先保留标签并闭合其引用内容", "部分覆盖：仅已知包边；完整标签内容闭包和来源仍未覆盖。"),
        new ArchivalClauseDefinition("6.12(a)", "移除扩展信息", "**部分：**`AC-014` 会使带直接文本目标的可识别 `Document/Extensions` 声明失败；空、嵌套或外来声明以及其他扩展位置仍为 Unknown。没有移除处理器。"),
        new ArchivalClauseDefinition("6.12(b)", "冻结扩展外观/打印效果", "外部证据：来源特定解释和视觉/打印比较。"),
        new ArchivalClauseDefinition("6.13.1(a)", "非去技术化签名的格式和合格 CA", "外部证据：GB/T 35275、CA 资质、证书链/吊销和可信时间。APF-006 仅负责清点。"),
        new ArchivalClauseDefinition("6.13.1(b)", "印章格式、注册和合格证书", "外部证据：GB/T 38540、注册和合格 CA 证据。"),
        new ArchivalClauseDefinition("6.13.1(c)", "档案机构接受长期证书状态服务", "外部证据：档案机构决策和服务连续性。"),
        new ArchivalClauseDefinition("6.13.2(a)(1)", "在保留批注中记录签署者/时间/摘要", "外部证据：经认证的来源签名和新批注来源。"),
        new ArchivalClauseDefinition("6.13.2(a)(2)", "移除保留批注的 SignRef", "未覆盖：只能作为经认证签名去技术化的一部分。"),
        new ArchivalClauseDefinition("6.13.2(b)", "不保留批注时，将签名批注扁平化", "外部证据：外观/所有权比较。"),
        new ArchivalClauseDefinition("6.13.2(c)", "去技术化后移除签名记录和目录", "未覆盖：已签名字节策略和完整引用闭包。R17 副本拒绝签名输入。"),
        new ArchivalClauseDefinition("6.13.3(a)(1)", "将印章外观转为批注，并记录签署者/证书/时间/摘要", "外部证据：经认证的印章和外观证据。"),
        new ArchivalClauseDefinition("6.13.3(a)(2)", "替代方案：以 Darken 将印章外观合成到页面", "外部证据：视觉检查和合成支持。"),
        new ArchivalClauseDefinition("6.13.3(b)", "去技术化后移除印章记录和目录", "未覆盖：必须先认证来源并闭合引用。"),
        new ArchivalClauseDefinition("6.14", "保留多版本信息", "未覆盖：需要转换前后字节和引用。"),
        new ArchivalClauseDefinition("6.15.1(a)", "优先保留文本格式附件", "外部证据：保留策略和实际内容类型。"),
        new ArchivalClauseDefinition("6.15.1(b)", "优先保留有归档文档记录的格式", "外部证据：文档和档案接收证据。"),
        new ArchivalClauseDefinition("6.15.1(c)", "优先保留符合条件的已签名附件", "外部证据：a/b、签名标准和信任验证。"),
        new ArchivalClauseDefinition("6.15.2(a)", "转换符合条件的已移除附件，并追加到主文档", "外部证据：附录 A 格式、转换保真度、顺序和页面检查。"),
        new ArchivalClauseDefinition("6.15.2(b)(1)", "优先将移除的音频/视频外部化并保留摘要", "外部证据：保留的摘要、描述和外部对象证据。"),
        new ArchivalClauseDefinition("6.15.2(b)(2)", "优先将超出查看需求的大图像外部化并保留摘要", "外部证据：大小/用途策略、摘要和描述。"),
        new ArchivalClauseDefinition("6.15.2(b)(3)", "优先将数据电子表格外部化并保留摘要", "外部证据：语义分类和信息损失。"),
        new ArchivalClauseDefinition("6.15.2(b)(4)", "优先将动画/3D 内容较多的网页/幻灯片外部化并保留摘要", "外部证据：内容检查和保存决策。"),
        new ArchivalClauseDefinition("6.16", "长期 OFD 不加密；归档前解密", "**部分：**当包可加载时，`AC-015` 会使保留的根 `Encryptions.xml` 标记失败，不受本地 APF 策略或根 XML 是否可用影响。没有通用加密发现、来源解密证据和明文证明时，未发现标记仍为 Unknown。APF-004 仍是独立的本地策略检查。"),
        new ArchivalClauseDefinition("7.1(1)", "写入端支持相关生成/转换和归档处理", "外部证据：产品能力和转换 E2E，不能只检查包。"),
        new ArchivalClauseDefinition("7.1(2)", "写入端遵循规定的处理", "外部证据：仍缺完整的条款级端到端证据。"),
        new ArchivalClauseDefinition("7.2(1)", "阅读端完整输出 OFD、检查格式并报告未处理项", "外部证据：独立阅读器及 GUI/日志行为。"),
        new ArchivalClauseDefinition("7.2(2)", "阅读端禁止编辑 OFD-A 并给出提示", "外部证据：目标阅读器 UI 和权限测试。"),
        new ArchivalClauseDefinition("A.1/1", "栅格图像", "未覆盖：图像转页面，以及带星号类别的判断。"),
        new ArchivalClauseDefinition("A.1/2", "文本", "未覆盖：编码和内容保真度。"),
        new ArchivalClauseDefinition("A.1/3", "文字处理", "未覆盖：DOC/DOCX/WPS/RTF/UOF/UOT 各不相同；现有 DOCX 转换器不代表完整附录覆盖。"),
        new ArchivalClauseDefinition("A.1/4", "电子表格", "未覆盖：工作表、公式、打印区域和带星号判断。"),
        new ArchivalClauseDefinition("A.1/5", "演示文稿", "未覆盖：备注/动画和带星号判断。"),
        new ArchivalClauseDefinition("A.1/6", "固定版式文档", "未覆盖：OFD/PDF/XPS/PS/CEB/SEP/AIP 互操作。"),
        new ArchivalClauseDefinition("A.1/7", "网页", "未覆盖：资源捕获和带星号判断。"),
        new ArchivalClauseDefinition("A.1/8", "音频", "未覆盖：AVS 转换和内容证据。"),
        new ArchivalClauseDefinition("A.1/9", "视频", "未覆盖：AVS 转换和内容证据。"),
        new ArchivalClauseDefinition("A.1/10", "CAD", "未覆盖：查看/输出是否为主要用途。"),
        new ArchivalClauseDefinition("A.1/11", "数据库", "未覆盖：按内容采用表单或图像处理。"),
        new ArchivalClauseDefinition("A.1/12", "电子邮件", "未覆盖：网页/表单处理以及标头/附件。"),
        new ArchivalClauseDefinition("A.1/13", "社交媒体", "未覆盖：按内容采用文本/图像/电子表格处理。"),
    });
    /// <summary>Undated references requiring an explicit edition and amendments in the target dossier.</summary>
    public static IReadOnlyList<string> ReferencedStandards { get; } = Array.AsReadOnly(new[]
    {
        "GB/T 13000", "GB 18030", "GB/T 18894", "GB/T 33190", "GB/T 35275", "GB/T 38540",
        "DA/T 46", "DA/T 47", "DA/T 54", "GB/T 20090"
    });
}
