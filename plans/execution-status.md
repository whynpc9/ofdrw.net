# 计划执行调度记录

开始：2026-10-15。授权：全部计划顺序执行；每计划独立 `gpt-6.1-sol / high` subagent；文件不冲突时并行。协调者维护索引、复核源码和验收，不直接修改生产代码。

源码起点：`d0f88c7321a57443a4cb0ec92631664f66ed8cc0`。原 checkout 保留既有未跟踪 `.agents/`、`.claude/`、`skills-lock.json`；所有实现位于独立受管 worktree。本轮授权本地 worktree 提交用于交接，不推送/发布或合入用户 main。

## 调度约定

- 每份计划启动前检查依赖和代码漂移；后续叠加已经复核的依赖实现，记录真实实施基线。
- 源码文件冲突的任务串行整合；独立 worktree 防止构建/产物互相污染。
- 各执行者记录实际测试/TRX，协调者检查完整差异、断言与复跑结果。
- macOS Preview 由协调者分配独占租约；执行者先生成本次实际 PDF/PNG，再请求验收。任何时刻仅一位获租约执行者操作 GUI。
- 原生 Windows 门缺实际执行证据时明确待验收，不把 macOS 或工作流编辑写成 Windows 已通过。
- 状态区分实现完成、独立复核、功能门、PNG、Preview、原生平台；缺项不能标整个计划 DONE。

## 当前任务

| 计划 | 执行者 | worktree | 状态 |
|---|---|---|---|
| 001 | execute_plan_001 | `/Users/wanghongyi/.codex/worktrees/ofdrw-plan-001/ofdrw.net` | DONE / `36468a50` |
| 002 | execute_plan_002 | `/Users/wanghongyi/.codex/worktrees/ofdrw-plan-002/ofdrw.net` | IN PROGRESS |
| 003 | execute_plan_003 | `/Users/wanghongyi/.codex/worktrees/ofdrw-plan-003/ofdrw.net` | IN PROGRESS |
| 004 | execute_plan_004 | `/Users/wanghongyi/.codex/worktrees/ofdrw-plan-004/ofdrw.net` | IN PROGRESS |

| 013 | execute_plan_013 | `/Users/wanghongyi/.codex/worktrees/ofdrw-plan-013/ofdrw.net` | IN PROGRESS |
| 026 | execute_plan_026 | `/Users/wanghongyi/.codex/worktrees/ofdrw-plan-026/ofdrw.net` | IN PROGRESS |
| 027 | execute_plan_027 | `/Users/wanghongyi/.codex/worktrees/ofdrw-plan-027/ofdrw.net` | IN PROGRESS |

其他计划按依赖与共享文件释放顺序启动；调度持续进行。

## 2026-10-15 独立复核更新

- 001：完整diff和新断言已复核；协调者独立 Signatures 13/13、CLI78/78、public-api文档Rebuild通过。候选 `36468a50ef7c2c5005a8ddd114253814b5c14caf`；仅此计划标DONE。
- 002：冻结 `010db666e80b869ba02885d733150653f163d0d3`；图像根路径4回归与包消费通过，协调者Preview处理中，暂未DONE。
- 003：冻结 `156541ec4523e248425c43fef4e386b0869b7ea5`；协调者完整diff和独立10回归通过，Preview待检查。
- 004：冻结 `efa8b7f6a36c92965da3695346449555f677f811`；本地功能完成，已发现可用的Legion WSL2内原生Windows .NET，执行者正补Windows实际门。Preview待检查。
- 005：在003冻结提交上执行，保留候选校验/一次提交边界；006在004冻结提交上执行；007在002冻结提交上执行。代码基线漂移已核对，未回退前序修复。
- 011：独立预算/字体任务进行中；013逐页包消费任务进行中；026文档候选 `ec121b087e095db4b4325cd4d61e61cbb4372142` 等协调者复核；027调查已确认描述符部分未知扩展信息丢失，结果待复核。
- 013校准：协调者复查 Test第5页源/往返全页与裁剪，内容/换行/位置保持，观察到栅格化边缘软化；同意以实际11页最大RMSE .0942145为依据使用.10阈值，必须有同尺寸非空负对照被同一比较器拒绝，不声称像素相等。

## 最新冻结候选与协调状态

此表按2026-10-16当前调度更新；机器清单见 [execution-state.json](execution-state.json)。下方段落保留发生顺序的历史，不覆盖本表。

| 计划 | 冻结提交 | 当前状态 |
|---|---|---|
| 001 | `36468a50` | DONE |
| 002 | `010db666` | DONE（独立复核及12页Preview） |
| 003 | `86adab14` | IMPLEMENTED; actual Preview7 PARTIAL/FAIL pre-existing direct subtitle italic |
| 004 | `d3b0a705` | DONE（Preview通过；Windows链接7项跳过；JPEG证据扩展名已纠正） |
| 005 | `34245397` | DONE（独立18项及12页Preview） |
| 006 | `85fbd17c` | DONE |
| 007 | `1c52e7d8` | DONE |
| 008 | `20f9e84d` | DONE（独立17项回归通过） |
| 009 | `2782cf5` | DONE（独立30项、Preview28页） |
| 010 | `5f4e9372` | DONE（独立28项及12页Preview；最终014兼容123项与11包通过） |
| 011 | `cac33dee` | DONE（独立40项、Preview14页、SVG9图；哈希URL复核通过） |
| 012 | `a0fa1da3` | DONE（独立验签40项、CLI78项） |
| 013 | `8b36416a` | DONE（独立包E2E、Preview11页；保留raster软化） |
| 014 | `4790e5c` | DONE（root独立97项、Preview9页、SVG2图） |
| 015 | `c2db7ee2` | DONE（独立37项、Preview15页、SVG9图；source identity分开） |
| 016 | `d9e68a39` | DONE（独立15项、Preview11页） |
| 017 | `ce991158` | DONE（独立21项、Preview12页） |
| 018 | `555b2bb4` | DONE（独立15项及9页Preview） |
| 019 | `5c53cfe5` | DONE（独立10项、Preview11页） |
| 020 | `d0ad0f61` | DONE（独立CLI128、Preview5页、SVG） |
| 021 | `d4958451` | DONE（独立CLI169、Preview4页） |
| 022 | `1d7e8ff7` | DONE（独立库4项、CLI2项） |
| 023 | `80e7b50` | DONE（独立9项、Preview7页） |
| 024 | `c5520dd3` | DONE（独立15项；三平台实测） |
| 025 | `90eba68e` | DONE（独立23项、Preview7页、SVG） |
| 026 | `ec121b08` | DONE |
| 027 | `19e58deb` | DONE（调查CONFIRMED；修复未实施） |
| 028 | `e35fb82e` | DONE（仅设计；独立26链接检查） |
| 029 | `1ca3e287` | DONE（仅设计；11链接复核） |
| 030 | `fe01184f` | DONE（仅设计；19链接复核） |

006 协调者完整diff/断言复核及 CLI 115/115；007 完整diff/断言及独立 Packaging 6/6、Signatures 8/8通过。026 完整文档与151链接检查通过。011第一轮复核发现新预算在3处选项克隆及来源指纹遗漏，已修复并重跑全部门，视觉排队。004原生Windows实际Core2、Packaging23、CLI14通过，7符号链接用例权限不足跳过，保留限制。

- 新独立验收：008完整source/test diff、17/17；012完整source/test diff、Signatures40/40与CLI78/78；027完整报告/探针复核，独立隔离复跑7case得到99MATCH/30LOSS/5UNDECIDABLE，监听控制可用；028六节设计契约人工复核、26链接检查通过，分别标DONE且保留调查/设计边界。
- Preview队列改由协调者授予单人独占租约，当前002执行者复查本计划12页，完成关闭后释放；root同步做源码和独立回归复核。旧root截图中repeated-page1实际为PNG窗口，不计PDF验收，002须逐份新开实际PDF重新检查。
- 014批准最小scope扩充Core/Models/OfdCompositeProjection.cs：现有admission后、清SourceXml前在clone用共享有效文本视图，原definitions/身份不变，新增composite两renderer控制。
- 024第一轮复核要求32位Unix明确failclosed、构造失败确保释放/保留原异常/最终reap；原生三平台15项初稿通过后仍须对最终revision重跑。

## 后续复核与租约

002已12页实际PDF Preview完成，root核对12份AX URL/页码与哈希并抽查重复双图/选页截图，标DONE。018全部2448零skip、可信65、包11及9页Preview完成，root完整源码/断言和独立15项通过，标DONE；019可启动。024最终c5520dd三原生平台15/15，root完整两轮源码复核及独立mac15/15，标DONE，保留32位Unix明确不支持/旧libc未测边界。029/030六节契约人工复核与11/19链接独立检查通过，仅设计DONE。

003实际Preview已7页，但直接BuiltIn副标题斜体缺失；已用原始d0f88c7重渲染证明两个页面与当前PNG字节相同，为既有视觉缺陷。仍不报全视觉通过，执行者只读定位最小修复提案。014复核发现新Writer往返literal斜杠风险并修正；初修709改变Extractor既有direct-XText契约，经既有测试失败后7763fd6拆LogicalText与严格render准入、加编码前预算/取消。所有新增scope经root核对批准，raw/source/owned身份保持；010/015必须接续最终修订再做本次最终门。

Preview目前004独占；018/003/002已释放。后续队列005→011→013→020→021→023→009及新计划，根据依赖释放调度。

## 2026-10-16 协调更新

004两页、005十二页Preview实际验收完成；004 Windows7符号链接权限skip保留，状态显示该限制。010图片准入自身source3a966经root全diff/tests、独立Pdf18+Docx10及12页AX/URL/hash核对、功能2505零skip/11包/PNG/Preview完成，允许016/017启动；root查出截图实为JPEG，执行者证据-only改正确扩展，bytes/SHA不变。010补继承014最新delta的兼容测试另录，不覆盖source355旧视觉证据。

014真实默认Auto字体子集回读重存新增4失败与decoded coverage正反2失败，批准仅FontEmbeddingPlan加入logical effective view并共享material；继有精确2参ctor的reflection取消测试又证实optional第3参不兼容，479保留2参委托3参。Root读全delta，最终479全套2411零skip/trusted65/docs/11包均过，PNG9+SVG2页已实际看过，Preview9页排队。010/015/019只做必要继承兼容和本次产物，旧whole门保留源码归属，最后全候选集成统一全面验证，避免重复中间失败头全跑。

022完整diff/断言与独立库4+CLI2通过，标DONE。011独立Core14/Packaging9/Pdf12/Docx5各分项目TRX通过，Preview/SVG租约当前归011。016和017分别在010+014最终源码上隔离执行；019字段最终分页source11b92已root完整review。025仍等待014完整视觉硬前置。


### 2026-10-16 最终桌面与集成调度

011 完成实际 Preview 14 页及 SVG 浏览器 9 图，root 核验全部14条 PDF/AX/截图SHA及fileURL，JPEG扩展名已纯重命名纠正。014 获独占Preview租约，root 独立 Audit014 36/36、Packaging 61/61通过。016、017全部source/test差异已读且无修订；016最终全套2526零跳过、PDF468通过，root独立15项启动；017最终全套2532零跳过，包/docs串行中。010最终014兼容123项及新11包通过，015独立37项与最终兼容132项通过，019独立10项通过。020 SVG浏览器补验通过且Ego租约释放。独立sol high integration_candidate 启动，仅隔离候选整合并保留每计划来源；025等待014完成后最后接入。主checkout生产源码未修改，未push/PR/merge/publish。

014 最终6PDF9页实际Preview通过，root复核9条真实URL/页码、6PDF原SHA一致并抽查literal反斜杠页面，标DONE；025独立sol high已在新worktree基于016+015启动。013获下一独占Preview租约。016独立15/15、017独立21/21均通过。综合候选ef6e9136的36项unique commits经root逐项stable patch-id比较，35项与原交付相同；026唯一feature-parity冲突保留010准入边界和026来源/依赖证据，已人工复核；候选尚缺025。

013 Test/Test3/ff共11页实际Preview完成，root核验全部11条URL/页码/JPEG及3原PDF哈希，并抽查第5页，保留栅格软化，标DONE。020接续5页Preview。综合候选pre025头ef6e9136全套2789/2789零跳过、trusted65、Python20、docs及11包全部通过；这些预验证不冒充最终含025验证。

020本次Preview5页通过，root全部5条URL/页码/JPEG与3PDF原SHA核验且抽查选页，标DONE；021获得4页租约。020观察到013遗留精确ff PDF，已要求013纠正仅聚焦用户窗口不足以证明全部关闭的报告，并授权021仅在fresh完整URL匹配时清这个own产物，不碰用户窗口。025 firstsource4c72c5d8核心diff/tests已读，revision1只补真实element内部取消确定性验证，保留原精确重载。

021全部4页Preview完成并由root核验真实URL/1page/原PDF与原始JPEG哈希，标DONE；013遗留ff经021fresh精确URL匹配后关闭且用户4文档窗口保留，旧关闭报告已纠正。009当前28页Preview租约。025最终ee325773二审通过，14+9专项通过，补真实scalar1后同步取消；Ego1SVG租约已授。综合候选b40e9c1接入全部38项唯一提交，025两项stable patch IDs独立比对一致，无源码冲突；最终完整门运行中。

最终综合头b40e9c1完整功能2812/2812零跳过、可信65、Python20、API docs与11包及013逐页比较全部通过；root独立六项目516/516零跳过通过（Core17/Packaging113/Signatures40/Pdf127/Docx124/CLI95）。本头13PDF20页PNG和2SVG实际查看通过，root抽查字段第2页和nested路径，核对全20页PDF/PNG及2SVG哈希一致，Preview仍排队，未冒充完成。025 root独立14+9=23通过；其全部功能、PNG7页、SVG已过，Preview待验。015实际9路径页已通过，正在收尾Native/default/CLI6页。

015实际15页Preview完成，root全部15条实际URL/页码、原PDF及JPEG哈希核验，保留path9页source2d74和最终DOCX6页c2db独立来源，标DONE；016获得下一11页Preview租约。全部Ego已结束，后续不再有跨应用GUI并行。

016本次7PDF11页Preview全部通过并关闭ownPDF，root全部11条实际URL/页码/原PDF与原JPEG哈希核验，标DONE；017接续7PDF12页Preview。源d9e68a3未改，不以测试代替页面。

017实际7PDF12页Preview全部完成，root全部12条完整URL/页码/原PDF、AX及JPEG哈希核验，标DONE；原JPEG未改，AX行尾规范化单独记录，不改变目标URL/页码。019当前11页租约，先行4页确认9/3/2、10/3/2、20/3/1。

019实际5PDF11页Preview完成，root全部11条URL/页码/原PDF/AX与JPEG哈希核验，字段四位置及selected数值正确，标DONE；023当前7页Preview租约。

023实际5PDF7页Preview全部完成，root全部7条URL/页码/原PDF、AX及JPEG哈希核验，标DONE；025接续最后一项计划7页Preview，其后综合候选20页。

025实际5PDF7页Preview完成，root全部7条真实URL/页码/原PDF、AX及原JPEG哈希核验，标DONE。全部30计划已分别执行：29项完成其修复/验证/调查/设计范围验收；003实现完成但保留基线既有直接BuiltIn副标题斜体视觉失败。最后综合候选20页Preview当前由integration_candidate独占验收，尚未宣称该门完成。


### 最终交付状态（2026-10-16）

30份计划已全部分别由独立sol/high执行者处理，29项完成各自修复、验证、调查或设计范围验收，003实现完成但保留既有直接BuiltIn副标题斜体视觉失败。上方“最新冻结候选与协调状态”表已按最终机器清单刷新；较早段落中的排队/租约状态仅为历史。

本地综合分支`codex/audit-integration`源码冻结于`b40e9c1dc5011ab08979e10e2bbbc9766b855770`，38项唯一提交集成映射已独立复核。最终全套2812/2812零失败零跳过、可信字体65/65、Python20/20、API文档、11包消费及013逐页栅格比较通过；协调者独立六项目516/516通过。

最终本次生成13PDF/20页：自动渲染、20页PNG目视、2SVG实际浏览器及20页macOS Preview均通过。协调者独立核验20条完整URL/页码AX、原始JPEG格式及SHA、13PDF及20PNG哈希，并抽查页码域第二页和native正文首页截图。13份任务PDF全部关闭，用户原四文档窗口保留，全部GUI租约释放。结论仅覆盖记录样例；003既有直接PDF斜体失败不因native链路通过而豁免。

[最终验收记录](/Users/wanghongyi/.codex/worktrees/audit-integration/ofdrw.net/artifacts/integration/final/acceptance.md) · [逐页Preview证据](/Users/wanghongyi/.codex/worktrees/audit-integration/ofdrw.net/artifacts/integration/final/preview/review.json)。004保留Windows链接7项权限跳过；027仅调查确认、修复未实施；028–030仅设计。主checkout生产源码保持原基线，未推送、建PR、合入main、打tag或发布。各独立worktree及综合候选保留供复查。
