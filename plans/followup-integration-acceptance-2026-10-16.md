# 五项后续实施统一集成验收

> Storage update: historical `artifacts/` paths below are restored from the
> [optional PR23 evidence bundle](../docs/evidence/pr23-integration/README.md).
> The original bytes and acceptance scope are unchanged.

共同起点：`3247c76f95c9a3312cc26144e24c4984767af26f`。目标分支：`codex/audit-integration`。已测试源码：`aef66e76d923d706bcd440eb1583c906f300eeb7`。

五项均由独立sol/high新会话、独立受管worktree完成，协调者先完整复核后逐一使用merge --no-ff合入。004按用户要求排除。所有五个最终交付HEAD均为当前集成HEAD祖先。唯一源码冲突是CLI Program.cs的命令列表及分发，同时保留text-to-ofd与ofd-to-html；其他各任务src/tests/e2e/scripts文件逐一与原交付核对一致。

| 任务 | 最终交付HEAD | 实际交付 |
|---|---|---|
| 003 | `0fda098f43d6662cfdf2fde1a4ba5b2e3f00df8b` | 直接BuiltIn PDF真实斜体；保留原事务修复，未升级依赖 |
| 027 | `470de1ebb98b2f9cf1448e0dfbc1763e527b3913` | 同文档未知附件XML来源包络保留、显式删除和受控克隆 |
| 028 | `8a37f4e23c982dd387f9d467f1898429e8ce0b9b` | 纯文本Flow API及text-to-ofd CLI |
| 029 | `924a9130d86d656c401e786117ec7814e06369d8` | 一次读包、有界静态HTML API及ofd-to-html CLI |
| 030 | `5e565eec31aff4d7374bf528571a5adfedee4ea3` | 六条只读本地APF四态预检，非标准认证 |

## 合并后实际功能门

| 门 | 结果 |
|---|---|
| 全套 | 3039/3039，11份TRX，0失败/0跳过 |
| 五项专项在全套中的实际执行 | 003 17、027 38、028 56、029 45、030 81，共237；其中003含原事务10项，不重复加到总数 |
| 独立分支复核复跑 | 上述237专项均由root在各交付源独立复跑通过 |
| 可信字体 | 仓库固定离线字体；强制65/65，零跳过 |
| Python | 20/20 |
| API文档Rebuild | 通过，新增未文档化API为0，原272历史符号未扩张 |
| 本地包 | 11个新候选包真实消费；新增PlainText string/UTF8 API和安装CLI通过 |
| HTML包消费 | 隔离nupkg API和安装CLI，重复页3/1/3与转义通过 |
| 预检包消费 | fresh cache、纯PackageReference，4例四态/只读控制通过 |

本地包版本`0.0.0-followup.20261016.1`，没有上传。所有dotnet使用可写CLI_HOME、禁首次体验/遥测、单节点与禁build server/shared compilation flags。restore先于脚本和包验证；文档Rebuild与测试/pack串行，整套门按共享validation锁执行并释放，不清全局NuGet缓存。

完整日志、实际命令/TRX、环境、原提交映射、专项计数、包产物位于本worktree的`artifacts/followup-integration/`。已将原始TRX、关键日志、页面产物/PNG/Preview原JPEG与AX、浏览器原截图及记录按显式清单纳入交接提交；NuGet缓存、11包文件、运行时目录和字体payload仍保留本地供复查，不递归提交。notice原始BOM/CRLF及尾随空格以notice-source.json的raw_utf8无损保留。`gates.json`、`test-summary.json`、`followup-case-counts.json`和`merge-map.json`记录实际源码归属。

## 合并后实际视觉门

自动渲染：7份PDF共12页已重新生成PNG，通过。PNG实际目视：12/12页通过。macOS Preview：7份本次PDF共12/12页通过，分别为native/default/CLI各2页、direct BuiltIn2页、选页1页、纯文本notice2页、真实italic边界1页。原始AX完整URL/页码与原JPEG保留，7PDF前后SHA均一致，本轮PDF已关闭，用户窗口保留。root核对全12条URL/页码/原JPEG魔数与SHA、12PNG及7PDF冻结SHA，并抽查新的direct副标题与italic边界显示。

浏览器：4份本次HTML、9个页面出现项全部实际查看通过（native2、纯文本短样例1、mixed3/1/3、真实包API3/1/3）。本次DOM脚本/事件属性/外部请求均为0，重复页ID唯一、clip引用均属本页，物理比例符合viewBox，HTML查看前后SHA一致。root核验所有精确URL/标题/页序/DOM与网络记录、4原HTML及截图SHA，并抽查native与mixed整页截图；浏览器任务空间已关闭、GUI租约已释放。[浏览器报告](../docs/evidence/pr23-integration/README.md)与[协调者复核](../docs/evidence/pr23-integration/README.md)可复查。冻结清单中初始PENDING字段保持原样，以避免改变被验收产物清单的SHA；最终状态以本验收记录和各实际执行记录为准。

实际路径、链路、观察、PDF前后哈希及关闭证据见[Preview报告](../docs/evidence/pr23-integration/README.md)、[PNG逐页记录](../docs/evidence/pr23-integration/README.md)、[冻结清单](../docs/evidence/pr23-integration/README.md)。直接BuiltIn DOCX→PDF单列，native/default均为本次DOCX→native OFD→PDF→Preview；纯文本为text→native OFD→PDF→Preview。未用直接DOCX→PDF代替native产物。

本次集成样例没有重新覆盖所有图片/页眉页脚/页码或复杂横向源文档场景；003分支自身有独立控制和报告，本次不将历史页面冒充新集成页面。

## 保留范围

- 004本轮排除，其Windows符号链接7项权限跳过未补验。
- 003仅对italic请求由regular物理字体承载的绘制做局部剪切，真实italic仅绘制一次，不扩为bold模拟、字体重新命名/嵌入或依赖升级；原候选回退/最终提交失败不回退边界保留。后续确认的既有DOCX source landscape PageSetup尺寸二次交换未扩修；prepared实际landscape绘制保持测试通过，不把它当完整DOCX landscape正确。
- 两页notice能生成native OFD/PDF，但默认HTML保守节点预算明确拒绝；首次失败日志保留于visual-generation-first-attempt.json，不提高预算或冒充该文档HTML通过。预算内短样例另查跨入口显示。
- 027不做物理擦除；不透明跨文档迁移和无法证明安全的来源迁移明确拒绝，未猜测重映射。
- 028只新增纯文本Flow适配入口和严格UTF-8流入口，沿用portable metrics/name-only字体及现有Flow分页，非富文本、真实字形测量或逐字节可逆文本抽取。
- 029首文档静态范围、保守媒体格式/预算及宿主字体回退边界不变；HTML不执行脚本、不抓外链，重复页独立ID及累计像素/节点/输出预算均保留。只接纳受限PNG与baseline Huffman JPEG，不把未支持的索引PNG/通用图片解码器或任意OFD图元算作支持。
- 030只检查本地有限条件，不检验字体合法性/覆盖/许可、不验签/解密、不作OFD-H或归档认证。
- 视觉结论只覆盖实际检查的确定性样例与页面，不推断任意Word/OFD均保真。

主checkout的main及既有未跟踪文件未改。本次后续合并仅本地，未推送、建PR、合入main、tag或发布。最终验收文档提交若在已测试源码之后，不含生产代码变更；验收源码仍以上述冻结SHA为准。

## 复跑与交接

实际完整命令及退出码见[gates.json](../docs/evidence/pr23-integration/README.md)，功能runner为[run-gates.py](../docs/evidence/pr23-integration/README.md)，产物runner为[generate-visual.py](../docs/evidence/pr23-integration/README.md)。重跑需按约定取得validation/GUI租约并用新目录保存证据；任何产物重新生成后必须重新完成适用视觉验收。本次验收未重生成冻结后产物。

五个独立会话及受管worktree均保留；所有最终交付HEAD为目标分支祖先。最终提交只补协调文档与可复查证据，生产/测试/e2e/scripts与已验证源码aef66完全相同。004原生权限门及上述已知边界均不是本次完成声明的一部分。
