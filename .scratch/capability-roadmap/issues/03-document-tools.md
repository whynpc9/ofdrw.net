# 文档工具：水印、按页拆分、Mix、签名清理

**What to build:** 在已有 OFD 上完成四件日常工具事：加水印、按页拆成新包、多页叠成一页、清掉签名声明和外观。合并与导出后水印还在；清签后验签报告是「无签名」，不是残章。

**Blocked by:** None (can start immediately)

**Priority:** P0

**Status:** implementation-in-progress (functional regressions passed; Preview and PR review pending)

- [ ] 文字/图片水印写入指定图层；合并与导出保留
- [ ] `split`：按页生成新包；资源与失效签名按现有合并约定处理
- [ ] Mix：多页叠成一页，页尺寸以第一页为准；图层顺序有文档
- [ ] 签名清理后 `verify-signatures` 报告无签名声明

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P0-06、P0-07、P0-08、P0-09

## What to build

编辑侧已有重排、删页、裁剪、自包含合并。这张票补齐上游工具里还缺的四件：水印、拆分、页面混合、签名清理。每件都要有 API 和 CLI，并且单独可演示。

水印是图层上的文字或图片，不是签章。Mix 把模板和注释一并叠上去。拆分沿用合并的资源拷贝与「字节变了就丢掉失效签名」约定。清签要删签名列表、签名值和外观，避免阅读器还显示残章。

## Acceptance criteria

- [ ] 可对指定页写入文字水印和图片水印；随后合并、导出 PDF/SVG 仍能看到水印；有样例
- [ ] 可按页码列表拆出新 OFD；未选中页的私有资源不残留；源包带签且拆分改变字节时，输出按现有约定处理失效签名
- [ ] 可将多个源页叠成一页；输出页尺寸等于第一页；图层（含模板、注释）顺序写在文档里
- [ ] 独立清签 API/CLI 删除签名列表、签名值与外观；之后 `verify-signatures` 报告无签名声明，而不是带着残缺签名对象
- [ ] 更新功能对照与 CLI 帮助

## Blocked by

- None (can start immediately)

## Implementation evidence

- API/CLI 契约：[工具教程](../../../docs/tutorials/15-document-tools.md)。
- 可复现四项 API/CLI 样例：[DocumentTools E2E](../../../e2e/Ofdrw.Net.DocumentTools.E2E/README.md)。
- 持久产物、完整性清单与分层验收：[证据](../../../docs/evidence/document-tools/README.md)。
- 完整视觉门及最新 PR review 未闭合前，本票保持未完成。
