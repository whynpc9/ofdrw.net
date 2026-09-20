# 公开流式布局：Paragraph / Span、折行、分页

**What to build:** 调用方不用手写毫米坐标，就能用公开的 `Paragraph` / `Span` 生成多页中英文 OFD；折行和分页由布局引擎完成。DOCX Native 仍走同一套引擎，行为不回退。

**Blocked by:** None (can start immediately)

**Status:** ready-for-agent

- [ ] 公开 Layout API 能排出带局部粗体/斜体/颜色的中英段落，并自动折行、分页
- [ ] 有可运行样例；Preview 视觉验收通过（Native OFD → PDF → Preview）
- [ ] 更新功能对照与教程实现对照

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P0-01

## What to build

今天生成 OFD 仍要往页面上堆 `TextObject`。这张票交付公开流式布局：调用方声明段落和行内样式，库负责折行、基线和分页，写出合法多页包。先把 DOCX BuiltIn 已有的折行/分页抽成 Layout 内部能力，再露出 `Paragraph` / `Span`；Native DOCX 继续用它，而不是再养一套排版。

不承诺任意 Word 浮动对象或复杂域保真。

## Acceptance criteria

- [ ] 不写页面坐标即可生成至少两页的中英混排段落；局部粗体、斜体、颜色只作用在对应 Span，不扩散到整段
- [ ] 中文不缺字乱码，英文间距与断词看起来正常；分页不裁切半行、不留下异常空白页
- [ ] 生成的 OFD 可抽出完整原文；视觉验收走本次 Native OFD，不得用直接 DOCX→PDF 代替
- [ ] DOCX Native/BuiltIn 既有转换测试仍通过
- [ ] 更新 `feature-parity.md` 状态，并在教程里对照公开布局 API
- [ ] 留下样例、产物和 Preview 验收记录（检查了哪些页、链路、遗留问题）

## Blocked by

- None (can start immediately)
