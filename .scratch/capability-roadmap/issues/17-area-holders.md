# 区域占位与事后回填

**What to build:** 先在版面上放命名区域，分页完成后再填文字或图片，不打乱已排好的页。

**Blocked by:** 01 公开流式布局：Paragraph / Span、折行、分页

**Status:** ready-for-agent

- [ ] 命名区域可先占位再回填文字/图片
- [ ] 回填不改变已分页的页数与其他块位置（按样例约定）
- [ ] 有样例

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P0-03

## What to build

表单和套打需要「先留空再填」。调用方声明命名区域，布局按占位尺寸分页；之后往区域里填字或图，而不重新排整篇。这不是交互式 PDF 表单，只是生成期回填。

## Acceptance criteria

- [ ] API 能声明命名区域并生成带占位的多页 OFD；再填文字或图片后保存
- [ ] 样例上回填前后页数不变，区域外内容坐标不变
- [ ] 超长回填有文档化行为（裁切、缩小或明确失败），不静默覆盖邻块
- [ ] 更新功能对照；有样例与 Preview 抽查

## Blocked by

- 01 公开流式布局：Paragraph / Span、折行、分页
