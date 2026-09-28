# OFD → HTML 预览

**What to build:** 把多页 OFD 导出成单个 HTML，页内容以现有 SVG 嵌入，浏览器打开即可预览。不承诺页面文字全部可选择、可复制。

**Blocked by:** None (can start immediately)

**Priority:** P1

**Status:** ready-for-agent

- [ ] 多页 SVG 包进单 HTML，浏览器能翻看每一页
- [ ] 文档写明不承诺可选中全部文字
- [ ] CLI 或 API 可指定输出路径

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P1-09

## What to build

给不装 OFD 阅读器的人一条「打开就能看」的路径：用已经能出的自包含 SVG 页打进一份 HTML。这是预览，不是可编辑网页，也不保证 HTML 里的字都能划选。

## Acceptance criteria

- [ ] API（及 CLI，若已有转换命令风格）把多页样例打成一个 HTML 文件；用浏览器打开能看到每一页的图面
- [ ] 页间顺序与源包一致；图片/路径/文字在 SVG 里可见，不是空白页
- [ ] 文档与功能对照写明：预览用途，不承诺可选中全部文字
- [ ] 无隐私样例可复现

## Blocked by

- None (can start immediately)
