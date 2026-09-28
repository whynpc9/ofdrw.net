# OFD 图片进出：栅格导出与一图一页导入

**What to build:** 指定页和分辨率，把 OFD 页导出为 PNG（默认）或 JPEG；反过来把 PNG/JPEG 做成一图一页、居中的 OFD。API 与 CLI 都能用。

**Blocked by:** None (can start immediately)

**Priority:** P0

**Status:** ready-for-agent

- [ ] `ofd-to-image`：指定页、指定 ppm，默认 PNG
- [ ] `image-to-ofd`：PNG/JPEG，可设页尺寸与 ppm，一图一页居中
- [ ] 有样例；导出图可用于目视核对页面内容

## Parent

[docs/capability-roadmap.md](../../../docs/capability-roadmap.md) P0-04、P0-05

## What to build

预览和简单制证现在卡在「只能出 SVG/PDF」。这张票打通图片进出：读包可按页出位图；给一张图可生成标准 OFD 页。栅格化可以走现有 SVG 或 PDF 页，不必新发明渲染器。页尺寸与 ppm 可配置。

## Acceptance criteria

- [ ] API 与 CLI 都能把指定 OFD 页写成 PNG 或 JPEG；未指定格式时为 PNG；可设 ppm
- [ ] API 与 CLI 都能把 PNG/JPEG 写成 OFD：每图一页、图像居中；可设页尺寸与 ppm
- [ ] 非法页号、不支持的输入格式有明确失败，不写出半包
- [ ] 有无隐私样例；导出的 PNG 能看出对应页的文字或图形，而不是空白或全黑
- [ ] 更新功能对照；CLI 帮助包含这两条命令

## Blocked by

- None (can start immediately)
