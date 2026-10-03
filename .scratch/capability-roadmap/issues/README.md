# 能力路线图 tickets

来源是 [docs/capability-roadmap.md](../../../docs/capability-roadmap.md)。一票一文件。文件编号是依赖顺序，不是开工顺序。

只领 `Status: ready-for-agent` 的票。`blocked` 要等「Blocked by」里的票完成。P2-06（证书加密）没有建票。

## 建议开工顺序

与路线图第 8 节一致。同一批次内可以并行。不要因为编号更小就先做 08–15。

1. **生成**：01，然后 16（P0 表格，被 01 阻塞）
2. **进出与工具**：02、03
3. **issue #3**：04、05
4. **无厂商密码上限**：11、12（不要接到默认验签）
5. **保真与检索**：21（被阻塞）、06、09
6. **其余**：按调用方需求抽，不默认全做

## 索引

| 票 | 优先级 | 状态 | 阻塞 | 路线图 |
| --- | --- | --- | --- | --- |
| [01 公开流式布局](01-public-flow-layout.md) | P0 | ready-for-agent | 无 | P0-01 |
| [16 公开表格](16-public-tables.md) | P0 | blocked | 01 | P0-02 |
| [17 区域占位](17-area-holders.md) | P0 | blocked | 01 | P0-03 |
| [02 图片进出](02-ofd-image-io.md) | P0 | ready-for-agent | 无 | P0-04、P0-05 |
| [03 文档工具](03-document-tools.md) | P0 | ready-for-agent | 无 | P0-06–P0-09 |
| [04 OfdGraphics](04-ofd-graphics-api.md) | P1 | implemented-validated-awaiting-final-review | 无 | P1-01 |
| [05 字体子集与复用](05-font-subset-and-reuse.md) | P1 | ready-for-agent | 无 | P1-02 |
| [20 布局 Canvas](20-layout-canvas-drawcontext.md) | P1 | blocked | 01、04 | P1-03 |
| [21 PDF 矢量](21-pdf-to-ofd-vectors.md) | P1 | blocked | 04、19 | P1-04 |
| [06 定位并替换](06-keyword-locate-and-replace.md) | P1 | ready-for-agent | 无 | P1-05、P1-06 |
| [07 附件增删](07-attachment-add-remove.md) | P1 | ready-for-agent | 无 | P1-07 |
| [18 纯文本 → OFD](18-plain-text-to-ofd.md) | P1 | blocked | 01 | P1-08 |
| [08 OFD → HTML](08-ofd-to-html.md) | P1 | ready-for-agent | 无 | P1-09 |
| [09 强类型对象](09-typed-object-model.md) | P1 | ready-for-agent | 无 | P1-10 |
| [10 多 DocBody](10-multi-docbody.md) | P1 | ready-for-agent | 无 | P1-11 |
| [19 SkiaSharp 适配](19-skiasharp-graphics-adapter.md) | P1 | blocked | 04 | P1-12 |
| [11 口令加解密](11-gmt0099-password-crypto.md) | P2 | ready-for-agent | 无 | P2-01 |
| [12 SES 自签](12-ses-self-sign-interop.md) | P2 | ready-for-agent | 无 | P2-02 |
| [22 骑缝外观](22-riding-seal-appearance.md) | P2 | blocked | 06 | P2-03 |
| [13 锁定签名](13-lock-and-continue-sign.md) | P2 | ready-for-agent | 无 | P2-04 |
| [14 OFD-A 检查器](14-ofd-a-checker.md) | P2 | ready-for-agent | 无 | P2-05 |
| [15 发布闸门](15-release-engineering-gates.md) | Q | ready-for-agent | 无 | Q-01–Q-05 |
