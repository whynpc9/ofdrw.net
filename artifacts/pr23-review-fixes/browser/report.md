# PR23 两项修复后 HTML 实际浏览器验收

冻结源码 `979a93e39dd8cec3fbb4bdd2121d301ea84b2e39`，清单SHA `6da5cc063ac6b67fbf3fb99fa5aa717dc248719bb6dd67a3b22447f9070d7976`。实际浏览器 `Chrome/152.0.7977.54`，独立 TaskSpace21。4文件 / 9页面出现项全部实际加载并查看本次原始完整截图，限定本清单 **PASS**。

| 文件 | 源页顺序 / 实看范围 | 本次截图 |
| --- | --- | --- |
| [native.html](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/pr23-review-fixes/visual/native.html) | 1,2 / 全部2出现项 | [1-native-full.png](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/pr23-review-fixes/browser/1-native-full.png) |
| [notice-short.html](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/pr23-review-fixes/visual/notice-short.html) | 1 / 全部1出现项 | [2-notice-short-full.png](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/pr23-review-fixes/browser/2-notice-short-full.png) |
| [mixed.html](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/pr23-review-fixes/visual/html-fixtures/mixed.html) | 3,1,3 / 全部3出现项 | [3-mixed-full.png](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/pr23-review-fixes/browser/3-mixed-full.png) |
| [package-api.html](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/pr23-review-fixes/html-package-output/package-api.html) | 3,1,3 / 全部3出现项 | [4-package-api-full.png](/Users/wanghongyi/.codex/worktrees/audit-followup-coordination/ofdrw.net/artifacts/pr23-review-fixes/browser/4-package-api-full.png) |

每个完整 file URL、实际标题、标签/源页顺序、viewBox/实际尺寸、clip局部目标、字体状态、真实请求、文件前后SHA与截图SHA均保存在 `observations.json`，截图原字节未改。

Native two occurrences: Chinese/English body, blue headings, italic subtitle, filled bordered table, local red emphasis and date visible.
Short notice one occurrence: two Chinese/English lines with iii WWW and literal <literal> & text visible.
Mixed occurrences 3,1,3: repeated source page3 consistent; red italic versus upright blue page1 text, blue/red images and triangular clip visible; injection strings remain literal.
Package API occurrences 3,1,3: Chinese Wi & <literal> preserved, page labels/order and repeated page3 visible.

所有出现项宽高比与viewBox之差小于0.000006；script节点0、on*属性0、外部网络请求0；DOM ID全局唯一，mixed clip均指向本section内定义。请求记录实际含file导航以及data字体/图片，可区分无外部请求与未采集网络。注入标题和正文仅作为字面文本；未触发其fetch。未重跑此前独立SVG/哨兵控制，未冒称本次运行。

## 边界及清理

本次只验收清单内notice-short.html。长版notice默认HTML节点预算拒绝是既有边界，本轮未上调预算、未绕过，也未将旧拒绝当作本次新测试；不声称长版notice HTML通过。未嵌入/不可用字体仍可能依赖宿主回退，本次可见正常不证明严格字体身份或任意文档保真。

功能/包门由父root独立负责。原生Preview结果在`../preview/report.md`；HTML验收未用截图或DOM代替Preview。`browser-cleanup.json`记录TaskSpace21成功finish({keep:[]})，仅本次独立空间关闭，用户TaskSpace20完全未操作。GUI租约由本持有者释放。

## 文件身份

| HTML | 前后相同的SHA256 | 截图SHA256 |
| --- | --- | --- |
| native.html | `e754a150d44b2c6e5dc6df001fc684b49ebea3efff5c09edc30e9df8ba571a5d` | `a004d09b274aa19be69e448adf25aa2de12db0ed311ca8e0671f166f8ad5d669` |
| notice-short.html | `64bdde3be37ab17a4105bed3148aba479b6ae8fcc61af183619cd90d8decc91d` | `d0e603d1da82e28da6ae812502763336d1e270dd5497ab918bdd3331ef04fe71` |
| mixed.html | `e231eca4f1662bbe7a67d2a5675a596b263b8289258762181052a5aac8d69614` | `81213750007cf6f0d15077a78b757082e9adef7acd47ff1c1524a1ca3ce2aeeb` |
| package-api.html | `e9cd1995aa0cb1d462d5f02993bdc42d4a6f2bbdc9f804700896a7b691d8440e` | `8c380e34ecb122e597f9c9580b8bebe10410afca4ccd2ba364917eb44e4b5e62` |
