# 最终集成 HTML 浏览器验收

源码：`aef66e76d923d706bcd440eb1583c906f300eeb7`。状态：**PASS，4 文件、9 个页出现项全部实际查看**。

实际浏览器：`Chrome/152.0.7977.54`，macOS，file URL。仅打开 root 已生成的清单文件，未重生 HTML，未运行 build/test/package，未修改源码、共享 plans、029 分支提交或原 029.json。

在 integration-preview-done.json 已存在、source 匹配且 gui_released=true 后才原子取得 GUI 租约；此前只做只读清单准备。[前置完成标记](preview-prerequisite.json)。

## 实际查看范围

| 文件/标题 | 页序与查看范围 | 截图 |
| --- | --- | --- |
| native.html / OFD preview | 源页 1,2；出现1–2全部查看 | [本次完整页面截图](1-native-full.png) |
| notice-short.html / OFD preview | 源页 1；出现1–1全部查看 | [本次完整页面截图](2-notice-short-full.png) |
| mixed.html / &lt;/title&gt;&lt;script&gt;fetch('http://127.0.0.1:9/attack')&lt;/script&gt;&" | 源页 3,1,3；出现1–3全部查看 | [本次完整页面截图](3-mixed-full.png) |
| package-api.html / Package &lt;literal&gt; & title | 源页 3,1,3；出现1–3全部查看 | [本次完整页面截图](4-package-api-full.png) |

每个精确 fileURL、浏览器实际标题/页序、范围、viewBox/实际尺寸、clip 目标闭包、字体加载、请求与 SHA 均在 [observations.json](observations.json)。四份截图均为本次真实浏览器完整页面捕获，覆盖表中全部9个出现项，未用旧截图冒充。

- 查看出现1、2；中英文正文、蓝色标题、灰色斜体、表格填充/边框与红色强调可见，分页标签和对齐正确，没有新增缺字、裁切或异常空白页。
- 查看出现1；PlainText→HTML 两行中英文、iii/WWW 比例文字及 <literal> & text. 均可见并作为字面文本；无裁切或重叠。
- 查看出现1、2、3（源页3、1、3）；红色斜体/普通文字区分、蓝/红图片、三角形 clip 和重复源页3一致；标题/正文注入文本保持字面，未生成脚本。
- 查看出现1、2、3（源页3、1、3）；真实 nupkg API 的 Package page 3/1/3 中文 Wi & <literal> 均完整可见，页标签、重复和物理页面比例正确。

所有实际页面宽高比与其 viewBox 比例之差小于0.0001。9个出现项的脚本节点0、事件属性0、外部请求0；DOM ID 全局唯一，mixed 的 clip 引用均指向本 section 的定义。没有发现本次样例新增的缺字、重叠、异常裁切或意外空白页。未内嵌或不可用字体仍依赖浏览器/语言/宿主回退，本次可见正常不等于严格字体身份或通用 Word/OFD 保真。

## 文件与截图身份

| 文件 | 原文件 SHA-256（查看前后相同） | 本次截图 SHA-256 |
| --- | --- | --- |
| native.html | `e754a150d44b2c6e5dc6df001fc684b49ebea3efff5c09edc30e9df8ba571a5d` | `a004d09b274aa19be69e448adf25aa2de12db0ed311ca8e0671f166f8ad5d669` |
| notice-short.html | `64bdde3be37ab17a4105bed3148aba479b6ae8fcc61af183619cd90d8decc91d` | `d0e603d1da82e28da6ae812502763336d1e270dd5497ab918bdd3331ef04fe71` |
| mixed.html | `e231eca4f1662bbe7a67d2a5675a596b263b8289258762181052a5aac8d69614` | `81213750007cf6f0d15077a78b757082e9adef7acd47ff1c1524a1ca3ce2aeeb` |
| package-api.html | `e9cd1995aa0cb1d462d5f02993bdc42d4a6f2bbdc9f804700896a7b691d8440e` | `8c380e34ecb122e597f9c9580b8bebe10410afca4ccd2ba364917eb44e4b5e62` |

## 网络与既有控制边界

本次 Network.requestWillBeSent 确实记录了本地 file 导航及需要的 data 字体/图片请求；四份文件均无外部请求。mixed 的 fetch 注入标题与正文保持字面文本，当前 DOM 也没有 script 或 on* 属性。请求记录保留在 observations.json，data URI 用类型前缀和完整 URI 的 SHA 记录。

按本轮授权，不重复搭建哨兵或重跑独立源 SVG。已有正向请求/零攻击哨兵与独立 SVG 字体控制来自实现源码 `9a67e2658e94b45d695fb383b1f10f91cef178b3`；经 git diff 核实该版本至最终 aef66 的整个 SVG 程序集源码无差异。[来源及等价性记录](source-verification.json)。这些既有控制不冒称为本次新运行；本次外部请求0由本次网络记录独立确认。

## 功能与预算边界

本轮只做浏览器验收，功能/包门由父 root 的联合验证负责，没有重复运行或扩大通过声明。Preview 由003另验，本轮没有操作其 PDF 窗口。

长版两页 notice.ofd 的默认 HTML 节点预算拒绝仍然成立，首次日志为 exit1 / HTML preview capacity exceeded: nodes。这里只验收清单内 notice-short.html；没有降低或绕过预算，也不声称长版 notice HTML 通过。[边界记录](known-budget-boundary.json)。

浏览器 TaskSpace19 已关闭，仅关闭本人创建的页面，用户空间保留；GUI 租约由 finally 释放。原029 HEAD 与029.json哈希保持不变。[关闭记录](browser-cleanup.json)。
