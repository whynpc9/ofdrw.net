# 04 原生绘图契约

设计基线：`df71c7f20e0c45e9cba9cb78f2d2a91d61413ec9`（PR #12）。Astra High 已完成源码设计审查。

公开类型位于 `Ofdrw.Net.Layout.Graphics`，输出复用 Core 的 `OfdPathElement`、`OfdTextElement` 与 `OfdTextRun`，由已有 Writer 分配 ID、写资源与失效签名清理。调用方使用 package/page 入口；不引入第二套文档或字体模型，不改 builder、Reader 公共 API。

- 所有坐标/线宽/字体 em 使用毫米，调用坐标原点为物理页框左上角，Y 向下。矩阵六值为 a,b,c,d,e,f；x'=a*x+c*y+e，y'=b*x+d*y+f。`Multiply` 是左乘右，右矩阵先执行；追加变换 `current=current*operation`。正角度顺时针。
- 路径 M/L/Q/B/C 为绝对命令；B 为 cubic、C 为 close。填充 NonZero 或 EvenOdd；纯色画笔/笔刷，butt cap、miter join、miter limit 10。线宽在用户空间，完整 affine CTM 同时变换笔画与几何。对象 Boundary 是变换控制多边形及笔画的保守包围框。
- `DrawString` y 为基线，单游程、原文镜像完整保留。无 advances 时使用阅读器字体字距；显式 advances 按已有 grapheme 规则，数量严格为 graphemeCount-1。文本 Boundary 是整物理页视口，无另建字体测量/换行服务；页外内容由物理页裁切。不保证复杂 shaping/任意 Word 保真。
- 图形路径在 draw/clip 时快照。clip 在设定时固定到页面坐标，每次交集单独写 Clip，输出时仅减 Boundary 原点，不能乘对象 CTM 逆。Save/Restore 保存变换和 clip，严格 LIFO；空栈失败，已输出对象不变。裁剪必须含可绘制段，空路径明确拒绝。
- 有限输入、派生乘法/边界、可序列化正尺寸、图元/路径/文字/状态/clip/累计几何预算和取消均在 append 前检查。每次绘制原子追加。上下文非线程安全，资源和 page 不应并发修改。

## 与 05 的字体绑定边界

`OfdFont` 仅描述 name/size/weight/italic/optional resourceId。显式 ID 是目标 `package.Fonts` 内 ordinal 逻辑句柄，必须恰好匹配一个有效命名资源；实际 FontName 从此资源读取，不允许静默 fallback。无 ID 沿用现有 Writer/Converter 的名称/风格选择。资源 Bold/Italic 描述字体文件，文本 Weight/Italic 描述请求的强调；现有服务合并两者。

04 不加载系统字体，不拥有字体字节、resolver、cache、subset 或 fallback 服务；样例调用方自行向 `package.Fonts` 注册许可明确的 OFL 字体。05 可按原始载荷内容复用并重映射序列化数字 ID，必须保持文本 face 绑定；不可按显示名称合并不同内容。04 不写 CGTransform/glyph ID，05 从实际 Runs 收集 Unicode，用字跨页合并；镜像 Text 不应重复累计。当前 Core 没有 TTC face index，不把 TTC 多 face 默认为已支持。

Graphics 将局部文字游程归一到 (0, size)，矩阵包含调用方 x/baseline 的平移。Writer 仅对新建 XML、已有 CTM、基线归一且无 DeltaY 的 name-only italic 组合 M*F；F=[1,0,-0.2,1,0.2*serializedSize,0]，只给 F 写 03 `FauxItalicMatrixV1`。PDF/SVG 使用既有去因子/anchor 逻辑，用户 shear 保留并且强调只应用一次。嵌入载荷/透明文字跳过；保留 SourceXml 和任意旧基线 CTM 不重新组合，序列化不修改模型。不扩大 Graphics 的整页 Boundary。

## 共享修改所有权

Core 新增内部 `OfdPathStyle` 与 `OfdNumericFormat`，分别读取已有 SourceXml 绘制参数及共享保真普通十进制；不增加公开模型。PDF 普通路径从平均线宽缩放改为完整 graphics CTM，clip 构建路径仍按 03 原有几何规则处理。SVG 普通路径输出 fill-rule 及已声明 cap/join/miter-limit。Writer 模型负责的新建 Text CTM、Path/Image CTM（含默认 image CTM）与 clip CTM 统一保真普通十进制；已有 Text SourceXml CTM 继续源 XML 所有权。Writer 对 fresh、normalized、name-only 的 CTM 文本组合明确生成的 faux 因子；已保留 XML 与资源绑定沿用原有所有权，Reader 无源码修改。此共享导出修改需全套回归、11 包消费 E2E 和本次实际页面验收。

可写矩阵验证按实际普通十进制系数，以有界 BigInteger significand/scale 精确比较 A*D 与 B*C，拒绝数学上退化的写出矩阵；不用 binary 浮点残差或 decimal 的28位范围。路径/clipCTM/DeltaX 使用共享普通十进制展开，保留 round-trip 数字且不输出 E exponent；不是把极小非零 advance 静默量化为零。Writer 生成的 F 和 M*F 使用相同保真普通十进制，避免再舍入造成基线漂移或退化。

Miter 映射核对固定 [上游 Graphics2D stroke 参数](https://github.com/ofdrw/ofdrw/blob/5fe9c4276c64e40b455e6ea649b695adf8a9a734/ofdrw-graphics2d/src/main/java/org/ofdrw/graphics2d/OFDGraphics2DDrawParam.java)，并新增尖角实际 PNG ink 断言。

R3：0.0004 缩放和 1.0004 近相关但非奇异矩阵不再受旧三位矩阵格式限制；实际写出读回保留系数。Boundary/Size/LineWidth 仍为原三位小数策略，文字 localBaseline 和归一化 translation 使用实际写出 Size，避免大倍数下漂移。新生成 M*F 再做精确奇异检查；不为历史 SourceXml/任意模型增加全局退化拒绝。窄页1000mm框线的0.9996同状态clip/PDF真实边界回归覆盖精度一致性。

SVG 导出矩阵、translation 与 text/tspan 原点也复用保真普通十进制。新增普通148mm页上的100000mm逻辑框线，通过负平移展示末端：object/clip同时0.9996，右边框在130mm；旧舍入会把边框移出40mm，实际PDF/SVG像素均做断言。MultiplyTransform 在运算前拒绝奇异operand，防止旋转的浮点残差把无效输入扰动成可写的非零det；之后再验证组合结果。包消费E2E直接从11包本地feed调用公开Graphics并检查原生CTM/文字/强调。

R4：现有Core.OfdTextEmphasis集中纯faux组合与finite/exact-singular检查；Graphics在Commit前通过同一OfdFontSelection获取effectiveItalic，并针对非嵌入/非透明的name-only预检。Writer复用该计算但保留既有兼容门槛。不复制/冻结字体，不新增字体服务，正常矩阵/提示/渲染数值不变。极大有限矩阵若faux组合溢出，当场拒绝且不消费图元/文字预算。
