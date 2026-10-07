# 独立验证

gpt-6-sol low 子代理完成仅读源码审查及独立真实NuGet消费，未修改实现，未做GUI。
首次仅读指出验收README尚未落盘；该记录已补齐。未发现阻断性实现缺陷。

实际消费路径 `/private/tmp/ofdrw-issue11-independent/consumer`，仅
`PackageReference Ofdrw.Net.Crypto.Password 0.1.0-issue11.local`、net10.0，
无ProjectReference；独立 `NUGET_PACKAGES=/private/tmp/ofdrw-issue11-independent/packages`。
NuGet.Config将Ofdrw.Net.*仅映射本票三包feed，BC2.6.2只映射第三方缓存。
最终隔离缓存只含Core/Packaging/Password/BC四个包ID。

三nupkg的尺寸/SHA256及nuspec依赖与package-manifest完全一致：
Password→Packaging+BC2.6.2，Packaging→Core，Core无依赖。
构建并直接执行Consumer.dll完成native/default×partial/all四组：原始8条目名称/
载荷字节精确恢复、错误口令和预取消保持旧输出、无.ofd-password-*.tmp残留。
结果见 independent-functional.json。

命令采用dotnet restore/build --disable-build-servers -m:1 /nodeReuse:false
/p:UseSharedCompilation=false；运行dotnet绝对DLL路径避免run参数转发。
env为可写DOTNET_CLI_HOME、SKIP_FIRST_TIME_EXPERIENCE=1、TELEMETRY_OPTOUT=1。
历史限制：初始net8消费者因缺8.0.31 reference packs恢复失败，改兼容net10消费者；
初次run参数误转发，改build+直接DLL后成功。未清缓存，不把失败历史冒充通过。

公开范围另由Astra High复核：private profile自闭环可作为11票受限交付，
不称通用GM/T0099；SM3库存为恢复故障检测非认证，未覆盖ZIP元数据/库存外
描述字段变化。核心能力flag与默认Converter依赖不变。

最终包装ReadmeFile改为docs/password-crypto.md后，Low再以0.1.0-issue11.review1
和新隔离packages-review1目录实际复跑四组并核对三个新包hash/依赖/readme，通过。
见independent-review1-functional.json和最新package-manifest.json。
