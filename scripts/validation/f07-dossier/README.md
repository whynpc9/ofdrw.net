# F07 隔离包消费者

此消费者只引用 `Ofdrw.Net.Packaging` NuGet，不引用源码项目。它生成一个无签名、无字体/附件的确定性结构样例，调用已有 R17 副本处理器移除 Permissions，然后将源、副本及真实审计文件编入 F07 档案。它验证证据身份和完整标准仍 Unknown；不证明页面视觉、媒体/字体或附件符合性。

从仓库根目录执行（本地候选版本不能复用已发布包身份）：

```sh
export DOTNET_CLI_HOME="$PWD/artifacts/f07/dotnet-home"
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1
export DOTNET_CLI_TELEMETRY_OPTOUT=1
export F07_PACKAGE_VERSION=0.2.0-f07-validation.3
mkdir -p artifacts/f07/packages artifacts/f07/isolated-consumer

dotnet pack src/Ofdrw.Net.Core/Ofdrw.Net.Core.csproj -c Release -o artifacts/f07/packages -p:PackageVersion="$F07_PACKAGE_VERSION" -p:Version="$F07_PACKAGE_VERSION" -m:1 /nodeReuse:false /p:UseSharedCompilation=false --disable-build-servers
dotnet pack src/Ofdrw.Net.Packaging/Ofdrw.Net.Packaging.csproj -c Release -o artifacts/f07/packages -p:PackageVersion="$F07_PACKAGE_VERSION" -p:Version="$F07_PACKAGE_VERSION" -m:1 /nodeReuse:false /p:UseSharedCompilation=false --disable-build-servers
cp scripts/validation/f07-dossier/Consumer.csproj scripts/validation/f07-dossier/Program.cs artifacts/f07/isolated-consumer/
dotnet restore artifacts/f07/isolated-consumer/Consumer.csproj -p:ArchivalPackageVersion="$F07_PACKAGE_VERSION" --source "$PWD/artifacts/f07/packages" --source https://api.nuget.org/v3/index.json --packages "$PWD/artifacts/f07/isolated-packages" -m:1 /nodeReuse:false /p:UseSharedCompilation=false --disable-build-servers
dotnet build artifacts/f07/isolated-consumer/Consumer.csproj --no-restore -p:ArchivalPackageVersion="$F07_PACKAGE_VERSION" -p:RestorePackagesPath="$PWD/artifacts/f07/isolated-packages" -m:1 /nodeReuse:false /p:UseSharedCompilation=false --disable-build-servers
dotnet artifacts/f07/isolated-consumer/bin/Debug/net10.0/Consumer.dll artifacts/f07/consumer-output
```

输出目录必须没有旧的 `copy.ofd` 和审计文件，避免混淆产物身份；重复执行应另选输出目录。变更候选时更换版本或使用新的隔离包目录，不清理全局 NuGet 缓存。SDK 10；产物支持 netstandard2.0/netstandard2.1。记录源码 SHA、包 SHA-256 和 `environment.json`；`dossier.json` 直接保留源/结果/证据摘要。所有日常运行产物留在忽略的 artifacts 下。
