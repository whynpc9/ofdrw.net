// Generated only after complete native asset verification. Do not edit.
namespace Ofdrw.Net.Core.Processes
{
    internal static partial class NativeSupervisorAssets
    {
        private static Asset FindAsset(string rid)
        {
            switch (rid)
            {
                case "linux-x64": return new Asset("linux-x64", 28904L, "79b3117e47835dff8f137d6e7363a87e622c7a1b891e278d7a1d94417c858595");
                case "linux-arm64": return new Asset("linux-arm64", 77376L, "1de0494b48e430cd69a47a3eabf3b933d8bdbd87f97cb59289afd2b07ec48956");
                case "osx-x64": return new Asset("osx-x64", 36704L, "2a8e4b0e362d2c8e8150d482e236b1d3bc2fb57f4c118c37e9f6687493cba5ae");
                case "osx-arm64": return new Asset("osx-arm64", 69424L, "1fe91767b723a916a9c75331117440654bee04435f654a7d8b69bdba65ee981f");
                default: throw new System.PlatformNotSupportedException("Unsupported native supervisor RID: " + rid);
            }
        }
    }
}
