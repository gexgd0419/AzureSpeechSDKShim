A compatibility shim layer that can make Azure Speech SDK, which normally requires Windows 10, able to work on Windows 7.

Maybe not everything in Azure Speech SDK can work well on Windows 7. But notably, local (embedded) natural voices work. So this can **make local Narrator voices on Windows 11 work on Windows 7**.

This is a sub-project of [NaturalVoiceSAPIAdapter](https://github.com/gexgd0419/NaturalVoiceSAPIAdapter), but this can also be used as a stand-alone program.

Extract SpeechSDKPatcher.exe and SpeechSDKShim.dll into the folder that has the Azure Speech SDK DLL files (Microsoft.CognitiveServices.Speech.*.dll), then run SpeechSDKPatcher.exe. It will automatically find and patch every Azure Speech SDK DLLs to use SpeechSDKShim.dll, which provides shim functions for unsupported system APIs, instead of the real system DLLs.

Azure Speech SDK also requires msvcp140.dll, msvcp140_codecvt_ids.dll, vcruntime_140.dll, vcruntime_140_1.dll, and the Universal CRT (UCRT). This patcher will make the SDK require only a single ucrtbase.dll instead of the whole UCRT, and if MSVC runtime DLLs are in the same folder as the SDK, those DLLs will be patched as well.

The patched SDK's local natural voice feature has been tested and proven to work on freshly installed Windows 7 SP0/SP1 on virtual machines. Although the patcher can work on Windows Vista and even XP, the patched SDK couldn't really work, and would just crash the client application for no apparent reason.