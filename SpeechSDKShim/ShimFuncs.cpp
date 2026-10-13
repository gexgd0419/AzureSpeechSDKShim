// Shim functions for unsupported APIs
// As the entire references to system DLLs such as Kernel32.dll will be changed to this DLL,
// this DLL has to provide ALL export functions of the original DLLs.
// APIs supported on all systems (e.g. CreateFileW) are just redirected to the original DLLs.
// Most of the others are handled by YY-Thunks, only a few are written down below.
// Inspired by the vista2xp project

#include "framework.h"
#include <appmodel.h>
#include <ncrypt.h>

class HModule
{
    HMODULE h;
public:
    explicit HModule(LPCWSTR lib) : h(LoadLibraryW(lib)) {}
    ~HModule() { FreeLibrary(h); }
    operator HMODULE() const { return h; }
};

#define DEFINE_PFN(hModule, Proc) static const auto pfn##Proc = (decltype(Proc)*)GetProcAddress(hModule, #Proc)

static const HModule hKernel32(L"kernel32");
DEFINE_PFN(hKernel32, GetPackageFamilyName);
DEFINE_PFN(hKernel32, LoadPackagedLibrary);


LONG WINAPI Shim_GetPackageFamilyName(
    _In_ HANDLE hProcess,
    _Inout_ UINT32* packageFamilyNameLength,
    _Out_writes_opt_(*packageFamilyNameLength) PWSTR packageFamilyName
)
{
    if (pfnGetPackageFamilyName)
        return pfnGetPackageFamilyName(hProcess, packageFamilyNameLength, packageFamilyName);
    return APPMODEL_ERROR_NO_PACKAGE;
}

_Ret_maybenull_
HMODULE WINAPI Shim_LoadPackagedLibrary(
    _In_       LPCWSTR lpwLibFileName,
    _Reserved_ DWORD Reserved
)
{
    if (pfnLoadPackagedLibrary)
        return pfnLoadPackagedLibrary(lpwLibFileName, Reserved);
    SetLastError(APPMODEL_ERROR_NO_PACKAGE);
    return nullptr;
}
