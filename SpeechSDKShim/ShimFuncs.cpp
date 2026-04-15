// Shim functions for unsupported APIs
// As the entire references to system DLLs such as Kernel32.dll will be changed to this DLL,
// this DLL has to provide ALL export functions of the original DLLs.
// APIs supported on all systems (e.g. CreateFileW) are just redirected to the original DLLs.
// Most of the others are handled by YY-Thunks, only a few are written down below.
// Inspired by the vista2xp project

#include "framework.h"
#include <appmodel.h>
#include <ncrypt.h>
#include <cstdlib>
#include <bit>

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
DEFINE_PFN(hKernel32, GetLogicalProcessorInformationEx);

static const HModule hNcrypt(L"ncrypt");
DEFINE_PFN(hNcrypt, NCryptOpenStorageProvider);
DEFINE_PFN(hNcrypt, NCryptImportKey);
DEFINE_PFN(hNcrypt, NCryptFreeObject);

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

SECURITY_STATUS WINAPI Shim_NCryptOpenStorageProvider(
    _Out_   NCRYPT_PROV_HANDLE* phProvider,
    _In_opt_ LPCWSTR pszProviderName,
    _In_    DWORD   dwFlags)
{
    if (pfnNCryptOpenStorageProvider)
        return pfnNCryptOpenStorageProvider(phProvider, pszProviderName, dwFlags);
    return ERROR_NOT_SUPPORTED;
}

SECURITY_STATUS WINAPI Shim_NCryptImportKey(
    _In_    NCRYPT_PROV_HANDLE hProvider,
    _In_opt_ NCRYPT_KEY_HANDLE hImportKey,
    _In_    LPCWSTR pszBlobType,
    _In_opt_ NCryptBufferDesc* pParameterList,
    _Out_   NCRYPT_KEY_HANDLE* phKey,
    _In_reads_bytes_(cbData) PBYTE pbData,
    _In_    DWORD   cbData,
    _In_    DWORD   dwFlags)
{
    if (pfnNCryptImportKey)
        return pfnNCryptImportKey(hProvider, hImportKey, pszBlobType, pParameterList, phKey, pbData, cbData, dwFlags);
    return ERROR_NOT_SUPPORTED;
}

SECURITY_STATUS WINAPI Shim_NCryptFreeObject(
    _In_    NCRYPT_HANDLE hObject)
{
    if (pfnNCryptFreeObject)
        return pfnNCryptFreeObject(hObject);
    return ERROR_NOT_SUPPORTED;
}

BOOL
WINAPI
Shim_GetLogicalProcessorInformationEx(
    _In_ LOGICAL_PROCESSOR_RELATIONSHIP RelationshipType,
    _Out_writes_bytes_to_opt_(*ReturnedLength, *ReturnedLength) PSYSTEM_LOGICAL_PROCESSOR_INFORMATION_EX Buffer,
    _Inout_ PDWORD ReturnedLength
)
{
    // There's a bug in YY-Thunks 1.1.9 that can corrupt the memory heap and crash the process
    // when GetLogicalProcessorInformationEx is called.
    // https://github.com/Chuyu-Team/YY-Thunks/issues/175
    // So we use our own modified implementation here.

    if (pfnGetLogicalProcessorInformationEx)
        return pfnGetLogicalProcessorInformationEx(RelationshipType, Buffer, ReturnedLength);

    // GetLogicalProcessorInformationEx calls the following under the hood:
    // NtQuerySystemInformationEx(107, &RelationshipType, sizeof(RelationshipType), Buffer, *ReturnedLength, ReturnedLength);
    // Before calling NtQuerySystemInformationEx, it checks whether the ReturnedLength pointer is NULL:

    if (!ReturnedLength)
    {
        SetLastError(ERROR_INVALID_PARAMETER);
        return FALSE;
    }

    DWORD cbBuffer = *ReturnedLength;

    // Since *ReturnedLength is passed to NtQuerySystemInformationEx as the output buffer length,
    // when *ReturnedLength != 0, NtQuerySystemInformationEx will always probe the output buffer for write,
    // even when Buffer == NULL.
    // So when *ReturnedLength != 0 and Buffer == NULL (or points to an inaccessible memory area),
    // this function will fail with ERROR_NOACCESS, and *ReturnedLength will not be updated.
    // Callers therefore should always initialize *ReturnedLength to zero before calling
    // if they want to get the requested buffer size.
    // Here we try to replicate the same behavior with IsBadWritePtr.

    if (cbBuffer != 0 && (!Buffer || IsBadWritePtr(Buffer, cbBuffer)))
    {
        SetLastError(ERROR_NOACCESS);
        return FALSE;
    }

    // Build the information we need with GetLogicalProcessorInformation.

    DWORD cbInfoBuf = 0;
    GetLogicalProcessorInformation(nullptr, &cbInfoBuf);
    if (GetLastError() != ERROR_INSUFFICIENT_BUFFER)
        return FALSE;

    auto pInfoBuf = static_cast<PSYSTEM_LOGICAL_PROCESSOR_INFORMATION>(malloc(cbInfoBuf));
    if (!pInfoBuf)
    {
        SetLastError(ERROR_NOT_ENOUGH_MEMORY);
        return FALSE;
    }

    if (!GetLogicalProcessorInformation(pInfoBuf, &cbInfoBuf))
    {
        free(pInfoBuf);
        return FALSE;
    }

    auto pInfoEnd = reinterpret_cast<PSYSTEM_LOGICAL_PROCESSOR_INFORMATION>(reinterpret_cast<PBYTE>(pInfoBuf) + cbInfoBuf);

    DWORD cbRequired = 0;
    PBYTE pOutBuffer = reinterpret_cast<PBYTE>(Buffer);

    for (auto pInfo = pInfoBuf; pInfo < pInfoEnd; pInfo++)
    {
        if (RelationshipType != RelationAll && pInfo->Relationship != RelationshipType)
            continue;

        DWORD cbItem;
        // These are the only relationship types supported by GetLogicalProcessorInformation
        switch (pInfo->Relationship)
        {
        case RelationProcessorCore:
        case RelationProcessorPackage:
            cbItem = RTL_SIZEOF_THROUGH_FIELD(SYSTEM_LOGICAL_PROCESSOR_INFORMATION_EX, Processor);
            break;
        case RelationNumaNode:
            cbItem = RTL_SIZEOF_THROUGH_FIELD(SYSTEM_LOGICAL_PROCESSOR_INFORMATION_EX, NumaNode);
            break;
        case RelationCache:
            cbItem = RTL_SIZEOF_THROUGH_FIELD(SYSTEM_LOGICAL_PROCESSOR_INFORMATION_EX, Cache);
            break;
        default:
            continue;
        }

        cbRequired += cbItem;

        if (cbRequired > cbBuffer)
            continue;

        memset(pOutBuffer, 0, cbItem);
        
        auto pOut = reinterpret_cast<PSYSTEM_LOGICAL_PROCESSOR_INFORMATION_EX>(pOutBuffer);

        pOut->Size = cbItem;
        pOut->Relationship = pInfo->Relationship;

        switch (pInfo->Relationship)
        {
        case RelationProcessorCore:
        case RelationProcessorPackage:
            pOut->Processor.Flags = pInfo->ProcessorCore.Flags;
            pOut->Processor.GroupCount = 1;
            pOut->Processor.GroupMask->Mask = pInfo->ProcessorMask;
            break;
        case RelationNumaNode:
            pOut->NumaNode.NodeNumber = pInfo->NumaNode.NodeNumber;
            pOut->NumaNode.GroupMask.Mask = pInfo->ProcessorMask;
            break;
        case RelationCache:
            pOut->Cache.Level = pInfo->Cache.Level;
            pOut->Cache.Associativity = pInfo->Cache.Associativity;
            pOut->Cache.LineSize = pInfo->Cache.LineSize;
            pOut->Cache.CacheSize = pInfo->Cache.Size;
            pOut->Cache.Type = pInfo->Cache.Type;
            pOut->Cache.GroupMask.Mask = pInfo->ProcessorMask;
            break;
        }

        pOutBuffer += cbItem;
    }

    free(pInfoBuf);

    if (RelationshipType == RelationAll || RelationshipType == RelationGroup)
    {
        DWORD cbItem = RTL_SIZEOF_THROUGH_FIELD(SYSTEM_LOGICAL_PROCESSOR_INFORMATION_EX, Group);

        cbRequired += cbItem;

        if (cbRequired <= cbBuffer)
        {
            memset(pOutBuffer, 0, cbItem);
            auto pOut = reinterpret_cast<PSYSTEM_LOGICAL_PROCESSOR_INFORMATION_EX>(pOutBuffer);

            pOut->Size = cbItem;
            pOut->Relationship = RelationGroup;

            pOut->Group.ActiveGroupCount = 1;
            pOut->Group.MaximumGroupCount = 1;

            SYSTEM_INFO sysInfo;

            GetSystemInfo(&sysInfo);

            pOut->Group.GroupInfo->ActiveProcessorMask = sysInfo.dwActiveProcessorMask;
            pOut->Group.GroupInfo->ActiveProcessorCount = static_cast<BYTE>(std::popcount(sysInfo.dwActiveProcessorMask));
            pOut->Group.GroupInfo->MaximumProcessorCount = static_cast<BYTE>(sysInfo.dwNumberOfProcessors);
        }
    }

    if (cbRequired == 0)
    {
        // When no item matching the specified relation is found, return this error.
        SetLastError(ERROR_GEN_FAILURE);
        return FALSE;
    }

    *ReturnedLength = cbRequired;

    if (cbRequired > cbBuffer)
    {
        SetLastError(ERROR_INSUFFICIENT_BUFFER);
        return FALSE;
    }

    return TRUE;
}