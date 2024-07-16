// SpeechSDKPatcher
// A program that patches the Azure Speech SDK DLL files to make them compatible with pre Windows 10 systems.
// Their import tables will be modified with Detours, so that references to unsupported APIs in system DLLs
// will be replaced by references to SpeechSDKShim.dll, which contains shims for those APIs.
// 32-bit Detours supports patching 32-bit DLLs, and 64-bit Detours supports patching 64-bit DLLs,
// so we need two separate SpeechSDKPatcher exes for different bitness.

#include "framework.h"
#include "SpeechSDKPatcher.h"

static bool ShouldReplaceImportFile(LPCSTR pszOrigFile)
{
    // Replace references to "API set" DLLs that usually points to Kernel32 APIs
    if (_strnicmp(pszOrigFile, "api-ms-win-core-", 16) == 0)
        return true;

    static const char* dllnames[] = { "kernel32.dll", "advapi32.dll", "ncrypt.dll", "api-ms-win-eventing-provider-l1-1-0.dll" };
    for (auto dllname : dllnames)
    {
        if (_stricmp(pszOrigFile, dllname) == 0)
            return true;
    }

    return false;
}

static BOOL CALLBACK DetourFileCallback(PVOID pContext, LPCSTR pszOrigFile, LPCSTR pszFile, LPCSTR* ppszOutFile)
{
    if (ShouldReplaceImportFile(pszOrigFile) && _stricmp(pszFile, "SpeechSDKShim.dll") != 0)
    {
        *ppszOutFile = "SpeechSDKShim.dll";
        *(bool*)pContext = true;
    }
    else if (_strnicmp(pszOrigFile, "api-ms-win-crt-", 15) == 0 && _stricmp(pszFile, "ucrtbase.dll") != 0)
    {
        // Replace references to "API set" DLLs that usually points to ucrtbase.dll
        *ppszOutFile = "ucrtbase.dll";
        *(bool*)pContext = true;
    }

    return TRUE;
}

static BOOL CALLBACK DetourResetFileCallback(PVOID pContext, LPCSTR pszOrigFile, LPCSTR pszFile, LPCSTR* ppszOutFile)
{
    // Reset previously modified references
    if (pszOrigFile != pszFile && _stricmp(pszOrigFile, pszFile) != 0)
    {
        *ppszOutFile = pszOrigFile;
        *(bool*)pContext = true;
    }
    return TRUE;
}

static BOOL PatchDll(LPCWSTR dllPath, bool revert = false)
{
    BOOL ret = TRUE;

    // Open for read-only first
    HANDLE hFile = CreateFileW(dllPath, GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, 0, nullptr);
    if (hFile == INVALID_HANDLE_VALUE)
        return FALSE;

    PDETOUR_BINARY pBinary = DetourBinaryOpen(hFile);
    if (!pBinary)
    {
        CloseHandle(hFile);
        return FALSE;
    }

    // Check if the DLL is modified
    bool modified = false;
    if (revert)
        ret = DetourBinaryEditImports(pBinary, &modified, nullptr, DetourResetFileCallback, nullptr, nullptr);
    else
        ret = DetourBinaryEditImports(pBinary, &modified, nullptr, DetourFileCallback, nullptr, nullptr);

    CloseHandle(hFile);
    if (!ret)
    {
        DetourBinaryClose(pBinary);
        return FALSE;
    }
    if (!modified)
    {
        DetourBinaryClose(pBinary);
        return TRUE;
    }

    // Open again for read & write, only if it should be modified
    // to avoid asking for write permission every time
    hFile = CreateFileW(dllPath, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ, nullptr, OPEN_EXISTING, 0, nullptr);
    if (hFile == INVALID_HANDLE_VALUE)
    {
        DetourBinaryClose(pBinary);
        return FALSE;
    }

    ret = DetourBinaryWrite(pBinary, hFile);

    CloseHandle(hFile);
    DetourBinaryClose(pBinary);

    return ret;
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hInstance);
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);
    UNREFERENCED_PARAMETER(nCmdShow);

    // read the files to be patched from command line args
    int argc = __argc;
    wchar_t** argv = __wargv;
    std::vector<std::wstring> paths;
    bool revert = false;

    for (int i = 1; i < argc; i++)
    {
        LPCWSTR arg = argv[i];
        if (_wcsicmp(arg, L"-revert") == 0)
            revert = true;
        else
            paths.emplace_back(arg);
    }

    if (paths.empty())
    {
        // no paths are provided through command line, add some default items
        // including all SpeechSDK DLLs and some UCRT-related DLLs

        WCHAR path[MAX_PATH];
        if (GetModuleFileNameW(nullptr, path, MAX_PATH) == MAX_PATH)
            return ERROR_FILENAME_EXCED_RANGE;
        PathRemoveFileSpecW(path);
        if (!SetCurrentDirectoryW(path))
            return GetLastError();

        static const wchar_t* pathsToFind[] = { L"Microsoft.CognitiveServices.Speech.*.dll", L"msvcp140.dll",
            L"msvcp140_codecvt_ids.dll", L"vcruntime140.dll", L"vcruntime140_1.dll", L"ucrtbase.dll"};
        for (auto pathToFind : pathsToFind)
        {
            WIN32_FIND_DATAW fd;
            HANDLE hFind = FindFirstFileW(pathToFind, &fd);
            if (hFind != INVALID_HANDLE_VALUE)
            {
                do
                {
                    paths.emplace_back(fd.cFileName);
                } while (FindNextFileW(hFind, &fd));
                FindClose(hFind);
            }
        }
    }

    for (const auto& path : paths)
    {
        if (!PatchDll(path.c_str(), revert))
        {
            return GetLastError();
        }
    }

    return 0;
}
