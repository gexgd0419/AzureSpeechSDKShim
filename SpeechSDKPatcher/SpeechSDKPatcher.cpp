// SpeechSDKPatcher
// A program that patches the Azure Speech SDK DLL files to make them compatible with pre Windows 10 systems.
// Their import tables will be modified with Detours, so that references to unsupported APIs in system DLLs
// will be replaced by references to SpeechSDKShim.dll, which contains shims for those APIs.
// 32-bit Detours supports patching 32-bit DLLs, and 64-bit Detours supports patching 64-bit DLLs,
// so we need two separate SpeechSDKPatcher exes for different bitness.

#include "framework.h"
#include "SpeechSDKPatcher.h"
#include <CommDlg.h>
#pragma comment (lib, "comdlg32.lib")

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

enum class PatchResult
{
    Failed = 0,
    NotModified,
    Patched
};

static PatchResult PatchDll(LPCWSTR dllPath, bool revert = false)
{
    BOOL ret = TRUE;

    // Open for read-only first
    HANDLE hFile = CreateFileW(dllPath, GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, 0, nullptr);
    if (hFile == INVALID_HANDLE_VALUE)
        return PatchResult::Failed;

    PDETOUR_BINARY pBinary = DetourBinaryOpen(hFile);
    if (!pBinary)
    {
        CloseHandle(hFile);
        return PatchResult::Failed;
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
        return PatchResult::Failed;
    }
    if (!modified)
    {
        DetourBinaryClose(pBinary);
        return PatchResult::NotModified;
    }

    // Open again for read & write, only if it should be modified
    // to avoid asking for write permission every time
    hFile = CreateFileW(dllPath, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ, nullptr, OPEN_EXISTING, 0, nullptr);
    if (hFile == INVALID_HANDLE_VALUE)
    {
        DetourBinaryClose(pBinary);
        return PatchResult::Failed;
    }

    ret = DetourBinaryWrite(pBinary, hFile);

    CloseHandle(hFile);
    DetourBinaryClose(pBinary);

    return ret ? PatchResult::Patched : PatchResult::Failed;
}

static bool FindFiles(LPCWSTR pathToFind, std::vector<std::wstring>& pathList)
{
    WIN32_FIND_DATAW fd;
    HANDLE hFind = FindFirstFileW(pathToFind, &fd);
    if (hFind == INVALID_HANDLE_VALUE)
        return false;
    do
    {
        pathList.emplace_back(fd.cFileName);
    } while (FindNextFileW(hFind, &fd));
    FindClose(hFind);
    return true;
}

static bool OpenFiles(std::vector<std::wstring>& pathList)
{
    auto pathBuf = std::make_unique<WCHAR[]>(4096);
    OPENFILENAMEW ofn = { sizeof ofn };
    ofn.lpstrTitle = L"Select Speech SDK DLL files to patch";
    ofn.lpstrFilter = L"Microsoft.CognitiveServices.Speech.*.dll\0Microsoft.CognitiveServices.Speech.*.dll\0";
    ofn.nFilterIndex = 1;
    ofn.lpstrFile = pathBuf.get();
    ofn.nMaxFile = 4096;
    ofn.Flags = OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY | OFN_ALLOWMULTISELECT;
    if (!GetOpenFileNameW(&ofn))
        return false;
    // Now the current directory has been changed
    // Just store the filenames (without directory path)
    for (LPCWSTR pFileName = pathBuf.get() + ofn.nFileOffset; *pFileName != L'\0'; pFileName += pathList.back().size() + 1)
    {
        pathList.emplace_back(pFileName);
    }
    return true;
}

static void ReportError(DWORD err = GetLastError())
{
    LPWSTR pStr = nullptr;
    FormatMessageW(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS | FORMAT_MESSAGE_ALLOCATE_BUFFER,
        nullptr, err, LANG_USER_DEFAULT, reinterpret_cast<LPWSTR>(&pStr), 0, nullptr);
    MessageBoxW(nullptr, pStr, L"SpeechSDKPatcher", MB_ICONEXCLAMATION);
    LocalFree(pStr);
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
    bool revert = false, quiet = false;

    for (int i = 1; i < argc; i++)
    {
        LPCWSTR arg = argv[i];
        if (_wcsicmp(arg, L"-revert") == 0)
            revert = true;
        else if (_wcsicmp(arg, L"-quiet") == 0)
            quiet = true;
        else if (arg[0] == L'-' || arg[0] == L'/')
        {
            MessageBoxW(nullptr,
                L"Command line usage: SpeechSDKPatcher [-revert] [-quiet] [<filename> <filename> ...]",
                L"SpeechSDKPatcher", MB_ICONINFORMATION);
            return 0;
        }
        else
            paths.emplace_back(arg);
    }

    if (paths.empty())
    {
        // no paths are provided through command line, add some default items
        // including all SpeechSDK DLLs and some UCRT-related DLLs

        WCHAR path[MAX_PATH];
        if (GetModuleFileNameW(nullptr, path, MAX_PATH) == MAX_PATH)
        {
            if (!quiet)
                ReportError(ERROR_FILENAME_EXCED_RANGE);
            return ERROR_FILENAME_EXCED_RANGE;
        }
        PathRemoveFileSpecW(path);
        if (!SetCurrentDirectoryW(path))
        {
            DWORD err = GetLastError();
            if (!quiet)
                ReportError(err);
            return err;
        }

        if (!FindFiles(L"Microsoft.CognitiveServices.Speech.*.dll", paths) && !quiet)
        {
            if (!OpenFiles(paths))
                return ERROR_CANCELLED;
            FindFiles(L"Microsoft.CognitiveServices.Speech.*.dll", paths);
        }

        // Other related files to patch
        static const wchar_t* pathsToFind[] = { L"msvcp140.dll", L"msvcp140_codecvt_ids.dll", L"vcruntime140.dll",
            L"vcruntime140_1.dll", L"ucrtbase.dll"};
        for (auto pathToFind : pathsToFind)
        {
            FindFiles(pathToFind, paths);
        }
    }

    UINT filecount = 0;
    for (const auto& path : paths)
    {
        PatchResult result = PatchDll(path.c_str(), revert);
        if (result == PatchResult::Failed)
        {
            DWORD err = GetLastError();
            if (!quiet)
            {
                if (err == ERROR_EXE_MARKED_INVALID)
                    MessageBoxW(nullptr,
                        L"The patcher you are using does not have the same bitness (32-bit or 64-bit) as the file(s) to be patched.",
                        L"SpeechSDKPatcher", MB_ICONEXCLAMATION);
                else
                    ReportError(err);
            }
            return err;
        }
        else if (result == PatchResult::Patched)
            filecount++;
    }

    if (!quiet)
    {
        auto msg = L"Patching completed.\r\n\r\n" + std::to_wstring(filecount) + L" file(s) changed.";
        if (!revert && !PathFileExistsW(L"SpeechSDKShim.dll"))
            msg += L"\r\n\r\nHowever, SpeechSDKShim.dll file does not exist in the same directory.\r\n"
                L"Put the shim DLL with the correct bitness in the directory to make the patched files work.";
        MessageBoxW(nullptr,
            msg.c_str(),
            L"SpeechSDKPatcher", MB_ICONINFORMATION);
    }

    return 0;
}
