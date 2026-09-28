#include <windows.h>
#include <tlhelp32.h>
#include <msi.h>
#include <msiquery.h>
#include <cwchar>

// Read-only checks run before the MSI transaction. Never close Excel or force a reboot.
extern "C" __declspec(dllexport) UINT __stdcall CheckEnvironment(MSIHANDLE session) {
    SYSTEM_INFO system{};
    GetNativeSystemInfo(&system);
    MsiSetPropertyW(session, L"CUPOLA_PLATFORM_SUPPORTED",
                    system.wProcessorArchitecture == PROCESSOR_ARCHITECTURE_AMD64 ? L"1" : L"");
    HANDLE snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
    if (snapshot == INVALID_HANDLE_VALUE) return ERROR_INSTALL_FAILURE;
    PROCESSENTRY32W entry{};
    entry.dwSize = sizeof(entry);
    bool running = false;
    if (!Process32FirstW(snapshot, &entry)) { CloseHandle(snapshot); return ERROR_INSTALL_FAILURE; }
    do {
        if (_wcsicmp(entry.szExeFile, L"EXCEL.EXE") == 0) { running = true; break; }
    } while (Process32NextW(snapshot, &entry));
    CloseHandle(snapshot);
    MsiSetPropertyW(session, L"CUPOLA_EXCEL_RUNNING", running ? L"1" : L"");
    return ERROR_SUCCESS;
}
