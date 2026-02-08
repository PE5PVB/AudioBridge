#include <windows.h>
#include <commctrl.h>
#include "ComHelper.h"
#include "resource.h"
#include "DialogProc.h"

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "msimg32.lib")
#pragma comment(lib, "uxtheme.lib")

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, LPWSTR, int nCmdShow) {
    // Single instance check
    HANDLE hMutex = CreateMutexW(nullptr, TRUE, L"AudioBridge_SingleInstance_PE5PVB");
    if (GetLastError() == ERROR_ALREADY_EXISTS) {
        // Find and activate existing window
        HWND existing = FindWindowW(L"AudioBridge", L"AudioBridge");
        if (existing) {
            ShowWindow(existing, SW_SHOW);
            ShowWindow(existing, SW_RESTORE);
            SetForegroundWindow(existing);
        }
        if (hMutex) CloseHandle(hMutex);
        return 0;
    }

    CoInitializeGuard comGuard(COINIT_APARTMENTTHREADED);
    if (!comGuard) {
        MessageBoxW(nullptr, L"COM initialisatie mislukt.", L"Error", MB_ICONERROR);
        if (hMutex) CloseHandle(hMutex);
        return 1;
    }

    INITCOMMONCONTROLSEX icc = {};
    icc.dwSize = sizeof(icc);
    icc.dwICC = ICC_STANDARD_CLASSES;
    InitCommonControlsEx(&icc);

    HWND hWnd = CreateMainWindow(hInstance);
    if (!hWnd) {
        MessageBoxW(nullptr, L"Venster aanmaken mislukt.", L"Error", MB_ICONERROR);
        if (hMutex) CloseHandle(hMutex);
        return 1;
    }

    if (ShouldStartMinimized()) {
        // Started via Task Scheduler with --startup flag and minimize to tray enabled
        ShowWindow(hWnd, SW_HIDE);
    } else {
        ShowWindow(hWnd, nCmdShow);
        UpdateWindow(hWnd);
    }

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        if (!IsDialogMessageW(hWnd, &msg)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    if (hMutex) CloseHandle(hMutex);
    return static_cast<int>(msg.wParam);
}
