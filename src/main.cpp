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
    CoInitializeGuard comGuard(COINIT_APARTMENTTHREADED);
    if (!comGuard) {
        MessageBoxW(nullptr, L"COM initialisatie mislukt.", L"Error", MB_ICONERROR);
        return 1;
    }

    INITCOMMONCONTROLSEX icc = {};
    icc.dwSize = sizeof(icc);
    icc.dwICC = ICC_STANDARD_CLASSES;
    InitCommonControlsEx(&icc);

    HWND hWnd = CreateMainWindow(hInstance);
    if (!hWnd) {
        MessageBoxW(nullptr, L"Venster aanmaken mislukt.", L"Error", MB_ICONERROR);
        return 1;
    }

    ShowWindow(hWnd, nCmdShow);
    UpdateWindow(hWnd);

    MSG msg;
    while (GetMessageW(&msg, nullptr, 0, 0)) {
        if (!IsDialogMessageW(hWnd, &msg)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    return static_cast<int>(msg.wParam);
}
