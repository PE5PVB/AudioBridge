#include "DialogProc.h"
#include "resource.h"
#include "DeviceEnumerator.h"
#include "AudioRouter.h"
#include <commctrl.h>
#include <windowsx.h>
#include <dwmapi.h>
#include <uxtheme.h>
#include <vssym32.h>
#include <cstdio>
#include <vector>
#include <string>
#include <ks.h>
#include <ksmedia.h>
#include <gdiplus.h>
#include <shellapi.h>

#pragma comment(lib, "dwmapi.lib")
#pragma comment(lib, "uxtheme.lib")
#pragma comment(lib, "gdiplus.lib")

// ── FMDX.org color scheme ──────────────────────────────────────────
static constexpr COLORREF CLR_BG        = RGB(0x25, 0x27, 0x2e);
static constexpr COLORREF CLR_PANEL     = RGB(0x20, 0x22, 0x28);
static constexpr COLORREF CLR_ACCENT    = RGB(0x4d, 0xb6, 0x91);
static constexpr COLORREF CLR_ACCENT_DK = RGB(0x2b, 0x3a, 0x35);
static constexpr COLORREF CLR_TEXT      = RGB(0xff, 0xff, 0xff);
static constexpr COLORREF CLR_TEXT_DIM  = RGB(0x88, 0x88, 0x88);
static constexpr COLORREF CLR_RED       = RGB(0xff, 0x42, 0x42);
static constexpr COLORREF CLR_GREEN     = RGB(0x42, 0xff, 0x62);
static constexpr COLORREF CLR_COMBO_BG  = RGB(0x1a, 0x1c, 0x22);
static constexpr COLORREF CLR_BTN_HOVER = RGB(0x5c, 0xd4, 0xa8);

// ── Window metrics ─────────────────────────────────────────────────
static constexpr int WIN_W  = 500;
static constexpr int WIN_H  = 490;
static constexpr int MARGIN = 20;
static constexpr int COMBO_H = 28;
static constexpr int BTN_W  = 100;
static constexpr int BTN_H  = 34;
static constexpr int LABEL_H = 16;
static constexpr int RADIO_H = 20;
static constexpr int STATUS_H = 16;
static constexpr int GAP    = 6;
static constexpr int PANEL_PAD = 16;
static constexpr int PANEL_RADIUS = 12;

// ── State ──────────────────────────────────────────────────────────
static std::vector<AudioDeviceInfo> g_captureDevices;
static std::vector<AudioDeviceInfo> g_renderDevices;
static std::unique_ptr<AudioRouter> g_router;
static HFONT g_fontNormal = nullptr;
static HFONT g_fontBold   = nullptr;
static HFONT g_fontSmall  = nullptr;
static HFONT g_fontTitle  = nullptr;
static HBRUSH g_bgBrush    = nullptr;
static HBRUSH g_panelBrush = nullptr;
static HBRUSH g_comboBrush = nullptr;
static HBRUSH g_accentDkBrush = nullptr;
static HICON g_appIcon = nullptr;
static HANDLE g_fontHandles[3] = {};
static DWORD  g_fontHandleCount = 0;
static bool   g_titilliumLoaded = false;

static ULONG_PTR g_gdipToken = 0;
static HBITMAP   g_bgBitmap  = nullptr;
static RECT      g_linkRc    = {};

static HWND g_hBtnApply = nullptr;
static HWND g_hBtnStop  = nullptr;
static HWND g_tooltip   = nullptr;
static bool g_hoverApply = false;
static bool g_hoverStop  = false;
static bool g_exclusiveMode = false; // radio state: false=shared, true=exclusive
static RECT g_radioSharedRc = {};
static RECT g_radioExclusiveRc = {};

#define WM_APP_AUTOSTART (WM_APP + 1)

// ── Forward declarations ───────────────────────────────────────────
static LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
static LRESULT CALLBACK BtnSubclassProc(HWND, UINT, WPARAM, LPARAM, UINT_PTR, DWORD_PTR);
static void populateDevices(HWND hWnd);
static void onApply(HWND hWnd);
static void onStop(HWND hWnd);
static std::wstring formatInfo(const wchar_t* prefix, const WAVEFORMATEXTENSIBLE& fmt, UINT32 bufFrames);

// ── Helper: dark mode title bar (Windows 10 1809+) ────────────────
static void enableDarkTitleBar(HWND hWnd) {
    BOOL darkMode = TRUE;
    DwmSetWindowAttribute(hWnd, 20, &darkMode, sizeof(darkMode));
}

// ── Helper: draw rounded filled rect ──────────────────────────────
static void fillRoundRect(HDC hdc, RECT rc, int radius, COLORREF color) {
    HBRUSH br = CreateSolidBrush(color);
    HPEN pen = CreatePen(PS_SOLID, 1, color);
    HBRUSH oldBr = (HBRUSH)SelectObject(hdc, br);
    HPEN oldPen = (HPEN)SelectObject(hdc, pen);
    RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, radius, radius);
    SelectObject(hdc, oldBr);
    SelectObject(hdc, oldPen);
    DeleteObject(br);
    DeleteObject(pen);
}

// ── Helper: draw text with color ──────────────────────────────────
static void drawText(HDC hdc, const wchar_t* text, RECT rc, COLORREF color,
                     HFONT font, UINT fmt = DT_LEFT | DT_SINGLELINE | DT_NOPREFIX) {
    HFONT oldFont = (HFONT)SelectObject(hdc, font);
    SetTextColor(hdc, color);
    SetBkMode(hdc, TRANSPARENT);
    DrawTextW(hdc, text, -1, &rc, fmt);
    SelectObject(hdc, oldFont);
}

// ── Load embedded Titillium Web font from resources ───────────────
static void loadEmbeddedFonts(HINSTANCE hInstance) {
    int resIds[] = { IDR_FONT_REGULAR, IDR_FONT_LIGHT, IDR_FONT_SEMIBOLD };
    g_fontHandleCount = 0;
    for (int id : resIds) {
        HRSRC res = FindResourceW(hInstance, MAKEINTRESOURCEW(id), RT_RCDATA);
        if (!res) continue;
        HGLOBAL mem = LoadResource(hInstance, res);
        if (!mem) continue;
        void* data = LockResource(mem);
        DWORD size = SizeofResource(hInstance, res);
        if (!data || size == 0) continue;

        DWORD numFonts = 0;
        HANDLE h = AddFontMemResourceEx(data, size, nullptr, &numFonts);
        if (h && numFonts > 0) {
            g_fontHandles[g_fontHandleCount++] = h;
        }
    }
    g_titilliumLoaded = (g_fontHandleCount == 3);
}

static void unloadEmbeddedFonts() {
    for (DWORD i = 0; i < g_fontHandleCount; ++i) {
        if (g_fontHandles[i]) {
            RemoveFontMemResourceEx(g_fontHandles[i]);
            g_fontHandles[i] = nullptr;
        }
    }
    g_fontHandleCount = 0;
    g_titilliumLoaded = false;
}

// ── Settings persistence ──────────────────────────────────────────
struct AppSettings {
    std::wstring captureDeviceId;
    std::wstring renderDeviceId;
    bool exclusiveMode = false;
    bool autoStart = false;
};

static std::wstring getSettingsPath() {
    wchar_t appData[MAX_PATH] = {};
    GetEnvironmentVariableW(L"APPDATA", appData, MAX_PATH);
    std::wstring dir = std::wstring(appData) + L"\\AudioBridge";
    CreateDirectoryW(dir.c_str(), nullptr);
    return dir + L"\\settings.ini";
}

static AppSettings loadSettings() {
    AppSettings s;
    std::wstring path = getSettingsPath();
    wchar_t buf[512] = {};

    GetPrivateProfileStringW(L"Audio", L"CaptureDevice", L"", buf, 512, path.c_str());
    s.captureDeviceId = buf;

    GetPrivateProfileStringW(L"Audio", L"RenderDevice", L"", buf, 512, path.c_str());
    s.renderDeviceId = buf;

    s.exclusiveMode = GetPrivateProfileIntW(L"Audio", L"ExclusiveMode", 0, path.c_str()) != 0;
    s.autoStart = GetPrivateProfileIntW(L"Audio", L"AutoStart", 0, path.c_str()) != 0;

    return s;
}

static void saveSettings(const std::wstring& captureId, const std::wstring& renderId, bool exclusive) {
    std::wstring path = getSettingsPath();
    WritePrivateProfileStringW(L"Audio", L"CaptureDevice", captureId.c_str(), path.c_str());
    WritePrivateProfileStringW(L"Audio", L"RenderDevice", renderId.c_str(), path.c_str());
    WritePrivateProfileStringW(L"Audio", L"ExclusiveMode", exclusive ? L"1" : L"0", path.c_str());
}

static void saveAutoStart(bool running) {
    std::wstring path = getSettingsPath();
    WritePrivateProfileStringW(L"Audio", L"AutoStart", running ? L"1" : L"0", path.c_str());
}

// ── Load bg.jpg from embedded resource via GDI+ ──────────────────
static HBITMAP loadBgImage(HINSTANCE hInstance) {
    HRSRC hRes = FindResourceW(hInstance, MAKEINTRESOURCEW(IDR_BG_IMAGE), RT_RCDATA);
    if (!hRes) return nullptr;
    HGLOBAL hMem = LoadResource(hInstance, hRes);
    if (!hMem) return nullptr;
    void* data = LockResource(hMem);
    DWORD size = SizeofResource(hInstance, hRes);
    if (!data || size == 0) return nullptr;

    HGLOBAL hGlobal = GlobalAlloc(GMEM_MOVEABLE, size);
    if (!hGlobal) return nullptr;
    void* pGlobal = GlobalLock(hGlobal);
    memcpy(pGlobal, data, size);
    GlobalUnlock(hGlobal);

    IStream* stream = nullptr;
    if (FAILED(CreateStreamOnHGlobal(hGlobal, TRUE, &stream))) {
        GlobalFree(hGlobal);
        return nullptr;
    }

    Gdiplus::Bitmap bmp(stream);
    stream->Release();

    HBITMAP hBitmap = nullptr;
    bmp.GetHBITMAP(Gdiplus::Color(0, 0, 0), &hBitmap);
    return hBitmap;
}

// ── Create main window ────────────────────────────────────────────
HWND CreateMainWindow(HINSTANCE hInstance) {
    // Initialize GDI+
    Gdiplus::GdiplusStartupInput gdipInput;
    Gdiplus::GdiplusStartup(&g_gdipToken, &gdipInput, nullptr);

    loadEmbeddedFonts(hInstance);
    g_bgBitmap = loadBgImage(hInstance);
    const wchar_t* fontFace = g_titilliumLoaded ? L"Titillium Web" : L"Segoe UI";

    LOGFONTW lf = {};
    lf.lfHeight = -16;
    lf.lfWeight = FW_NORMAL;
    lf.lfCharSet = DEFAULT_CHARSET;
    lf.lfQuality = CLEARTYPE_QUALITY;
    wcscpy_s(lf.lfFaceName, fontFace);
    g_fontNormal = CreateFontIndirectW(&lf);

    lf.lfWeight = FW_SEMIBOLD;
    g_fontBold = CreateFontIndirectW(&lf);

    lf.lfHeight = -14;
    lf.lfWeight = FW_NORMAL;
    g_fontSmall = CreateFontIndirectW(&lf);

    lf.lfHeight = -24;
    lf.lfWeight = FW_LIGHT;
    g_fontTitle = CreateFontIndirectW(&lf);

    g_bgBrush      = CreateSolidBrush(CLR_BG);
    g_panelBrush   = CreateSolidBrush(CLR_PANEL);
    g_comboBrush   = CreateSolidBrush(CLR_COMBO_BG);
    g_accentDkBrush = CreateSolidBrush(CLR_ACCENT_DK);

    WNDCLASSEXW wc = {};
    wc.cbSize        = sizeof(wc);
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = hInstance;
    wc.hIcon         = LoadIconW(hInstance, MAKEINTRESOURCEW(IDI_APP_ICON));
    wc.hIconSm       = wc.hIcon;
    wc.hCursor       = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = g_bgBrush;
    wc.lpszClassName = L"AudioBridge";
    g_appIcon = wc.hIcon;
    RegisterClassExW(&wc);

    int screenW = GetSystemMetrics(SM_CXSCREEN);
    int screenH = GetSystemMetrics(SM_CYSCREEN);
    RECT wr = {0, 0, WIN_W, WIN_H};
    AdjustWindowRectEx(&wr, WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX, FALSE, 0);
    int ww = wr.right - wr.left;
    int wh = wr.bottom - wr.top;

    HWND hWnd = CreateWindowExW(
        0, L"AudioBridge", L"AudioBridge",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_MINIMIZEBOX,
        (screenW - ww) / 2, (screenH - wh) / 2, ww, wh,
        nullptr, nullptr, hInstance, nullptr);
    return hWnd;
}

// ── Create child controls ─────────────────────────────────────────
static void createControls(HWND hWnd) {
    HINSTANCE hInst = (HINSTANCE)GetWindowLongPtrW(hWnd, GWLP_HINSTANCE);
    int x = MARGIN + PANEL_PAD;
    int y = MARGIN + PANEL_PAD + 30;
    int innerW = WIN_W - 2 * MARGIN - 2 * PANEL_PAD;

    // Capture combo - owner-draw for dark styling
    y += LABEL_H + GAP;
    CreateWindowExW(0, L"COMBOBOX", nullptr,
        WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS | WS_VSCROLL | WS_TABSTOP,
        x, y, innerW, 200,
        hWnd, (HMENU)(INT_PTR)IDC_COMBO_CAPTURE, hInst, nullptr);
    y += COMBO_H + GAP + GAP;

    // Render combo - owner-draw for dark styling
    y += LABEL_H + GAP;
    CreateWindowExW(0, L"COMBOBOX", nullptr,
        WS_CHILD | WS_VISIBLE | CBS_DROPDOWNLIST | CBS_OWNERDRAWFIXED | CBS_HASSTRINGS | WS_VSCROLL | WS_TABSTOP,
        x, y, innerW, 200,
        hWnd, (HMENU)(INT_PTR)IDC_COMBO_RENDER, hInst, nullptr);
    y += COMBO_H + GAP + GAP;

    // WASAPI mode - custom drawn radio buttons (positions stored for hit testing)
    y += LABEL_H + GAP;
    g_radioSharedRc = {x, y, x + 130, y + RADIO_H};
    g_radioExclusiveRc = {x + 140, y, x + 280, y + RADIO_H};
    g_exclusiveMode = false;
    y += RADIO_H + GAP + GAP;

    // Buttons
    int btnAreaW = BTN_W * 2 + 12;
    int btnX = x + (innerW - btnAreaW) / 2;

    g_hBtnApply = CreateWindowExW(0, L"BUTTON", L"Apply",
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW | WS_TABSTOP,
        btnX, y, BTN_W, BTN_H,
        hWnd, (HMENU)(INT_PTR)IDC_BTN_APPLY, hInst, nullptr);

    g_hBtnStop = CreateWindowExW(0, L"BUTTON", L"Stop",
        WS_CHILD | WS_VISIBLE | BS_OWNERDRAW | WS_TABSTOP,
        btnX + BTN_W + 12, y, BTN_W, BTN_H,
        hWnd, (HMENU)(INT_PTR)IDC_BTN_STOP, hInst, nullptr);

    SetWindowSubclass(g_hBtnApply, BtnSubclassProc, IDC_BTN_APPLY, 0);
    SetWindowSubclass(g_hBtnStop, BtnSubclassProc, IDC_BTN_STOP, 0);

    HWND controls[] = {
        GetDlgItem(hWnd, IDC_COMBO_CAPTURE),
        GetDlgItem(hWnd, IDC_COMBO_RENDER),
        g_hBtnApply, g_hBtnStop
    };
    for (HWND h : controls) {
        if (h) SendMessageW(h, WM_SETFONT, (WPARAM)g_fontNormal, TRUE);
    }

    populateDevices(hWnd);

    // Create tooltips for radio buttons
    g_tooltip = CreateWindowExW(0, TOOLTIPS_CLASSW, nullptr,
        WS_POPUP | TTS_ALWAYSTIP,
        CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
        hWnd, nullptr, hInst, nullptr);

    if (g_tooltip) {
        SetWindowPos(g_tooltip, HWND_TOPMOST, 0, 0, 0, 0,
                     SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        SendMessageW(g_tooltip, TTM_SETMAXTIPWIDTH, 0, 350);

        TTTOOLINFOW ti = {};
        ti.cbSize = sizeof(ti);
        ti.uFlags = TTF_SUBCLASS;
        ti.hwnd = hWnd;

        ti.uId = 1;
        ti.rect = g_radioSharedRc;
        ti.lpszText = (LPWSTR)L"Uses the Windows audio engine.\r\nOther apps can play audio simultaneously.\r\nHigher latency but very compatible.";
        SendMessageW(g_tooltip, TTM_ADDTOOLW, 0, (LPARAM)&ti);

        ti.uId = 2;
        ti.rect = g_radioExclusiveRc;
        ti.lpszText = (LPWSTR)L"Bypasses the Windows audio engine for\r\ndirect hardware access. Lowest latency\r\nbut locks the device exclusively.";
        SendMessageW(g_tooltip, TTM_ADDTOOLW, 0, (LPARAM)&ti);
    }

    // Auto-start if was running when last closed
    AppSettings settings = loadSettings();
    if (settings.autoStart) {
        PostMessageW(hWnd, WM_APP_AUTOSTART, 0, 0);
    }
}

// ── Owner-draw combobox item painting ─────────────────────────────
static void drawComboItem(const DRAWITEMSTRUCT* dis) {
    if (dis->itemID == (UINT)-1) return;

    HDC hdc = dis->hDC;
    RECT rc = dis->rcItem;
    bool selected = (dis->itemState & ODS_SELECTED);
    bool inList = !(dis->itemState & ODS_COMBOBOXEDIT);

    // Background
    COLORREF bg = selected ? CLR_ACCENT_DK : (inList ? CLR_COMBO_BG : CLR_COMBO_BG);
    HBRUSH bgBr = CreateSolidBrush(bg);
    FillRect(hdc, &rc, bgBr);
    DeleteObject(bgBr);

    // Get text
    wchar_t text[256] = {};
    SendMessageW(dis->hwndItem, CB_GETLBTEXT, dis->itemID, (LPARAM)text);

    // Draw text
    rc.left += 6;
    HFONT oldFont = (HFONT)SelectObject(hdc, g_fontNormal);
    SetTextColor(hdc, selected ? CLR_ACCENT : CLR_TEXT);
    SetBkMode(hdc, TRANSPARENT);
    DrawTextW(hdc, text, -1, &rc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
    SelectObject(hdc, oldFont);
}

// ── Owner-draw button painting ────────────────────────────────────
static void drawButton(const DRAWITEMSTRUCT* dis) {
    HDC hdc = dis->hDC;
    RECT rc = dis->rcItem;
    bool isApply = (dis->CtlID == IDC_BTN_APPLY);
    bool hover = isApply ? g_hoverApply : g_hoverStop;
    bool pressed = (dis->itemState & ODS_SELECTED);

    // First: fill entire rect with panel color to kill white corners
    HBRUSH panelBr = CreateSolidBrush(CLR_PANEL);
    FillRect(hdc, &rc, panelBr);
    DeleteObject(panelBr);

    COLORREF bgColor, textColor;
    if (isApply) {
        bgColor = pressed ? CLR_ACCENT_DK : (hover ? CLR_BTN_HOVER : CLR_ACCENT);
        textColor = CLR_BG;
    } else {
        bgColor = pressed ? RGB(0x35, 0x37, 0x3e) : (hover ? RGB(0x3a, 0x3c, 0x44) : RGB(0x30, 0x32, 0x3a));
        textColor = CLR_TEXT;
    }

    fillRoundRect(hdc, rc, 10, bgColor);

    wchar_t text[64];
    GetWindowTextW(dis->hwndItem, text, 64);

    HFONT oldFont = (HFONT)SelectObject(hdc, g_fontBold);
    SetTextColor(hdc, textColor);
    SetBkMode(hdc, TRANSPARENT);
    DrawTextW(hdc, text, -1, &rc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);
    SelectObject(hdc, oldFont);
}

// ── Button subclass for hover ─────────────────────────────────────
static LRESULT CALLBACK BtnSubclassProc(HWND hWnd, UINT msg, WPARAM wParam,
                                         LPARAM lParam, UINT_PTR id, DWORD_PTR) {
    bool* hoverFlag = (id == IDC_BTN_APPLY) ? &g_hoverApply : &g_hoverStop;
    switch (msg) {
        case WM_MOUSEMOVE:
            if (!*hoverFlag) {
                *hoverFlag = true;
                TRACKMOUSEEVENT tme = {sizeof(tme), TME_LEAVE, hWnd, 0};
                TrackMouseEvent(&tme);
                InvalidateRect(hWnd, nullptr, FALSE);
            }
            break;
        case WM_MOUSELEAVE:
            *hoverFlag = false;
            InvalidateRect(hWnd, nullptr, FALSE);
            break;
        case WM_NCDESTROY:
            RemoveWindowSubclass(hWnd, BtnSubclassProc, id);
            break;
    }
    return DefSubclassProc(hWnd, msg, wParam, lParam);
}

// ── Paint the main window ─────────────────────────────────────────
static void onPaint(HWND hWnd) {
    PAINTSTRUCT ps;
    HDC hdc = BeginPaint(hWnd, &ps);

    RECT clientRect;
    GetClientRect(hWnd, &clientRect);

    HDC memDC = CreateCompatibleDC(hdc);
    HBITMAP memBmp = CreateCompatibleBitmap(hdc, clientRect.right, clientRect.bottom);
    HBITMAP oldBmp = (HBITMAP)SelectObject(memDC, memBmp);

    // Fill background with image or solid color
    if (g_bgBitmap) {
        HDC bgDC = CreateCompatibleDC(memDC);
        HBITMAP oldBg = (HBITMAP)SelectObject(bgDC, g_bgBitmap);

        BITMAP bm;
        GetObject(g_bgBitmap, sizeof(bm), &bm);

        int dstW = clientRect.right, dstH = clientRect.bottom;
        float scaleX = (float)dstW / bm.bmWidth;
        float scaleY = (float)dstH / bm.bmHeight;
        float scale = (scaleX > scaleY) ? scaleX : scaleY;

        int cropW = (int)(dstW / scale);
        int cropH = (int)(dstH / scale);
        int cropX = (bm.bmWidth - cropW) / 2;
        int cropY = (bm.bmHeight - cropH) / 2;

        SetStretchBltMode(memDC, HALFTONE);
        SetBrushOrgEx(memDC, 0, 0, nullptr);
        StretchBlt(memDC, 0, 0, dstW, dstH, bgDC, cropX, cropY, cropW, cropH, SRCCOPY);

        SelectObject(bgDC, oldBg);
        DeleteDC(bgDC);
    } else {
        HBRUSH bgBr = CreateSolidBrush(CLR_BG);
        FillRect(memDC, &clientRect, bgBr);
        DeleteObject(bgBr);
    }

    // Main panel - semi-transparent so background image shows through
    RECT panelRc = {MARGIN, MARGIN, WIN_W - MARGIN, WIN_H - MARGIN};
    {
        int panelW = panelRc.right - panelRc.left;
        int panelH = panelRc.bottom - panelRc.top;

        // Clip to rounded rect shape
        HRGN panelRgn = CreateRoundRectRgn(panelRc.left, panelRc.top,
                                            panelRc.right + 1, panelRc.bottom + 1,
                                            PANEL_RADIUS, PANEL_RADIUS);
        SelectClipRgn(memDC, panelRgn);

        // AlphaBlend panel color over background
        HDC panelDC = CreateCompatibleDC(memDC);
        HBITMAP panelBmp = CreateCompatibleBitmap(memDC, panelW, panelH);
        HBITMAP oldPnlBmp = (HBITMAP)SelectObject(panelDC, panelBmp);

        RECT localRc = {0, 0, panelW, panelH};
        HBRUSH pnlBr = CreateSolidBrush(CLR_PANEL);
        FillRect(panelDC, &localRc, pnlBr);
        DeleteObject(pnlBr);

        BLENDFUNCTION bf = {};
        bf.BlendOp = AC_SRC_OVER;
        bf.SourceConstantAlpha = 190; // ~75% opaque, bg shows through
        AlphaBlend(memDC, panelRc.left, panelRc.top, panelW, panelH,
                   panelDC, 0, 0, panelW, panelH, bf);

        SelectObject(panelDC, oldPnlBmp);
        DeleteObject(panelBmp);
        DeleteDC(panelDC);

        SelectClipRgn(memDC, nullptr);
        DeleteObject(panelRgn);
    }

    // Panel border
    HPEN borderPen = CreatePen(PS_SOLID, 1, RGB(0x35, 0x3a, 0x40));
    HPEN oldPen = (HPEN)SelectObject(memDC, borderPen);
    HBRUSH oldBr = (HBRUSH)SelectObject(memDC, GetStockObject(NULL_BRUSH));
    RoundRect(memDC, panelRc.left, panelRc.top, panelRc.right, panelRc.bottom,
              PANEL_RADIUS, PANEL_RADIUS);
    SelectObject(memDC, oldBr);
    SelectObject(memDC, oldPen);
    DeleteObject(borderPen);

    int x = MARGIN + PANEL_PAD;
    int y = MARGIN + PANEL_PAD;
    int innerW = WIN_W - 2 * MARGIN - 2 * PANEL_PAD;

    // Title
    RECT titleRc = {x, y, x + innerW, y + 34};
    drawText(memDC, L"AudioBridge", titleRc, CLR_ACCENT, g_fontTitle);
    y += 30;

    // Labels (drawn between controls)
    RECT lblCapture = {x, y, x + innerW, y + LABEL_H + 4};
    drawText(memDC, L"Input Device (Capture):", lblCapture, CLR_TEXT_DIM, g_fontSmall);
    y += LABEL_H + GAP + COMBO_H + GAP + GAP;

    RECT lblRender = {x, y, x + innerW, y + LABEL_H + 4};
    drawText(memDC, L"Output Device (Render):", lblRender, CLR_TEXT_DIM, g_fontSmall);
    y += LABEL_H + GAP + COMBO_H + GAP + GAP;

    RECT lblMode = {x, y, x + innerW, y + LABEL_H + 4};
    drawText(memDC, L"WASAPI Mode:", lblMode, CLR_TEXT_DIM, g_fontSmall);
    y += LABEL_H + GAP;

    // Custom radio buttons
    auto drawRadio = [&](RECT rc, const wchar_t* label, bool selected) {
        int circleSize = 14;
        int cx = rc.left;
        int cy = rc.top + (RADIO_H - circleSize) / 2;

        // Outer circle
        HPEN outerPen = CreatePen(PS_SOLID, 1, selected ? CLR_ACCENT : CLR_TEXT_DIM);
        HBRUSH innerBr = CreateSolidBrush(CLR_PANEL);
        HPEN oldP = (HPEN)SelectObject(memDC, outerPen);
        HBRUSH oldB = (HBRUSH)SelectObject(memDC, innerBr);
        Ellipse(memDC, cx, cy, cx + circleSize, cy + circleSize);
        SelectObject(memDC, oldP);
        SelectObject(memDC, oldB);
        DeleteObject(outerPen);
        DeleteObject(innerBr);

        // Inner filled circle when selected
        if (selected) {
            int pad = 3;
            HBRUSH accentBr = CreateSolidBrush(CLR_ACCENT);
            HPEN accentPen = CreatePen(PS_SOLID, 1, CLR_ACCENT);
            oldP = (HPEN)SelectObject(memDC, accentPen);
            oldB = (HBRUSH)SelectObject(memDC, accentBr);
            Ellipse(memDC, cx + pad, cy + pad, cx + circleSize - pad, cy + circleSize - pad);
            SelectObject(memDC, oldP);
            SelectObject(memDC, oldB);
            DeleteObject(accentBr);
            DeleteObject(accentPen);
        }

        // Label text
        RECT textRc = {cx + circleSize + 6, rc.top, rc.right, rc.bottom};
        drawText(memDC, label, textRc, selected ? CLR_TEXT : CLR_TEXT_DIM, g_fontNormal,
                 DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_NOPREFIX);
    };

    drawRadio(g_radioSharedRc, L"Shared Mode", !g_exclusiveMode);
    drawRadio(g_radioExclusiveRc, L"Exclusive Mode", g_exclusiveMode);
    y += RADIO_H + GAP + GAP + BTN_H + GAP + GAP;

    // ── Status panel ──────────────────────────────────────────────
    int statusY = y;
    int statusPanelH = 4 * STATUS_H + 3 * GAP + 16;
    RECT statusPanel = {MARGIN + PANEL_PAD - 4, statusY,
                        WIN_W - MARGIN - PANEL_PAD + 4, statusY + statusPanelH};
    fillRoundRect(memDC, statusPanel, 8, CLR_ACCENT_DK);

    int sx = MARGIN + PANEL_PAD + 8;
    int sw = WIN_W - 2 * MARGIN - 2 * PANEL_PAD - 16;
    int sy = statusY + 8;

    wchar_t statusBuf[256] = L"Status: Stopped";
    wchar_t capBuf[256]    = L"Capture: -";
    wchar_t renBuf[256]    = L"Render:  -";
    wchar_t latBuf[256]    = L"Latency: -  |  Underruns: 0";
    COLORREF statusClr = CLR_TEXT_DIM;

    if (g_router) {
        RouterStatus rs = g_router->getStatus();
        const wchar_t* stateStr = L"Stopped";
        switch (rs.state) {
            case RouterState::Running: stateStr = L"Running"; statusClr = CLR_GREEN; break;
            case RouterState::Error:   stateStr = L"Error";   statusClr = CLR_RED;   break;
            default: break;
        }
        if (rs.state == RouterState::Error && !rs.errorMessage.empty())
            swprintf_s(statusBuf, L"Status: %s - %s", stateStr, rs.errorMessage.c_str());
        else
            swprintf_s(statusBuf, L"Status: %s", stateStr);

        if (rs.state == RouterState::Running) {
            std::wstring capStr = formatInfo(L"Capture:", rs.captureFormat, rs.captureBufferFrames);
            wcscpy_s(capBuf, capStr.c_str());
            std::wstring renStr = formatInfo(L"Render: ", rs.renderFormat, rs.renderBufferFrames);
            wcscpy_s(renBuf, renStr.c_str());

            double capLatMs = 0, renLatMs = 0;
            if (rs.captureFormat.Format.nSamplesPerSec > 0)
                capLatMs = 1000.0 * rs.captureBufferFrames / rs.captureFormat.Format.nSamplesPerSec;
            if (rs.renderFormat.Format.nSamplesPerSec > 0)
                renLatMs = 1000.0 * rs.renderBufferFrames / rs.renderFormat.Format.nSamplesPerSec;
            swprintf_s(latBuf, L"Latency: ~%.1f ms  |  Underruns: %llu%s",
                       capLatMs + renLatMs, rs.underruns,
                       rs.resamplerActive ? L"  |  Resampler: active" : L"");
        }
    }

    RECT statusTextRc = {sx, sy, sx + sw, sy + STATUS_H + 4};
    drawText(memDC, statusBuf, statusTextRc, statusClr, g_fontSmall);

    sy += STATUS_H + GAP;
    RECT capRc = {sx, sy, sx + sw, sy + STATUS_H + 4};
    drawText(memDC, capBuf, capRc, CLR_TEXT, g_fontSmall);

    sy += STATUS_H + GAP;
    RECT renRc = {sx, sy, sx + sw, sy + STATUS_H + 4};
    drawText(memDC, renBuf, renRc, CLR_TEXT, g_fontSmall);

    sy += STATUS_H + GAP;
    RECT latRc = {sx, sy, sx + sw, sy + STATUS_H + 4};
    drawText(memDC, latBuf, latRc, CLR_ACCENT, g_fontSmall);

    // Footer - version and GitHub link
    int footerY = statusPanel.bottom + 14;
    RECT versionRc = {MARGIN + PANEL_PAD, footerY, WIN_W - MARGIN - PANEL_PAD, footerY + STATUS_H + 4};
    drawText(memDC, L"v1.00 - Sjef Verhoeven PE5PVB", versionRc, CLR_TEXT_DIM, g_fontSmall,
             DT_CENTER | DT_SINGLELINE | DT_NOPREFIX);

    g_linkRc = {MARGIN + PANEL_PAD, footerY + STATUS_H + 2,
                WIN_W - MARGIN - PANEL_PAD, footerY + 2 * STATUS_H + 6};
    drawText(memDC, L"github.com/PE5PVB/AudioBridge", g_linkRc, CLR_ACCENT, g_fontSmall,
             DT_CENTER | DT_SINGLELINE | DT_NOPREFIX);

    // Blit
    BitBlt(hdc, 0, 0, clientRect.right, clientRect.bottom, memDC, 0, 0, SRCCOPY);
    SelectObject(memDC, oldBmp);
    DeleteObject(memBmp);
    DeleteDC(memDC);

    EndPaint(hWnd, &ps);
}

// ── Window procedure ──────────────────────────────────────────────
static LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
        case WM_CREATE:
            enableDarkTitleBar(hWnd);
            createControls(hWnd);
            return 0;

        case WM_PAINT:
            onPaint(hWnd);
            return 0;

        case WM_ERASEBKGND:
            return 1;

        case WM_CTLCOLORLISTBOX: {
            HDC hdcLB = (HDC)wParam;
            SetTextColor(hdcLB, CLR_TEXT);
            SetBkColor(hdcLB, CLR_COMBO_BG);
            return (LRESULT)g_comboBrush;
        }

        case WM_CTLCOLOREDIT:
        case WM_CTLCOLORSTATIC: {
            HDC hdcCtl = (HDC)wParam;
            SetTextColor(hdcCtl, CLR_TEXT);
            SetBkColor(hdcCtl, CLR_PANEL);
            SetBkMode(hdcCtl, TRANSPARENT);
            return (LRESULT)g_panelBrush;
        }

        case WM_MEASUREITEM: {
            auto* mis = reinterpret_cast<MEASUREITEMSTRUCT*>(lParam);
            if (mis->CtlID == IDC_COMBO_CAPTURE || mis->CtlID == IDC_COMBO_RENDER) {
                mis->itemHeight = 24;
                return TRUE;
            }
            break;
        }

        case WM_DRAWITEM: {
            auto* dis = reinterpret_cast<DRAWITEMSTRUCT*>(lParam);
            if (dis->CtlID == IDC_BTN_APPLY || dis->CtlID == IDC_BTN_STOP) {
                drawButton(dis);
                return TRUE;
            }
            if (dis->CtlID == IDC_COMBO_CAPTURE || dis->CtlID == IDC_COMBO_RENDER) {
                drawComboItem(dis);
                return TRUE;
            }
            break;
        }

        case WM_SETCURSOR: {
            if (LOWORD(lParam) == HTCLIENT) {
                POINT pt;
                GetCursorPos(&pt);
                ScreenToClient(hWnd, &pt);
                if (PtInRect(&g_linkRc, pt)) {
                    SetCursor(LoadCursorW(nullptr, IDC_HAND));
                    return TRUE;
                }
            }
            break;
        }

        case WM_LBUTTONDOWN: {
            POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
            if (PtInRect(&g_linkRc, pt)) {
                ShellExecuteW(nullptr, L"open", L"https://github.com/PE5PVB/AudioBridge", nullptr, nullptr, SW_SHOWNORMAL);
            } else if (PtInRect(&g_radioSharedRc, pt) && g_exclusiveMode) {
                g_exclusiveMode = false;
                InvalidateRect(hWnd, nullptr, FALSE);
            } else if (PtInRect(&g_radioExclusiveRc, pt) && !g_exclusiveMode) {
                g_exclusiveMode = true;
                InvalidateRect(hWnd, nullptr, FALSE);
            }
            break;
        }

        case WM_COMMAND:
            switch (LOWORD(wParam)) {
                case IDC_BTN_APPLY:
                    onApply(hWnd);
                    return 0;
                case IDC_BTN_STOP:
                    onStop(hWnd);
                    return 0;
            }
            break;

        case WM_TIMER:
            if (wParam == IDT_STATUS_TIMER) {
                InvalidateRect(hWnd, nullptr, FALSE);
                if (g_router) {
                    RouterStatus rs = g_router->getStatus();
                    if (rs.state == RouterState::Error)
                        KillTimer(hWnd, IDT_STATUS_TIMER);
                }
                return 0;
            }
            break;

        case WM_APP_AUTOSTART:
            onApply(hWnd);
            return 0;

        case WM_CLOSE: {
            bool wasRunning = g_router && g_router->getStatus().state == RouterState::Running;
            onStop(hWnd);              // This sets AutoStart=0
            saveAutoStart(wasRunning); // Override: restore true if was running
            DestroyWindow(hWnd);
            return 0;
        }

        case WM_DESTROY:
            if (g_fontNormal) { DeleteObject(g_fontNormal); g_fontNormal = nullptr; }
            if (g_fontBold)   { DeleteObject(g_fontBold);   g_fontBold = nullptr; }
            if (g_fontSmall)  { DeleteObject(g_fontSmall);  g_fontSmall = nullptr; }
            if (g_fontTitle)  { DeleteObject(g_fontTitle);  g_fontTitle = nullptr; }
            if (g_panelBrush) { DeleteObject(g_panelBrush); g_panelBrush = nullptr; }
            if (g_comboBrush) { DeleteObject(g_comboBrush); g_comboBrush = nullptr; }
            if (g_accentDkBrush) { DeleteObject(g_accentDkBrush); g_accentDkBrush = nullptr; }
            if (g_tooltip) { DestroyWindow(g_tooltip); g_tooltip = nullptr; }
            if (g_bgBitmap) { DeleteObject(g_bgBitmap); g_bgBitmap = nullptr; }
            unloadEmbeddedFonts();
            if (g_gdipToken) { Gdiplus::GdiplusShutdown(g_gdipToken); g_gdipToken = 0; }
            PostQuitMessage(0);
            return 0;
    }
    return DefWindowProcW(hWnd, msg, wParam, lParam);
}

// ── Device enumeration ────────────────────────────────────────────
static void populateDevices(HWND hWnd) {
    HWND hCapture = GetDlgItem(hWnd, IDC_COMBO_CAPTURE);
    HWND hRender  = GetDlgItem(hWnd, IDC_COMBO_RENDER);

    SendMessageW(hCapture, CB_RESETCONTENT, 0, 0);
    SendMessageW(hRender, CB_RESETCONTENT, 0, 0);

    g_captureDevices.clear();
    g_renderDevices.clear();

    DeviceEnumerator::enumerateCapture(g_captureDevices);
    DeviceEnumerator::enumerateRender(g_renderDevices);

    for (auto& dev : g_captureDevices)
        SendMessageW(hCapture, CB_ADDSTRING, 0, (LPARAM)dev.name.c_str());
    for (auto& dev : g_renderDevices)
        SendMessageW(hRender, CB_ADDSTRING, 0, (LPARAM)dev.name.c_str());

    if (!g_captureDevices.empty()) SendMessageW(hCapture, CB_SETCURSEL, 0, 0);
    if (!g_renderDevices.empty())  SendMessageW(hRender, CB_SETCURSEL, 0, 0);

    // Restore saved settings
    AppSettings settings = loadSettings();
    g_exclusiveMode = settings.exclusiveMode;

    for (int i = 0; i < (int)g_captureDevices.size(); ++i) {
        if (g_captureDevices[i].id == settings.captureDeviceId) {
            SendMessageW(hCapture, CB_SETCURSEL, i, 0);
            break;
        }
    }
    for (int i = 0; i < (int)g_renderDevices.size(); ++i) {
        if (g_renderDevices[i].id == settings.renderDeviceId) {
            SendMessageW(hRender, CB_SETCURSEL, i, 0);
            break;
        }
    }
}

// ── Format info string ────────────────────────────────────────────
static std::wstring formatInfo(const wchar_t* prefix, const WAVEFORMATEXTENSIBLE& fmt, UINT32 bufFrames) {
    const wchar_t* sampleType = L"PCM";
    if (fmt.Format.wFormatTag == WAVE_FORMAT_EXTENSIBLE) {
        if (IsEqualGUID(fmt.SubFormat, KSDATAFORMAT_SUBTYPE_IEEE_FLOAT))
            sampleType = L"float";
    } else if (fmt.Format.wFormatTag == WAVE_FORMAT_IEEE_FLOAT) {
        sampleType = L"float";
    }

    wchar_t buf[256];
    swprintf_s(buf, L"%s %u Hz, %u-bit %s, %uch, buf=%u",
               prefix,
               fmt.Format.nSamplesPerSec,
               (fmt.Format.wFormatTag == WAVE_FORMAT_EXTENSIBLE)
                   ? fmt.Samples.wValidBitsPerSample : fmt.Format.wBitsPerSample,
               sampleType,
               fmt.Format.nChannels,
               bufFrames);
    return buf;
}

// ── Apply / Stop handlers ─────────────────────────────────────────
static void onApply(HWND hWnd) {
    int capIdx = (int)SendDlgItemMessageW(hWnd, IDC_COMBO_CAPTURE, CB_GETCURSEL, 0, 0);
    int renIdx = (int)SendDlgItemMessageW(hWnd, IDC_COMBO_RENDER, CB_GETCURSEL, 0, 0);

    if (capIdx < 0 || capIdx >= (int)g_captureDevices.size() ||
        renIdx < 0 || renIdx >= (int)g_renderDevices.size()) {
        InvalidateRect(hWnd, nullptr, FALSE);
        return;
    }

    bool exclusive = g_exclusiveMode;

    // Save settings
    saveSettings(g_captureDevices[capIdx].id, g_renderDevices[renIdx].id, exclusive);

    if (g_router) g_router->stop();
    g_router = std::make_unique<AudioRouter>();

    g_router->start(g_captureDevices[capIdx].id,
                    g_renderDevices[renIdx].id, exclusive);

    SetTimer(hWnd, IDT_STATUS_TIMER, 500, nullptr);
    InvalidateRect(hWnd, nullptr, FALSE);
}

static void onStop(HWND hWnd) {
    KillTimer(hWnd, IDT_STATUS_TIMER);
    if (g_router) {
        g_router->stop();
        g_router.reset();
    }
    saveAutoStart(false);
    InvalidateRect(hWnd, nullptr, FALSE);
}
