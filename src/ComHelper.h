#pragma once

#include <wrl/client.h>
#include <objbase.h>
#include <cstdio>

using Microsoft::WRL::ComPtr;

#define RETURN_IF_FAILED(hr)                                    \
    do {                                                        \
        HRESULT _hr = (hr);                                     \
        if (FAILED(_hr)) {                                      \
            return _hr;                                         \
        }                                                       \
    } while (false)

struct CoInitializeGuard {
    HRESULT hr;
    CoInitializeGuard(DWORD model = COINIT_MULTITHREADED) {
        hr = CoInitializeEx(nullptr, model);
    }
    ~CoInitializeGuard() {
        if (SUCCEEDED(hr))
            CoUninitialize();
    }
    explicit operator bool() const { return SUCCEEDED(hr); }
    CoInitializeGuard(const CoInitializeGuard&) = delete;
    CoInitializeGuard& operator=(const CoInitializeGuard&) = delete;
};

struct PropVariantGuard {
    PROPVARIANT pv;
    PropVariantGuard() { PropVariantInit(&pv); }
    ~PropVariantGuard() { PropVariantClear(&pv); }
    PROPVARIANT* operator&() { return &pv; }
    PropVariantGuard(const PropVariantGuard&) = delete;
    PropVariantGuard& operator=(const PropVariantGuard&) = delete;
};

struct CoTaskMemFreeGuard {
    void* ptr;
    explicit CoTaskMemFreeGuard(void* p) : ptr(p) {}
    ~CoTaskMemFreeGuard() { if (ptr) CoTaskMemFree(ptr); }
    CoTaskMemFreeGuard(const CoTaskMemFreeGuard&) = delete;
    CoTaskMemFreeGuard& operator=(const CoTaskMemFreeGuard&) = delete;
};
