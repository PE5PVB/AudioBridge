#pragma once

#include <windows.h>
#include <audioclient.h>
#include <mmdeviceapi.h>
#include <atomic>
#include <string>
#include "ComHelper.h"
#include "RingBuffer.h"

class WasapiCapture {
public:
    WasapiCapture();
    ~WasapiCapture();

    HRESULT init(const std::wstring& deviceId, bool exclusive, RingBuffer* ringBuffer);
    HRESULT start();
    void    stop();

    const WAVEFORMATEXTENSIBLE& format() const { return m_format; }
    UINT32 bufferFrames() const { return m_bufferFrames; }
    bool   isRunning()    const { return m_running.load(std::memory_order_relaxed); }

private:
    static DWORD WINAPI captureThread(LPVOID param);
    void captureLoop();

    HRESULT initShared();
    HRESULT initExclusive();
    HRESULT negotiateExclusiveFormat();
    HRESULT tryExclusiveFormat(WORD channels, DWORD sampleRate,
                                WORD bitsPerSample, bool isFloat);

    ComPtr<IMMDevice>           m_device;
    ComPtr<IAudioClient>        m_audioClient;
    ComPtr<IAudioCaptureClient> m_captureClient;
    HANDLE                      m_eventHandle = nullptr;
    HANDLE                      m_threadHandle = nullptr;
    HANDLE                      m_stopEvent = nullptr;

    WAVEFORMATEXTENSIBLE m_format = {};
    UINT32               m_bufferFrames = 0;
    bool                 m_exclusive = false;
    std::atomic<bool>    m_running{false};

    RingBuffer*          m_ringBuffer = nullptr;
};
