#pragma once

#include <windows.h>
#include <audioclient.h>
#include <mmdeviceapi.h>
#include <atomic>
#include <string>
#include "ComHelper.h"
#include "RingBuffer.h"

class WasapiRender {
public:
    WasapiRender();
    ~WasapiRender();

    HRESULT init(const std::wstring& deviceId, bool exclusive, RingBuffer* ringBuffer,
                 const WAVEFORMATEXTENSIBLE* preferredFormat = nullptr);
    HRESULT start();
    void    stop();
    void    setRingBuffer(RingBuffer* rb) { m_ringBuffer = rb; }

    const WAVEFORMATEXTENSIBLE& format() const { return m_format; }
    UINT32 bufferFrames()  const { return m_bufferFrames; }
    bool   isRunning()     const { return m_running.load(std::memory_order_relaxed); }
    UINT64 underrunCount() const { return m_underruns.load(std::memory_order_relaxed); }

private:
    static DWORD WINAPI renderThread(LPVOID param);
    void renderLoop();

    HRESULT initShared();
    HRESULT initExclusive();
    HRESULT negotiateExclusiveFormat();
    HRESULT tryExclusiveFormat(WORD channels, DWORD sampleRate,
                                WORD bitsPerSample, bool isFloat);

    ComPtr<IMMDevice>          m_device;
    ComPtr<IAudioClient>       m_audioClient;
    ComPtr<IAudioRenderClient> m_renderClient;
    HANDLE                     m_eventHandle = nullptr;
    HANDLE                     m_threadHandle = nullptr;
    HANDLE                     m_stopEvent = nullptr;

    WAVEFORMATEXTENSIBLE m_format = {};
    UINT32               m_bufferFrames = 0;
    bool                 m_exclusive = false;
    bool                 m_hasPreferredFormat = false;
    WAVEFORMATEXTENSIBLE m_preferredFormat = {};
    std::atomic<bool>    m_running{false};
    std::atomic<UINT64>  m_underruns{0};

    RingBuffer*          m_ringBuffer = nullptr;
};
