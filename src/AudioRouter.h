#pragma once

#include <windows.h>
#include <string>
#include <memory>
#include <atomic>
#include "WasapiCapture.h"
#include "WasapiRender.h"
#include "AudioResampler.h"
#include "RingBuffer.h"

enum class RouterState {
    Stopped,
    Running,
    Error
};

struct RouterStatus {
    RouterState state = RouterState::Stopped;
    std::wstring errorMessage;
    WAVEFORMATEXTENSIBLE captureFormat = {};
    WAVEFORMATEXTENSIBLE renderFormat = {};
    UINT32 captureBufferFrames = 0;
    UINT32 renderBufferFrames = 0;
    UINT64 underruns = 0;
    bool   resamplerActive = false;
};

class AudioRouter {
public:
    AudioRouter();
    ~AudioRouter();

    HRESULT start(const std::wstring& captureDeviceId,
                  const std::wstring& renderDeviceId,
                  bool exclusive);
    void    stop();

    RouterStatus getStatus() const;

private:
    static DWORD WINAPI resamplerThread(LPVOID param);
    void resamplerLoop();

    std::unique_ptr<WasapiCapture>  m_capture;
    std::unique_ptr<WasapiRender>   m_render;
    std::unique_ptr<AudioResampler> m_resampler;

    // Ring buffer between capture and render (or capture and resampler)
    std::unique_ptr<RingBuffer> m_captureToRender;
    // Ring buffer between resampler and render (only when resampling)
    std::unique_ptr<RingBuffer> m_resamplerToRender;

    HANDLE m_resamplerThread = nullptr;
    HANDLE m_resamplerStopEvent = nullptr;
    std::atomic<bool> m_resamplerRunning{false};

    std::atomic<RouterState> m_state{RouterState::Stopped};
    std::wstring             m_errorMessage;
};
