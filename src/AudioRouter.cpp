#include "AudioRouter.h"
#include <mfapi.h>

AudioRouter::AudioRouter() {}

AudioRouter::~AudioRouter() {
    stop();
}

HRESULT AudioRouter::start(const std::wstring& captureDeviceId,
                            const std::wstring& renderDeviceId,
                            bool exclusive) {
    // Stop any existing session
    stop();

    m_errorMessage.clear();

    // MFStartup for resampler
    HRESULT hr = MFStartup(MF_VERSION);
    if (FAILED(hr)) {
        m_errorMessage = L"MFStartup mislukt";
        m_state.store(RouterState::Error);
        return hr;
    }

    // Ring buffer: 500ms at 48kHz stereo 32-bit float = ~192KB
    // Generous size to absorb jitter between capture and render clocks
    const size_t ringBufferSize = 48000 * 8 * 500 / 1000;
    m_captureToRender = std::make_unique<RingBuffer>(ringBufferSize);

    // Init capture
    m_capture = std::make_unique<WasapiCapture>();
    hr = m_capture->init(captureDeviceId, exclusive, m_captureToRender.get());
    if (FAILED(hr)) {
        m_errorMessage = L"Capture init mislukt (0x" + std::to_wstring(hr) + L")";
        m_state.store(RouterState::Error);
        return hr;
    }

    // Init render - pass capture format as preferred so render tries it first
    // This maximizes the chance both devices use the same format (no resampling needed)
    m_render = std::make_unique<WasapiRender>();
    hr = m_render->init(renderDeviceId, exclusive, m_captureToRender.get(),
                        &m_capture->format());
    if (FAILED(hr)) {
        m_errorMessage = L"Render init mislukt (0x" + std::to_wstring(hr) + L")";
        m_state.store(RouterState::Error);
        return hr;
    }

    // Check if resampling is needed between capture and render formats
    m_resampler = std::make_unique<AudioResampler>();
    hr = m_resampler->init(&m_capture->format().Format, &m_render->format().Format);

    if (hr == S_FALSE || !m_resampler->isNeeded()) {
        // No resampling needed - render reads directly from captureToRender (already set)
        m_resampler.reset();
        m_resamplerToRender.reset();
    } else if (SUCCEEDED(hr)) {
        // Resampling needed - redirect render to read from resampler output buffer
        m_resamplerToRender = std::make_unique<RingBuffer>(ringBufferSize);
        m_render->setRingBuffer(m_resamplerToRender.get());
    } else {
        m_errorMessage = L"Resampler init mislukt (0x" + std::to_wstring(hr) + L")";
        m_state.store(RouterState::Error);
        return hr;
    }

    // Start resampler thread if needed
    if (m_resampler && m_resampler->isNeeded()) {
        m_resamplerStopEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
        m_resamplerRunning.store(true);
        m_resamplerThread = CreateThread(nullptr, 0, resamplerThread, this, 0, nullptr);
    }

    // Start capture FIRST so the ring buffer fills up
    hr = m_capture->start();
    if (FAILED(hr)) {
        m_errorMessage = L"Capture start mislukt";
        stop();
        m_state.store(RouterState::Error);
        return hr;
    }

    // Pre-buffer: wait until ring buffer has enough data before starting render.
    // Target: 2x the render buffer size, so render never starves on first callback.
    RingBuffer* renderSource = m_resamplerToRender ? m_resamplerToRender.get()
                                                   : m_captureToRender.get();
    size_t preBufferTarget = static_cast<size_t>(m_render->bufferFrames())
                             * m_render->format().Format.nBlockAlign * 2;
    for (int wait = 0; wait < 500; ++wait) { // max 500ms wachten
        if (renderSource->availableToRead() >= preBufferTarget)
            break;
        Sleep(1);
    }

    hr = m_render->start();
    if (FAILED(hr)) {
        m_errorMessage = L"Render start mislukt";
        stop();
        m_state.store(RouterState::Error);
        return hr;
    }

    m_state.store(RouterState::Running);
    return S_OK;
}

void AudioRouter::stop() {
    // Stop resampler thread
    if (m_resamplerRunning.load()) {
        m_resamplerRunning.store(false);
        if (m_resamplerStopEvent) SetEvent(m_resamplerStopEvent);
        if (m_resamplerThread) {
            WaitForSingleObject(m_resamplerThread, 5000);
            CloseHandle(m_resamplerThread);
            m_resamplerThread = nullptr;
        }
        if (m_resamplerStopEvent) {
            CloseHandle(m_resamplerStopEvent);
            m_resamplerStopEvent = nullptr;
        }
    }

    if (m_capture) {
        m_capture->stop();
        m_capture.reset();
    }

    if (m_render) {
        m_render->stop();
        m_render.reset();
    }

    m_resampler.reset();
    m_captureToRender.reset();
    m_resamplerToRender.reset();

    MFShutdown();

    m_state.store(RouterState::Stopped);
}

RouterStatus AudioRouter::getStatus() const {
    RouterStatus status;
    status.state = m_state.load();
    status.errorMessage = m_errorMessage;

    if (m_capture) {
        status.captureFormat = m_capture->format();
        status.captureBufferFrames = m_capture->bufferFrames();
    }
    if (m_render) {
        status.renderFormat = m_render->format();
        status.renderBufferFrames = m_render->bufferFrames();
        status.underruns = m_render->underrunCount();
    }
    if (m_resampler) {
        status.resamplerActive = m_resampler->isNeeded();
    }

    return status;
}

DWORD WINAPI AudioRouter::resamplerThread(LPVOID param) {
    CoInitializeGuard comGuard(COINIT_MULTITHREADED);
    auto* self = static_cast<AudioRouter*>(param);
    self->resamplerLoop();
    return 0;
}

void AudioRouter::resamplerLoop() {
    // Process audio from captureToRender → resampler → resamplerToRender
    const size_t chunkSize = 4096; // Process in 4KB chunks
    std::vector<BYTE> inBuf(chunkSize);
    std::vector<BYTE> outBuf;

    while (m_resamplerRunning.load(std::memory_order_relaxed)) {
        size_t avail = m_captureToRender->availableToRead();
        if (avail == 0) {
            // Wait a short time for data
            WaitForSingleObject(m_resamplerStopEvent, 1);
            continue;
        }

        size_t toRead = (std::min)(avail, chunkSize);
        size_t bytesRead = m_captureToRender->read(inBuf.data(), toRead);

        if (bytesRead > 0) {
            outBuf.clear();
            HRESULT hr = m_resampler->process(inBuf.data(), static_cast<DWORD>(bytesRead), outBuf);
            if (SUCCEEDED(hr) && !outBuf.empty()) {
                m_resamplerToRender->write(outBuf.data(), outBuf.size());
            }
        }
    }
}
