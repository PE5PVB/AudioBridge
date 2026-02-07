#include "WasapiRender.h"
#include "DeviceEnumerator.h"
#include <avrt.h>
#include <audioclient.h>

WasapiRender::WasapiRender() {}

WasapiRender::~WasapiRender() {
    stop();
}

HRESULT WasapiRender::init(const std::wstring& deviceId, bool exclusive, RingBuffer* ringBuffer,
                            const WAVEFORMATEXTENSIBLE* preferredFormat) {
    m_exclusive = exclusive;
    m_ringBuffer = ringBuffer;

    if (preferredFormat) {
        m_hasPreferredFormat = true;
        m_preferredFormat = *preferredFormat;
    }

    RETURN_IF_FAILED(DeviceEnumerator::getDeviceById(deviceId, eRender, m_device));
    RETURN_IF_FAILED(m_device->Activate(__uuidof(IAudioClient), CLSCTX_ALL,
                                         nullptr, reinterpret_cast<void**>(m_audioClient.GetAddressOf())));

    m_eventHandle = CreateEventW(nullptr, FALSE, FALSE, nullptr);
    if (!m_eventHandle) return HRESULT_FROM_WIN32(GetLastError());

    m_stopEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    if (!m_stopEvent) return HRESULT_FROM_WIN32(GetLastError());

    if (exclusive) {
        return initExclusive();
    } else {
        return initShared();
    }
}

HRESULT WasapiRender::initShared() {
    WAVEFORMATEX* mixFormat = nullptr;
    RETURN_IF_FAILED(m_audioClient->GetMixFormat(&mixFormat));
    CoTaskMemFreeGuard fmtGuard(mixFormat);

    REFERENCE_TIME defaultPeriod, minPeriod;
    RETURN_IF_FAILED(m_audioClient->GetDevicePeriod(&defaultPeriod, &minPeriod));

    HRESULT hr = m_audioClient->Initialize(
        AUDCLNT_SHAREMODE_SHARED,
        AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
        defaultPeriod, 0, mixFormat, nullptr);
    RETURN_IF_FAILED(hr);

    if (mixFormat->cbSize >= 22 && mixFormat->wFormatTag == WAVE_FORMAT_EXTENSIBLE) {
        memcpy(&m_format, mixFormat, sizeof(WAVEFORMATEXTENSIBLE));
    } else {
        m_format.Format = *mixFormat;
        m_format.Format.wFormatTag = WAVE_FORMAT_EXTENSIBLE;
        m_format.Format.cbSize = 22;
        m_format.Samples.wValidBitsPerSample = mixFormat->wBitsPerSample;
        if (mixFormat->nChannels == 1)
            m_format.dwChannelMask = SPEAKER_FRONT_CENTER;
        else if (mixFormat->nChannels == 2)
            m_format.dwChannelMask = SPEAKER_FRONT_LEFT | SPEAKER_FRONT_RIGHT;
        m_format.SubFormat = (mixFormat->wBitsPerSample == 32)
            ? KSDATAFORMAT_SUBTYPE_IEEE_FLOAT : KSDATAFORMAT_SUBTYPE_PCM;
    }

    RETURN_IF_FAILED(m_audioClient->SetEventHandle(m_eventHandle));
    RETURN_IF_FAILED(m_audioClient->GetBufferSize(&m_bufferFrames));
    RETURN_IF_FAILED(m_audioClient->GetService(IID_PPV_ARGS(&m_renderClient)));

    return S_OK;
}

HRESULT WasapiRender::initExclusive() {
    RETURN_IF_FAILED(negotiateExclusiveFormat());
    RETURN_IF_FAILED(m_audioClient->SetEventHandle(m_eventHandle));
    RETURN_IF_FAILED(m_audioClient->GetBufferSize(&m_bufferFrames));
    RETURN_IF_FAILED(m_audioClient->GetService(IID_PPV_ARGS(&m_renderClient)));
    return S_OK;
}

HRESULT WasapiRender::negotiateExclusiveFormat() {
    // Try preferred format first (capture device's format) for zero-conversion path
    if (m_hasPreferredFormat) {
        bool isFloat = IsEqualGUID(m_preferredFormat.SubFormat, KSDATAFORMAT_SUBTYPE_IEEE_FLOAT);
        HRESULT hr = tryExclusiveFormat(
            m_preferredFormat.Format.nChannels,
            m_preferredFormat.Format.nSamplesPerSec,
            m_preferredFormat.Samples.wValidBitsPerSample,
            isFloat);
        if (SUCCEEDED(hr)) return S_OK;
    }

    // Fall back to standard priority list
    struct FormatAttempt { WORD ch; DWORD rate; WORD bits; bool isFloat; };
    FormatAttempt attempts[] = {
        {2, 48000, 32, true},
        {2, 48000, 24, false},
        {2, 48000, 16, false},
        {2, 44100, 32, true},
        {2, 44100, 24, false},
        {2, 44100, 16, false},
        {1, 48000, 16, false},
        {1, 44100, 16, false},
    };

    for (auto& a : attempts) {
        HRESULT hr = tryExclusiveFormat(a.ch, a.rate, a.bits, a.isFloat);
        if (SUCCEEDED(hr)) return S_OK;
    }

    return AUDCLNT_E_UNSUPPORTED_FORMAT;
}

HRESULT WasapiRender::tryExclusiveFormat(WORD channels, DWORD sampleRate,
                                           WORD bitsPerSample, bool isFloat) {
    WAVEFORMATEXTENSIBLE wfx = {};
    wfx.Format.wFormatTag = WAVE_FORMAT_EXTENSIBLE;
    wfx.Format.nChannels = channels;
    wfx.Format.nSamplesPerSec = sampleRate;
    wfx.Format.wBitsPerSample = (bitsPerSample == 24) ? 32 : bitsPerSample;
    wfx.Format.nBlockAlign = wfx.Format.nChannels * wfx.Format.wBitsPerSample / 8;
    wfx.Format.nAvgBytesPerSec = wfx.Format.nSamplesPerSec * wfx.Format.nBlockAlign;
    wfx.Format.cbSize = 22;
    wfx.Samples.wValidBitsPerSample = bitsPerSample;
    wfx.SubFormat = isFloat ? KSDATAFORMAT_SUBTYPE_IEEE_FLOAT : KSDATAFORMAT_SUBTYPE_PCM;

    if (channels == 1) wfx.dwChannelMask = SPEAKER_FRONT_CENTER;
    else if (channels == 2) wfx.dwChannelMask = SPEAKER_FRONT_LEFT | SPEAKER_FRONT_RIGHT;

    HRESULT hr = m_audioClient->IsFormatSupported(AUDCLNT_SHAREMODE_EXCLUSIVE,
                                                   &wfx.Format, nullptr);
    if (hr != S_OK) return AUDCLNT_E_UNSUPPORTED_FORMAT;

    REFERENCE_TIME defaultPeriod, minPeriod;
    RETURN_IF_FAILED(m_audioClient->GetDevicePeriod(&defaultPeriod, &minPeriod));

    REFERENCE_TIME requestedDuration = minPeriod;

    hr = m_audioClient->Initialize(
        AUDCLNT_SHAREMODE_EXCLUSIVE,
        AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
        requestedDuration, requestedDuration, &wfx.Format, nullptr);

    if (hr == AUDCLNT_E_BUFFER_SIZE_NOT_ALIGNED) {
        UINT32 alignedFrames = 0;
        RETURN_IF_FAILED(m_audioClient->GetBufferSize(&alignedFrames));

        m_audioClient.Reset();
        RETURN_IF_FAILED(m_device->Activate(__uuidof(IAudioClient), CLSCTX_ALL,
                                             nullptr, reinterpret_cast<void**>(m_audioClient.GetAddressOf())));

        requestedDuration = static_cast<REFERENCE_TIME>(
            10000000.0 * alignedFrames / wfx.Format.nSamplesPerSec + 0.5);

        hr = m_audioClient->Initialize(
            AUDCLNT_SHAREMODE_EXCLUSIVE,
            AUDCLNT_STREAMFLAGS_EVENTCALLBACK,
            requestedDuration, requestedDuration, &wfx.Format, nullptr);
    }

    if (FAILED(hr)) return hr;

    m_format = wfx;
    return S_OK;
}

HRESULT WasapiRender::start() {
    if (m_running.load()) return S_FALSE;

    m_running.store(true, std::memory_order_release);
    m_underruns.store(0, std::memory_order_relaxed);
    ResetEvent(m_stopEvent);

    m_threadHandle = CreateThread(nullptr, 0, renderThread, this, 0, nullptr);
    if (!m_threadHandle) {
        m_running.store(false);
        return HRESULT_FROM_WIN32(GetLastError());
    }

    return S_OK;
}

void WasapiRender::stop() {
    if (!m_running.load()) return;

    m_running.store(false, std::memory_order_release);
    SetEvent(m_stopEvent);

    if (m_threadHandle) {
        WaitForSingleObject(m_threadHandle, 5000);
        CloseHandle(m_threadHandle);
        m_threadHandle = nullptr;
    }

    if (m_audioClient) {
        m_audioClient->Stop();
    }

    if (m_eventHandle) {
        CloseHandle(m_eventHandle);
        m_eventHandle = nullptr;
    }
    if (m_stopEvent) {
        CloseHandle(m_stopEvent);
        m_stopEvent = nullptr;
    }
}

DWORD WINAPI WasapiRender::renderThread(LPVOID param) {
    CoInitializeGuard comGuard(COINIT_MULTITHREADED);
    auto* self = static_cast<WasapiRender*>(param);
    self->renderLoop();
    return 0;
}

void WasapiRender::renderLoop() {
    DWORD taskIndex = 0;
    HANDLE mmcssHandle = AvSetMmThreadCharacteristicsW(L"Pro Audio", &taskIndex);

    // Pre-roll: fill buffer with silence before starting
    {
        BYTE* data = nullptr;
        HRESULT hr = m_renderClient->GetBuffer(m_bufferFrames, &data);
        if (SUCCEEDED(hr)) {
            memset(data, 0, static_cast<size_t>(m_bufferFrames) * m_format.Format.nBlockAlign);
            m_renderClient->ReleaseBuffer(m_bufferFrames, AUDCLNT_BUFFERFLAGS_SILENT);
        }
    }

    HRESULT hr = m_audioClient->Start();
    if (FAILED(hr)) {
        m_running.store(false);
        if (mmcssHandle) AvRevertMmThreadCharacteristics(mmcssHandle);
        return;
    }

    HANDLE waitHandles[2] = { m_stopEvent, m_eventHandle };

    while (m_running.load(std::memory_order_relaxed)) {
        DWORD waitResult = WaitForMultipleObjects(2, waitHandles, FALSE, 2000);

        if (waitResult == WAIT_OBJECT_0) {
            break;
        }
        if (waitResult != WAIT_OBJECT_0 + 1 && waitResult != WAIT_TIMEOUT) {
            break;
        }

        UINT32 padding = 0;
        if (m_exclusive) {
            // In exclusive mode, buffer is always fully available after event
            padding = 0;
        } else {
            hr = m_audioClient->GetCurrentPadding(&padding);
            if (FAILED(hr)) break;
        }

        UINT32 framesAvailable = m_bufferFrames - padding;
        if (framesAvailable == 0) continue;

        BYTE* data = nullptr;
        hr = m_renderClient->GetBuffer(framesAvailable, &data);
        if (FAILED(hr)) continue;

        size_t bytesNeeded = static_cast<size_t>(framesAvailable) * m_format.Format.nBlockAlign;
        size_t bytesRead = m_ringBuffer->read(data, bytesNeeded);

        if (bytesRead < bytesNeeded) {
            // Underrun: fill remainder with silence
            memset(data + bytesRead, 0, bytesNeeded - bytesRead);
            m_underruns.fetch_add(1, std::memory_order_relaxed);
            if (bytesRead == 0) {
                m_renderClient->ReleaseBuffer(framesAvailable, AUDCLNT_BUFFERFLAGS_SILENT);
            } else {
                m_renderClient->ReleaseBuffer(framesAvailable, 0);
            }
        } else {
            m_renderClient->ReleaseBuffer(framesAvailable, 0);
        }
    }

    m_audioClient->Stop();
    if (mmcssHandle) AvRevertMmThreadCharacteristics(mmcssHandle);
}
