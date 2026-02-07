#include "WasapiCapture.h"
#include "DeviceEnumerator.h"
#include <avrt.h>
#include <audioclient.h>

WasapiCapture::WasapiCapture() {}

WasapiCapture::~WasapiCapture() {
    stop();
}

HRESULT WasapiCapture::init(const std::wstring& deviceId, bool exclusive, RingBuffer* ringBuffer) {
    m_exclusive = exclusive;
    m_ringBuffer = ringBuffer;

    RETURN_IF_FAILED(DeviceEnumerator::getDeviceById(deviceId, eCapture, m_device));
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

HRESULT WasapiCapture::initShared() {
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

    // Store format
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
    RETURN_IF_FAILED(m_audioClient->GetService(IID_PPV_ARGS(&m_captureClient)));

    return S_OK;
}

HRESULT WasapiCapture::initExclusive() {
    RETURN_IF_FAILED(negotiateExclusiveFormat());
    RETURN_IF_FAILED(m_audioClient->SetEventHandle(m_eventHandle));
    RETURN_IF_FAILED(m_audioClient->GetBufferSize(&m_bufferFrames));
    RETURN_IF_FAILED(m_audioClient->GetService(IID_PPV_ARGS(&m_captureClient)));
    return S_OK;
}

HRESULT WasapiCapture::negotiateExclusiveFormat() {
    // Try formats in priority order
    struct FormatAttempt { WORD ch; DWORD rate; WORD bits; bool isFloat; };
    FormatAttempt attempts[] = {
        {2, 48000, 32, true},   // 32-bit float 48kHz stereo
        {2, 48000, 24, false},  // 24-bit PCM 48kHz
        {2, 48000, 16, false},  // 16-bit PCM 48kHz
        {2, 44100, 32, true},   // 32-bit float 44.1kHz
        {2, 44100, 24, false},  // 24-bit PCM 44.1kHz
        {2, 44100, 16, false},  // 16-bit PCM 44.1kHz
        {1, 48000, 16, false},  // mono fallbacks
        {1, 44100, 16, false},
    };

    for (auto& a : attempts) {
        HRESULT hr = tryExclusiveFormat(a.ch, a.rate, a.bits, a.isFloat);
        if (SUCCEEDED(hr)) return S_OK;
    }

    return AUDCLNT_E_UNSUPPORTED_FORMAT;
}

HRESULT WasapiCapture::tryExclusiveFormat(WORD channels, DWORD sampleRate,
                                            WORD bitsPerSample, bool isFloat) {
    WAVEFORMATEXTENSIBLE wfx = {};
    wfx.Format.wFormatTag = WAVE_FORMAT_EXTENSIBLE;
    wfx.Format.nChannels = channels;
    wfx.Format.nSamplesPerSec = sampleRate;
    wfx.Format.wBitsPerSample = (bitsPerSample == 24) ? 32 : bitsPerSample; // container size
    wfx.Format.nBlockAlign = wfx.Format.nChannels * wfx.Format.wBitsPerSample / 8;
    wfx.Format.nAvgBytesPerSec = wfx.Format.nSamplesPerSec * wfx.Format.nBlockAlign;
    wfx.Format.cbSize = 22;
    wfx.Samples.wValidBitsPerSample = bitsPerSample;
    wfx.SubFormat = isFloat ? KSDATAFORMAT_SUBTYPE_IEEE_FLOAT : KSDATAFORMAT_SUBTYPE_PCM;

    if (channels == 1) wfx.dwChannelMask = SPEAKER_FRONT_CENTER;
    else if (channels == 2) wfx.dwChannelMask = SPEAKER_FRONT_LEFT | SPEAKER_FRONT_RIGHT;

    // Check if format is supported
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

    // Handle buffer alignment dance
    if (hr == AUDCLNT_E_BUFFER_SIZE_NOT_ALIGNED) {
        UINT32 alignedFrames = 0;
        RETURN_IF_FAILED(m_audioClient->GetBufferSize(&alignedFrames));

        // Release and recreate audio client
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

HRESULT WasapiCapture::start() {
    if (m_running.load()) return S_FALSE;

    m_running.store(true, std::memory_order_release);
    ResetEvent(m_stopEvent);

    m_threadHandle = CreateThread(nullptr, 0, captureThread, this, 0, nullptr);
    if (!m_threadHandle) {
        m_running.store(false);
        return HRESULT_FROM_WIN32(GetLastError());
    }

    return S_OK;
}

void WasapiCapture::stop() {
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

DWORD WINAPI WasapiCapture::captureThread(LPVOID param) {
    CoInitializeGuard comGuard(COINIT_MULTITHREADED);
    auto* self = static_cast<WasapiCapture*>(param);
    self->captureLoop();
    return 0;
}

void WasapiCapture::captureLoop() {
    // Boost thread priority for pro audio
    DWORD taskIndex = 0;
    HANDLE mmcssHandle = AvSetMmThreadCharacteristicsW(L"Pro Audio", &taskIndex);

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
            // Stop event signaled
            break;
        }
        if (waitResult != WAIT_OBJECT_0 + 1 && waitResult != WAIT_TIMEOUT) {
            break;
        }

        // Read available capture packets
        UINT32 packetLength = 0;
        while (SUCCEEDED(m_captureClient->GetNextPacketSize(&packetLength)) && packetLength > 0) {
            BYTE* data = nullptr;
            UINT32 framesAvailable = 0;
            DWORD flags = 0;

            hr = m_captureClient->GetBuffer(&data, &framesAvailable, &flags, nullptr, nullptr);
            if (FAILED(hr)) break;

            size_t byteCount = static_cast<size_t>(framesAvailable) * m_format.Format.nBlockAlign;

            if (flags & AUDCLNT_BUFFERFLAGS_SILENT) {
                // Write silence to the ring buffer
                std::vector<uint8_t> silence(byteCount, 0);
                m_ringBuffer->write(silence.data(), byteCount);
            } else {
                m_ringBuffer->write(data, byteCount);
            }

            m_captureClient->ReleaseBuffer(framesAvailable);
        }
    }

    m_audioClient->Stop();
    if (mmcssHandle) AvRevertMmThreadCharacteristics(mmcssHandle);
}
