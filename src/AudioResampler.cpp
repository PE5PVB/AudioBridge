#include "AudioResampler.h"
#include <mfapi.h>
#include <mftransform.h>
#include <mferror.h>
#include <wmcodecdsp.h>
#include <ks.h>
#include <ksmedia.h>

AudioResampler::AudioResampler() {}

AudioResampler::~AudioResampler() {
    m_transform.Reset();
}

static bool formatsMatch(const WAVEFORMATEX* a, const WAVEFORMATEX* b) {
    if (a->nSamplesPerSec != b->nSamplesPerSec) return false;
    if (a->nChannels != b->nChannels) return false;
    if (a->wBitsPerSample != b->wBitsPerSample) return false;

    // Check subformat for WAVEFORMATEXTENSIBLE
    if (a->wFormatTag == WAVE_FORMAT_EXTENSIBLE && b->wFormatTag == WAVE_FORMAT_EXTENSIBLE) {
        auto* ea = reinterpret_cast<const WAVEFORMATEXTENSIBLE*>(a);
        auto* eb = reinterpret_cast<const WAVEFORMATEXTENSIBLE*>(b);
        if (memcmp(&ea->SubFormat, &eb->SubFormat, sizeof(GUID)) != 0) return false;
        if (ea->Samples.wValidBitsPerSample != eb->Samples.wValidBitsPerSample) return false;
    }

    return true;
}

HRESULT AudioResampler::init(const WAVEFORMATEX* inputFormat, const WAVEFORMATEX* outputFormat) {
    m_needed = false;

    if (formatsMatch(inputFormat, outputFormat)) {
        return S_FALSE; // No resampling needed
    }

    m_needed = true;

    // Create the MF resampler DSP (CLSID_CResamplerMediaObject)
    RETURN_IF_FAILED(CoCreateInstance(CLSID_CResamplerMediaObject, nullptr,
                                      CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&m_transform)));

    // Set quality to best
    ComPtr<IWMResamplerProps> resamplerProps;
    if (SUCCEEDED(m_transform.As(&resamplerProps))) {
        resamplerProps->SetHalfFilterLength(60); // max quality
    }

    // Set input type
    ComPtr<IMFMediaType> inputType;
    RETURN_IF_FAILED(createMediaType(inputFormat, &inputType));
    RETURN_IF_FAILED(m_transform->SetInputType(0, inputType.Get(), 0));

    // Set output type
    ComPtr<IMFMediaType> outputType;
    RETURN_IF_FAILED(createMediaType(outputFormat, &outputType));
    RETURN_IF_FAILED(m_transform->SetOutputType(0, outputType.Get(), 0));

    // Send stream start message
    RETURN_IF_FAILED(m_transform->ProcessMessage(MFT_MESSAGE_NOTIFY_BEGIN_STREAMING, 0));
    RETURN_IF_FAILED(m_transform->ProcessMessage(MFT_MESSAGE_NOTIFY_START_OF_STREAM, 0));

    return S_OK;
}

HRESULT AudioResampler::createMediaType(const WAVEFORMATEX* wfx, IMFMediaType** ppType) {
    ComPtr<IMFMediaType> type;
    RETURN_IF_FAILED(MFCreateMediaType(&type));

    RETURN_IF_FAILED(type->SetGUID(MF_MT_MAJOR_TYPE, MFMediaType_Audio));

    // Determine subformat
    GUID subType;
    if (wfx->wFormatTag == WAVE_FORMAT_EXTENSIBLE) {
        auto* wfxe = reinterpret_cast<const WAVEFORMATEXTENSIBLE*>(wfx);
        if (IsEqualGUID(wfxe->SubFormat, KSDATAFORMAT_SUBTYPE_IEEE_FLOAT))
            subType = MFAudioFormat_Float;
        else
            subType = MFAudioFormat_PCM;
    } else if (wfx->wFormatTag == WAVE_FORMAT_IEEE_FLOAT) {
        subType = MFAudioFormat_Float;
    } else {
        subType = MFAudioFormat_PCM;
    }

    RETURN_IF_FAILED(type->SetGUID(MF_MT_SUBTYPE, subType));
    RETURN_IF_FAILED(type->SetUINT32(MF_MT_AUDIO_NUM_CHANNELS, wfx->nChannels));
    RETURN_IF_FAILED(type->SetUINT32(MF_MT_AUDIO_SAMPLES_PER_SECOND, wfx->nSamplesPerSec));
    RETURN_IF_FAILED(type->SetUINT32(MF_MT_AUDIO_BLOCK_ALIGNMENT, wfx->nBlockAlign));
    RETURN_IF_FAILED(type->SetUINT32(MF_MT_AUDIO_AVG_BYTES_PER_SECOND, wfx->nAvgBytesPerSec));
    RETURN_IF_FAILED(type->SetUINT32(MF_MT_AUDIO_BITS_PER_SAMPLE, wfx->wBitsPerSample));
    RETURN_IF_FAILED(type->SetUINT32(MF_MT_ALL_SAMPLES_INDEPENDENT, TRUE));

    if (wfx->wFormatTag == WAVE_FORMAT_EXTENSIBLE) {
        auto* wfxe = reinterpret_cast<const WAVEFORMATEXTENSIBLE*>(wfx);
        RETURN_IF_FAILED(type->SetUINT32(MF_MT_AUDIO_VALID_BITS_PER_SAMPLE,
                                          wfxe->Samples.wValidBitsPerSample));
        RETURN_IF_FAILED(type->SetUINT32(MF_MT_AUDIO_CHANNEL_MASK, wfxe->dwChannelMask));
    }

    *ppType = type.Detach();
    return S_OK;
}

HRESULT AudioResampler::process(const BYTE* inData, DWORD inBytes,
                                 std::vector<BYTE>& outBuffer) {
    if (!m_transform) return E_NOT_VALID_STATE;

    // Create input sample
    ComPtr<IMFSample> inputSample;
    RETURN_IF_FAILED(MFCreateSample(&inputSample));

    ComPtr<IMFMediaBuffer> inputBuffer;
    RETURN_IF_FAILED(MFCreateMemoryBuffer(inBytes, &inputBuffer));

    BYTE* bufPtr = nullptr;
    RETURN_IF_FAILED(inputBuffer->Lock(&bufPtr, nullptr, nullptr));
    memcpy(bufPtr, inData, inBytes);
    RETURN_IF_FAILED(inputBuffer->Unlock());
    RETURN_IF_FAILED(inputBuffer->SetCurrentLength(inBytes));

    RETURN_IF_FAILED(inputSample->AddBuffer(inputBuffer.Get()));

    // Feed to transform
    HRESULT hr = m_transform->ProcessInput(0, inputSample.Get(), 0);
    if (FAILED(hr)) return hr;

    // Drain all output
    return drainOutput(outBuffer);
}

HRESULT AudioResampler::drainOutput(std::vector<BYTE>& outBuffer) {
    for (;;) {
        MFT_OUTPUT_STREAM_INFO streamInfo = {};
        RETURN_IF_FAILED(m_transform->GetOutputStreamInfo(0, &streamInfo));

        ComPtr<IMFSample> outputSample;
        ComPtr<IMFMediaBuffer> outputMediaBuf;

        DWORD allocSize = (streamInfo.cbSize > 0) ? streamInfo.cbSize : 65536;

        RETURN_IF_FAILED(MFCreateSample(&outputSample));
        RETURN_IF_FAILED(MFCreateMemoryBuffer(allocSize, &outputMediaBuf));
        RETURN_IF_FAILED(outputSample->AddBuffer(outputMediaBuf.Get()));

        MFT_OUTPUT_DATA_BUFFER outputData = {};
        outputData.pSample = outputSample.Get();

        DWORD status = 0;
        HRESULT hr = m_transform->ProcessOutput(0, 1, &outputData, &status);

        if (hr == MF_E_TRANSFORM_NEED_MORE_INPUT) {
            return S_OK; // No more output available
        }
        RETURN_IF_FAILED(hr);

        // Extract data from output sample
        ComPtr<IMFMediaBuffer> buf;
        RETURN_IF_FAILED(outputData.pSample->ConvertToContiguousBuffer(&buf));

        BYTE* data = nullptr;
        DWORD dataLen = 0;
        RETURN_IF_FAILED(buf->Lock(&data, nullptr, &dataLen));

        size_t prevSize = outBuffer.size();
        outBuffer.resize(prevSize + dataLen);
        memcpy(outBuffer.data() + prevSize, data, dataLen);

        buf->Unlock();
    }
}

HRESULT AudioResampler::flush(std::vector<BYTE>& outBuffer) {
    if (!m_transform) return E_NOT_VALID_STATE;

    RETURN_IF_FAILED(m_transform->ProcessMessage(MFT_MESSAGE_COMMAND_DRAIN, 0));
    return drainOutput(outBuffer);
}
