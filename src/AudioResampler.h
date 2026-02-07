#pragma once

#include <windows.h>
#include <mfapi.h>
#include <mftransform.h>
#include <mmreg.h>
#include <vector>
#include "ComHelper.h"

class AudioResampler {
public:
    AudioResampler();
    ~AudioResampler();

    // Initialize with input and output WAVEFORMATEX.
    // Returns S_FALSE if no resampling is needed (formats match).
    HRESULT init(const WAVEFORMATEX* inputFormat, const WAVEFORMATEX* outputFormat);

    // Process a block of audio data. Output is appended to outBuffer.
    HRESULT process(const BYTE* inData, DWORD inBytes,
                    std::vector<BYTE>& outBuffer);

    // Flush any remaining data in the resampler.
    HRESULT flush(std::vector<BYTE>& outBuffer);

    bool isNeeded() const { return m_needed; }

private:
    HRESULT createMediaType(const WAVEFORMATEX* wfx, IMFMediaType** ppType);
    HRESULT drainOutput(std::vector<BYTE>& outBuffer);

    ComPtr<IMFTransform> m_transform;
    bool                 m_needed = false;
    DWORD                m_outputStreamId = 0;
};
