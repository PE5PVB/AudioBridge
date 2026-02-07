#pragma once

#include <string>
#include <vector>
#include <mmdeviceapi.h>
#include "ComHelper.h"

struct AudioDeviceInfo {
    std::wstring id;
    std::wstring name;
};

class DeviceEnumerator {
public:
    static HRESULT enumerateCapture(std::vector<AudioDeviceInfo>& devices);
    static HRESULT enumerateRender(std::vector<AudioDeviceInfo>& devices);

    static HRESULT getDeviceById(const std::wstring& id, EDataFlow flow,
                                  ComPtr<IMMDevice>& device);
private:
    static HRESULT enumerate(EDataFlow flow, std::vector<AudioDeviceInfo>& devices);
};
