#include "DeviceEnumerator.h"
#include <functiondiscoverykeys_devpkey.h>

HRESULT DeviceEnumerator::enumerateCapture(std::vector<AudioDeviceInfo>& devices) {
    return enumerate(eCapture, devices);
}

HRESULT DeviceEnumerator::enumerateRender(std::vector<AudioDeviceInfo>& devices) {
    return enumerate(eRender, devices);
}

HRESULT DeviceEnumerator::enumerate(EDataFlow flow, std::vector<AudioDeviceInfo>& devices) {
    devices.clear();

    ComPtr<IMMDeviceEnumerator> enumerator;
    RETURN_IF_FAILED(CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr,
                                      CLSCTX_ALL, IID_PPV_ARGS(&enumerator)));

    ComPtr<IMMDeviceCollection> collection;
    RETURN_IF_FAILED(enumerator->EnumAudioEndpoints(flow, DEVICE_STATE_ACTIVE, &collection));

    UINT count = 0;
    RETURN_IF_FAILED(collection->GetCount(&count));

    for (UINT i = 0; i < count; ++i) {
        ComPtr<IMMDevice> device;
        RETURN_IF_FAILED(collection->Item(i, &device));

        LPWSTR deviceId = nullptr;
        RETURN_IF_FAILED(device->GetId(&deviceId));
        CoTaskMemFreeGuard idGuard(deviceId);

        ComPtr<IPropertyStore> props;
        RETURN_IF_FAILED(device->OpenPropertyStore(STGM_READ, &props));

        PropVariantGuard pvName;
        RETURN_IF_FAILED(props->GetValue(PKEY_Device_FriendlyName, &pvName.pv));

        AudioDeviceInfo info;
        info.id = deviceId;
        info.name = (pvName.pv.vt == VT_LPWSTR) ? pvName.pv.pwszVal : L"(unknown)";
        devices.push_back(std::move(info));
    }

    return S_OK;
}

HRESULT DeviceEnumerator::getDeviceById(const std::wstring& id, EDataFlow /*flow*/,
                                         ComPtr<IMMDevice>& device) {
    ComPtr<IMMDeviceEnumerator> enumerator;
    RETURN_IF_FAILED(CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr,
                                      CLSCTX_ALL, IID_PPV_ARGS(&enumerator)));

    RETURN_IF_FAILED(enumerator->GetDevice(id.c_str(), &device));
    return S_OK;
}
