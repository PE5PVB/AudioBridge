cmake_minimum_required(VERSION 3.20)
project(AudioBridge LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

add_executable(AudioBridge WIN32
    src/main.cpp
    src/DeviceEnumerator.cpp
    src/WasapiCapture.cpp
    src/WasapiRender.cpp
    src/AudioResampler.cpp
    src/AudioRouter.cpp
    src/DialogProc.cpp
    src/AudioBridge.rc
)

target_include_directories(AudioBridge PRIVATE src)

target_link_libraries(AudioBridge PRIVATE
    ole32
    uuid
    avrt
    mfplat
    mfuuid
    wmcodecdspuuid
    propsys
    comctl32
    dwmapi
    uxtheme
    gdi32
    msimg32
    gdiplus
    shell32
)

target_compile_definitions(AudioBridge PRIVATE
    UNICODE
    _UNICODE
    WIN32_LEAN_AND_MEAN
    NOMINMAX
)

# Embed the manifest
set_target_properties(AudioBridge PROPERTIES
    LINK_FLAGS "/MANIFEST:EMBED /MANIFESTINPUT:\"${CMAKE_SOURCE_DIR}/AudioBridge.manifest\""
)
