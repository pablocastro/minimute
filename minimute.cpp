#include <initguid.h>
#include <combaseapi.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <CommCtrl.h>
#include <strsafe.h>
#include "resource.h"

#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

DEFINE_GUID(IID_IAudioEndpointVolume, 0x5CDF2C82, 0x841E, 0x4546, 0x97, 0x22, 0x0C, 0xF7, 0x40, 0x78, 0x22, 0x9A);

const char* MiniMute = "MiniMute";
HICON hIconMuted;
HICON hIconUnmuted;

void Error(const char *msg)
{
    MessageBox(NULL, msg, MiniMute, MB_ICONERROR | MB_OK);
}

template<typename F>
bool EnumerateMics(F action)
{
    IMMDeviceEnumerator* pEnum = nullptr;
    bool result = false;
    if (CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnum) == S_OK)
    {
        IMMDeviceCollection* pColl = nullptr;
        if (pEnum->EnumAudioEndpoints(eCapture, DEVICE_STATE_ACTIVE, &pColl) == S_OK)
        {
            UINT count = 0;
            if (pColl->GetCount(&count) == S_OK)
            {
                if (count == 0)
                {
                    Error("No audio devices");
                }
                else
                {
                    result = true;
                    for (UINT i = 0; i < count; i++)
                    {
                        IMMDevice* pDevice = nullptr;
                        if (pColl->Item(i, &pDevice) == S_OK)
                        {
                            IAudioEndpointVolume* pVol = nullptr;
                            if (pDevice->Activate(IID_IAudioEndpointVolume, CLSCTX_ALL, nullptr, (void**)&pVol) == S_OK)
                            {
                                action(i, pVol);
                                pVol->Release();
                            }
                            else
                            {
                                Error("Failed to retrieve volume interface");
                                result = false;
                                break;
                            }
                            pDevice->Release();
                        }
                        else
                        {
                            Error("Failed to retrieve audio device");
                            result = false;
                            break;
                        }
                    }
                }
            }
            else Error("Failed to get audio endpoint count");
            pColl->Release();
        }
        else Error("Failed to enumerate audio endpoints");
        pEnum->Release();
    }
    else Error("Failed to create MM device enumerator");

    return result;
}

// 0 not sure, 1 muted, 2 unmuted
int FlipMute()
{
    BOOL mute = true;
    int result = 0;
    EnumerateMics([&mute, &result](int index, IAudioEndpointVolume* pVol)
        {
            if (index == 0)
            {
                if (pVol->GetMute(&mute) == S_OK)
                {
                    mute = !mute;
                    result = mute ? 1 : 2;
                }
                else
                {
                    Error("Failed to get muted state");
                    mute = true;
                }
            }

            if (FAILED(pVol->SetMute(mute, nullptr)))
            {
                Error("Failed to mute/unmute at last one microphone");
                result = 0;
            }
        }
    );

    return result;
}

bool Muted()
{
    bool anyUnmuted = false;
    int result = EnumerateMics([&anyUnmuted](int, IAudioEndpointVolume* pVol)
        {
            BOOL muted;
            if (pVol->GetMute(&muted) != S_OK || !muted) anyUnmuted = true;
        }
    );

    return result && !anyUnmuted;
}

LRESULT CALLBACK NotifyWindowProcedure(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    return DefWindowProc(hWnd, uMsg, wParam, lParam);
}

HWND CreateNotifyWindow(HINSTANCE hInstance)
{
    WNDCLASSEX wndClass;
    SecureZeroMemory(&wndClass, sizeof(wndClass));
    wndClass.cbSize = sizeof(wndClass);
    wndClass.lpfnWndProc = NotifyWindowProcedure;
    wndClass.lpszClassName = MiniMute;
    wndClass.hInstance = hInstance;
    wndClass.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_APP));
    ATOM classAtom = RegisterClassEx(&wndClass);
    if (classAtom == 0) return NULL;

    return CreateWindowEx(0, (LPCTSTR)classAtom, MiniMute, 0, 0, 0, 0, 0, HWND_MESSAGE, NULL, hInstance, NULL);
}

void SetupNotifyIconData(NOTIFYICONDATA &iconData, HWND hWnd, bool muted)
{
    SecureZeroMemory(&iconData, sizeof(iconData));
    iconData.cbSize = sizeof(iconData);
    iconData.hWnd = hWnd;
    iconData.uID = 1;
    iconData.uFlags = NIF_ICON | NIF_TIP | NIF_SHOWTIP;
    iconData.hIcon = muted ? hIconMuted : hIconUnmuted;
    StringCchCopy(iconData.szTip, sizeof(iconData.szTip) / sizeof(TCHAR), muted ? "Muted" : "Unmuted");
}

bool SetNotifyIcon(HINSTANCE hInstance, HWND hWnd, bool muted)
{
    NOTIFYICONDATA iconData;
    SetupNotifyIconData(iconData, hWnd, muted);
    return Shell_NotifyIcon(NIM_MODIFY, &iconData);
}

bool InitializeNotifyIcon(HINSTANCE hInstance, HWND hWnd)
{
    if (LoadIconMetric(hInstance, MAKEINTRESOURCEW(IDI_MUTED), LIM_SMALL, &hIconMuted) != S_OK ||
        LoadIconMetric(hInstance, MAKEINTRESOURCEW(IDI_UNMUTED), LIM_SMALL, &hIconUnmuted) != S_OK) return false;

    NOTIFYICONDATA iconData;
    SetupNotifyIconData(iconData, hWnd, false);
    return Shell_NotifyIcon(NIM_ADD, &iconData);
}

int __stdcall main()
{
    HINSTANCE hInstance = GetModuleHandle(nullptr);
    HWND hWnd;

    if (CoInitializeEx(NULL, COINIT_MULTITHREADED) == S_OK &&
        (hWnd = CreateNotifyWindow(hInstance)) != NULL &&
        InitializeNotifyIcon(hInstance, hWnd))
    {
        if (RegisterHotKey(NULL, 1, 0, VK_SCROLL))
        {
            SetNotifyIcon(hInstance, hWnd, Muted());

            MSG msg;
            while (GetMessage(&msg, NULL, 0, 0) != 0)
            {
                if (msg.message == WM_HOTKEY)
                {
                    int r = FlipMute();
                    SetNotifyIcon(hInstance, hWnd, r == 1);
                }
            }
        }
        else Error("Failed to register hot key. Another instance already running?");
    }
    else Error("Failed to initialize");
    CoUninitialize();
    ExitProcess(0);
}
