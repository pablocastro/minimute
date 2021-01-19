#define __INLINE_ISEQUAL_GUID // avoid dependency on memcmp
#include <initguid.h>
#include <combaseapi.h>
#include <mmdeviceapi.h>
#include <endpointvolume.h>
#include <CommCtrl.h>
#include <strsafe.h>
#include "resource.h"

// To keep binary and runtime size to a minimum, minimute doesn't depend on the C runtime. Most of the weirdness
// of this code comes from not being able to use the CRT (no new/delete/malloc, no exceptions, etc.)

#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#define WM_ENUMERATE_MICS (WM_USER + 1)

bool Muted(bool rescan);
bool SetNotifyIcon(bool muted);
bool EnumerateMics();

// Lots of simplifying assumptions for minimalism: not thread safe (running in STA), not honoring the Release() behavior
class CEndpointCallback : public IAudioEndpointVolumeCallback, public IMMNotificationClient
{
    LONG _refCount;

public:
    CEndpointCallback() : _refCount(1) {}

    STDMETHOD(QueryInterface)(REFIID riid, void** ppvObject)
    {
        if (IID_IUnknown == riid)
        {
            AddRef();
            *ppvObject = (IUnknown*)(IAudioEndpointVolumeCallback*)this;
        }
        else if (__uuidof(IAudioEndpointVolumeCallback) == riid)
        {
            AddRef();
            *ppvObject = (IAudioEndpointVolumeCallback*)this;
        }
        else if (__uuidof(IMMNotificationClient) == riid)
        {
            AddRef();
            *ppvObject = (IMMNotificationClient*)this;
        }
        else
        {
            *ppvObject = nullptr;
            return E_NOINTERFACE;
        }
        return S_OK;
    }

    STDMETHOD_(ULONG, AddRef)()
    {
        return ++_refCount;
    }

    STDMETHOD_(ULONG, Release)()
    {
        if (_refCount > 0) _refCount--;
        return _refCount;
    }

    STDMETHOD(OnNotify)(PAUDIO_VOLUME_NOTIFICATION_DATA pNotify)
    {
        SetNotifyIcon(Muted(false));
        return S_OK;
    }

    STDMETHOD(OnDeviceStateChanged)(LPCWSTR pwstrDeviceId, DWORD dwNewState)
    {
        PostMessage(NULL, WM_ENUMERATE_MICS, 0, 0);
        return S_OK;
    }

    STDMETHOD(OnDeviceAdded)(LPCWSTR pwstrDeviceId)
    {
        PostMessage(NULL, WM_ENUMERATE_MICS, 0, 0);
        return S_OK;
    }

    STDMETHOD(OnDeviceRemoved)(LPCWSTR pwstrDeviceId)
    {
        PostMessage(NULL, WM_ENUMERATE_MICS, 0, 0);
        return S_OK;
    }

    STDMETHOD(OnDefaultDeviceChanged)(EDataFlow flow, ERole role, LPCWSTR pwstrDefaultDeviceId)
    {
        return S_OK;
    }

    STDMETHOD(OnPropertyValueChanged)(LPCWSTR pwstrDeviceId, const PROPERTYKEY key)
    {
        return S_OK;
    }
};

const char* g_MiniMute = "MiniMute";
HICON g_hIconMuted;
HICON g_hIconUnmuted;
HWND g_hNotificationWindow;
IAudioEndpointVolume** g_mics = nullptr;
int g_micCount = 0;
CEndpointCallback *g_endpointCallback;
bool registeredForNotifications = false;

void Error(const char *msg, UINT type = MB_ICONERROR)
{
    MessageBox(NULL, msg, g_MiniMute, type | MB_OK);
}

void ClearMics()
{
    if (g_mics != nullptr)
    {
        for (int i = 0; i < g_micCount; i++)
        {
            if (g_mics[i] != nullptr)
            {
                g_mics[i]->UnregisterControlChangeNotify(g_endpointCallback);
                g_mics[i]->Release();
                g_mics[i]= nullptr;
            }
        }

        HeapFree(GetProcessHeap(), 0, g_mics);
        g_mics = nullptr;
    }
}

bool ExtractAudioEndpoint(IMMDeviceCollection* pColl, UINT index, IAudioEndpointVolume **ppVol)
{
    IMMDevice* pDevice = nullptr;
    *ppVol = nullptr;
    if (pColl->Item(index, &pDevice) == S_OK)
    {
        if (pDevice->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, nullptr, (void**)ppVol) == S_OK)
        {
            if ((*ppVol)->RegisterControlChangeNotify(g_endpointCallback) != S_OK)
            {
                Error("Failed to register for change notifications. Mute/unmute may continue to work, but mute state in traybar may not be accurate.",
                      MB_ICONWARNING);
            }
        }
        else Error("Failed to retrieve volume interface");
        pDevice->Release();
    }
    else Error("Failed to retrieve audio device");

    return *ppVol != nullptr;
}

bool EnumerateMics()
{
    ClearMics();

    IMMDeviceEnumerator* pEnum = nullptr;
    bool result = false;
    if (CoCreateInstance(__uuidof(MMDeviceEnumerator), NULL, CLSCTX_ALL, __uuidof(IMMDeviceEnumerator), (void**)&pEnum) == S_OK)
    {
        if (!registeredForNotifications)
        {
            if (SUCCEEDED(pEnum->RegisterEndpointNotificationCallback(g_endpointCallback))) registeredForNotifications = true;
            else Error("Warning: failed to register for device notifications, new devices may not be detected automatically.", MB_ICONWARNING);
        }

        IMMDeviceCollection* pColl = nullptr;
        if (pEnum->EnumAudioEndpoints(eCapture, DEVICE_STATE_ACTIVE, &pColl) == S_OK)
        {
            UINT count = 0;
            if (pColl->GetCount(&count) == S_OK)
            {
                if (count > 0)
                {
                    g_mics = (IAudioEndpointVolume**)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, count * sizeof(IAudioEndpointVolume*));
                    if (g_mics != nullptr) 
                    {
                        g_micCount = count;
                        result = true;
                        for (UINT i = 0; i < count; i++)
                        {
                            if (!ExtractAudioEndpoint(pColl, i, &g_mics[i]))
                            {
                                result = false;
                                break;
                            }
                        }
                    }
                    else Error("Not enough memory to enumerate devices");
                }
                else Error("No audio devices");
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
    if (EnumerateMics())
    {
        for (int i = 0; i < g_micCount; i++)
        {
            if (g_mics[i] != nullptr)
            {
                if (i == 0)
                {
                    if (g_mics[i]->GetMute(&mute) == S_OK)
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

                if (FAILED(g_mics[i]->SetMute(mute, nullptr)))
                {
                    Error("Failed to mute/unmute at last one microphone");
                    result = 0;
                }
            }
        }
    }

    return result;
}

bool Muted(bool rescan)
{
    bool anyUnmuted = false;
    if (!rescan || EnumerateMics())
    {
        for (int i = 0; i < g_micCount; i++)
        {
            if (g_mics[i] != nullptr)
            {
                BOOL muted;
                if (g_mics[i]->GetMute(&muted) != S_OK || !muted) anyUnmuted = true;
            }
        }

        return !anyUnmuted;
    }

    return false;
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
    wndClass.lpszClassName = g_MiniMute;
    wndClass.hInstance = hInstance;
    wndClass.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_APP));
    ATOM classAtom = RegisterClassEx(&wndClass);
    if (classAtom == 0) return NULL;

    return CreateWindowEx(0, (LPCTSTR)classAtom, g_MiniMute, 0, 0, 0, 0, 0, HWND_MESSAGE, NULL, hInstance, NULL);
}

void SetupNotifyIconData(NOTIFYICONDATA &iconData, HWND hWnd, bool muted)
{
    SecureZeroMemory(&iconData, sizeof(iconData));
    iconData.cbSize = sizeof(iconData);
    iconData.hWnd = hWnd;
    iconData.uID = 1;
    iconData.uFlags = NIF_ICON | NIF_TIP | NIF_SHOWTIP;
    iconData.hIcon = muted ? g_hIconMuted : g_hIconUnmuted;
    StringCchCopy(iconData.szTip, sizeof(iconData.szTip) / sizeof(TCHAR), muted ? "Muted" : "Unmuted");
}

bool SetNotifyIcon(bool muted)
{
    NOTIFYICONDATA iconData;
    SetupNotifyIconData(iconData, g_hNotificationWindow, muted);
    return Shell_NotifyIcon(NIM_MODIFY, &iconData);
}

bool InitializeNotifyIcon(HINSTANCE hInstance, HWND hWnd)
{
    if (LoadIconMetric(hInstance, MAKEINTRESOURCEW(IDI_MUTED), LIM_SMALL, &g_hIconMuted) != S_OK ||
        LoadIconMetric(hInstance, MAKEINTRESOURCEW(IDI_UNMUTED), LIM_SMALL, &g_hIconUnmuted) != S_OK) return false;

    NOTIFYICONDATA iconData;
    SetupNotifyIconData(iconData, hWnd, false);
    return Shell_NotifyIcon(NIM_ADD, &iconData);
}

int __stdcall main()
{
    HINSTANCE hInstance = GetModuleHandle(nullptr);

    static CEndpointCallback endpointCallback;
    g_endpointCallback = &endpointCallback;

    if (CoInitializeEx(NULL, COINIT_APARTMENTTHREADED) == S_OK &&
        (g_hNotificationWindow = CreateNotifyWindow(hInstance)) != NULL &&
        InitializeNotifyIcon(hInstance, g_hNotificationWindow))
    {
        if (RegisterHotKey(NULL, 1, 0, VK_PAUSE))
        {
            SetNotifyIcon(Muted(true));

            MSG msg;
            while (GetMessage(&msg, NULL, 0, 0) != 0)
            {
                if (msg.message == WM_HOTKEY)
                {
                    int r = FlipMute();
                    SetNotifyIcon(r == 1);
                }
                else if (msg.message == WM_ENUMERATE_MICS)
                {
                    EnumerateMics();
                }

                DispatchMessage(&msg);
            }
        }
        else Error("Failed to register hot key. Another instance already running?");
    }
    else Error("Failed to initialize");
    CoUninitialize();
    ExitProcess(0);
}
