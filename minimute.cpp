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

struct MicEntry
{
    IAudioEndpointVolume* pVol;
};

const char* g_MiniMute = "MiniMute";
HICON g_hIconMuted;
HICON g_hIconUnmuted;
MicEntry* g_mics = nullptr;
int g_micCount = 0;

void Error(const char *msg)
{
    MessageBox(NULL, msg, g_MiniMute, MB_ICONERROR | MB_OK);
}

void ClearMics()
{
    if (g_mics != nullptr)
    {
        for (int i = 0; i < g_micCount; i++)
        {
            if (g_mics[i].pVol != nullptr)
            {
                g_mics[i].pVol->Release();
                g_mics[i].pVol = nullptr;
            }
        }

        HeapFree(GetProcessHeap(), 0, g_mics);
        g_mics = nullptr;
    }
}

bool ExtractAudioEndpoint(IMMDeviceCollection* pColl, UINT index, MicEntry *pEntry)
{
    IMMDevice* pDevice = nullptr;
    pEntry->pVol = nullptr;
    if (pColl->Item(index, &pDevice) == S_OK)
    {
        IAudioEndpointVolume* pVol = nullptr;
        if (pDevice->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, nullptr, (void**)&pVol) == S_OK)
        {
            pEntry->pVol = pVol;
        }
        else Error("Failed to retrieve volume interface");
        pDevice->Release();
    }
    else Error("Failed to retrieve audio device");

    return pEntry->pVol != nullptr;
}

bool EnumerateMics()
{
    ClearMics();

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
                if (count > 0)
                {
                    g_mics = (MicEntry*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, count * sizeof(MicEntry));
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
            if (g_mics[i].pVol != nullptr)
            {
                if (i == 0)
                {
                    if (g_mics[i].pVol->GetMute(&mute) == S_OK)
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

                if (FAILED(g_mics[i].pVol->SetMute(mute, nullptr)))
                {
                    Error("Failed to mute/unmute at last one microphone");
                    result = 0;
                }
            }
        }
    }

    return result;
}

bool Muted()
{
    bool anyUnmuted = false;
    if (EnumerateMics())
    {
        for (int i = 0; i < g_micCount; i++)
        {
            if (g_mics[i].pVol != nullptr)
            {
                BOOL muted;
                if (g_mics[i].pVol->GetMute(&muted) != S_OK || !muted) anyUnmuted = true;
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

bool SetNotifyIcon(HINSTANCE hInstance, HWND hWnd, bool muted)
{
    NOTIFYICONDATA iconData;
    SetupNotifyIconData(iconData, hWnd, muted);
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
    HWND hWnd;

    if (CoInitializeEx(NULL, COINIT_MULTITHREADED) == S_OK &&
        (hWnd = CreateNotifyWindow(hInstance)) != NULL &&
        InitializeNotifyIcon(hInstance, hWnd))
    {
        if (RegisterHotKey(NULL, 1, 0, VK_PAUSE))
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
