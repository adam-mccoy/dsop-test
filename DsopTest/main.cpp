#include <Windows.h>
#include <ObjIdl.h>
#include <ObjSel.h>

void ShowPicker(HWND hwnd)
{
    HRESULT hr;
    DSOP_INIT_INFO initInfo = { 0 };
    DSOP_SCOPE_INIT_INFO scopeInfo = { 0 };
    IDataObject *pDataObject;
    IDsObjectPicker *pDsObjectPicker = NULL;

    hr = CoCreateInstance(CLSID_DsObjectPicker, NULL, CLSCTX_INPROC_SERVER, IID_IDsObjectPicker, (void**)&pDsObjectPicker);

    scopeInfo.cbSize = sizeof(DSOP_SCOPE_INIT_INFO);
    scopeInfo.flType =
        DSOP_SCOPE_TYPE_GLOBAL_CATALOG |
        DSOP_SCOPE_TYPE_TARGET_COMPUTER |
        DSOP_SCOPE_TYPE_UPLEVEL_JOINED_DOMAIN |
        DSOP_SCOPE_TYPE_DOWNLEVEL_JOINED_DOMAIN |
        DSOP_SCOPE_TYPE_ENTERPRISE_DOMAIN |
        DSOP_SCOPE_TYPE_EXTERNAL_UPLEVEL_DOMAIN |
        DSOP_SCOPE_TYPE_EXTERNAL_DOWNLEVEL_DOMAIN |
        DSOP_SCOPE_TYPE_USER_ENTERED_UPLEVEL_SCOPE |
        DSOP_SCOPE_TYPE_USER_ENTERED_DOWNLEVEL_SCOPE;
    scopeInfo.FilterFlags.Uplevel.flBothModes = DSOP_FILTER_USERS;
    scopeInfo.FilterFlags.Uplevel.flMixedModeOnly = DSOP_FILTER_USERS;
    scopeInfo.FilterFlags.Uplevel.flNativeModeOnly = DSOP_FILTER_USERS;
    scopeInfo.FilterFlags.flDownlevel = DSOP_DOWNLEVEL_FILTER_USERS;

    initInfo.cbSize = sizeof(DSOP_INIT_INFO);
    initInfo.pwzTargetComputer = NULL;
    initInfo.cDsScopeInfos = 1;
    initInfo.aDsScopeInfos = &scopeInfo;

    hr = pDsObjectPicker->Initialize(&initInfo);
    if (!SUCCEEDED(hr))
        MessageBox(NULL, L"Failed to Initialize Picker", L"Error!", MB_ICONEXCLAMATION | MB_OK);

    hr = pDsObjectPicker->InvokeDialog(hwnd, &pDataObject);
    if (!SUCCEEDED(hr))
        MessageBox(NULL, L"Failed to Invoke Picker", L"Error!", MB_ICONEXCLAMATION | MB_OK);
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_CLOSE:
        DestroyWindow(hWnd);
        break;

    case WM_DESTROY:
        PostQuitMessage(0);
        break;

    case WM_KEYUP:
        ShowPicker(hWnd);
        break;

    default:
        return DefWindowProc(hWnd, msg, wParam, lParam);
    }
    return 0;
}

int __stdcall WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
    WNDCLASSEX wc;
    HWND hwnd;
    MSG msg;

    HRESULT hr;

    hr = CoInitialize(NULL);
    if (!SUCCEEDED(hr))
        return -1;

    wc.cbSize = sizeof(WNDCLASSEX);
    wc.style = 0;
    wc.lpfnWndProc = WndProc;
    wc.cbClsExtra = 0;
    wc.cbWndExtra = 0;
    wc.hInstance = hInstance;
    wc.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszMenuName = NULL;
    wc.lpszClassName = L"My Window Class";
    wc.hIconSm = LoadIcon(NULL, IDI_APPLICATION);

    if (!RegisterClassEx(&wc))
    {
        MessageBox(NULL, L"Window Registration Failed!", L"Error!", MB_ICONEXCLAMATION | MB_OK);
        return 0;
    }

    hwnd = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        L"My Window Class",
        L"The Title of my Window",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 240, 120,
        NULL, NULL, hInstance, NULL);

    if (hwnd == NULL)
    {
        MessageBox(NULL, L"Window Creation Failed!", L"Error!", MB_ICONEXCLAMATION | MB_OK);
        return 0;
    }

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    while (GetMessage(&msg, NULL, 0, 0) > 0)
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return msg.wParam;
}
