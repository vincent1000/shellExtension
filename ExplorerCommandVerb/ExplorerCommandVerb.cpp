// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved

// ExplorerCommand handlers are an inproc verb implementation method that can provide
// dynamic behavior including computing the name of the command, its icon and its visibility state.
// only use this verb implemetnation method if you are implementing a command handler on
// the commands module and need the same functionality on a context menu.
//
// each ExplorerCommand handler needs to have a unique COM object, run uuidgen to
// create new CLSID values for your handler. a handler can implement multiple
// different verbs using the information provided via IInitializeCommand (the verb name).
// your code can switch off those different verb names or the properties provided
// in the property bag

#include "Dll.h"
#include <iostream>
#include <fstream>
#include <Windows.h>
#include <codecvt>
#include <filesystem>
#include <string>
#include <winreg.h>
static WCHAR const c_szVerbDisplayName[] = L"This is a custom context menu...";
static WCHAR const c_szVerbName[] = L"Sample.ExplorerCommandVerb";


class CExplorerCommandVerb : public IExplorerCommand,
	public IInitializeCommand,
	public IObjectWithSite
{
public:
	CExplorerCommandVerb() : _cRef(1), _punkSite(NULL), _hwnd(NULL), _pstmShellItemArray(NULL)
	{
		DllAddRef();
	}

	// IUnknown
	IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
	{
		static const QITAB qit[] =
		{
			QITABENT(CExplorerCommandVerb, IExplorerCommand),       // required
			QITABENT(CExplorerCommandVerb, IInitializeCommand),     // optional
			QITABENT(CExplorerCommandVerb, IObjectWithSite),        // optional
			{ 0 },
		};
		return QISearch(this, qit, riid, ppv);
	}

	IFACEMETHODIMP_(ULONG) AddRef()
	{
		return InterlockedIncrement(&_cRef);
	}

	IFACEMETHODIMP_(ULONG) Release()
	{
		long cRef = InterlockedDecrement(&_cRef);
		if (!cRef)
		{
			delete this;
		}
		return cRef;
	}

	// IExplorerCommand
	IFACEMETHODIMP GetTitle(IShellItemArray* /* psiItemArray */, LPWSTR* ppszName)
	{
		// the verb name can be computed here, in this example it is static
		return SHStrDup(c_szVerbDisplayName, ppszName);
	}

	IFACEMETHODIMP GetIcon(IShellItemArray* /* psiItemArray */, LPWSTR* ppszIcon)
	{
			HKEY hKey;
			LPCWSTR keyPath = L"SOFTWARE\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\CurrentVersion\\AppModel\\PackageRepository\\Packages";
		
			// Open the key for reading
			if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, keyPath, 0, KEY_READ, &hKey) != ERROR_SUCCESS) {
				std::cerr << "Failed to open registry key." << std::endl;
				return E_NOTIMPL;
			}
			std::wstring subkeyNameToFind = L"App";
			std::wstring subkeyNameToExclude = L"neutral";
			std::wstring retValue = L"";
			WCHAR subkeyName[MAX_PATH];
			WCHAR value[MAX_PATH];
			DWORD index = 0;
		
			// Enumerate the subkeys to find the one starting with "abc_"
			while (RegEnumKey(hKey, index, subkeyName, MAX_PATH) != ERROR_NO_MORE_ITEMS) {
				std::wstring subkey(subkeyName);
				if ((subkey.find(subkeyNameToFind) == 0) && (subkey.find(subkeyNameToExclude) == std::wstring::npos))
				{
					HKEY hSubKey;
					if (RegOpenKeyEx(hKey, subkey.c_str(), 0, KEY_READ, &hSubKey) == ERROR_SUCCESS) {
						DWORD valueLength = MAX_PATH;
						// Query the value named "Path"
						if (RegQueryValueEx(hSubKey, L"Path", NULL, NULL, (LPBYTE)value, &valueLength) == ERROR_SUCCESS) {
							std::wcout << L"Value of 'Path' under subkey '" << subkey << L"' is: " << value << std::endl;
							//MessageBox(hwnd, value, L"ExplorerCommand Sample Verb", MB_OK);
							retValue = value;
						}
						RegCloseKey(hSubKey);
					}
				}
		
				index++;
			}
		
			RegCloseKey(hKey);
			std::wstring regPath = retValue + L"\\Images\\SmallTile.scale-200.png";
			std::filesystem::path currentPath = std::filesystem::current_path();
			
			iconPath = regPath;
		
		*ppszIcon = const_cast<LPWSTR>(iconPath.c_str());
		return S_OK;
	}

	IFACEMETHODIMP GetToolTip(IShellItemArray* /* psiItemArray */, LPWSTR* ppszInfotip)
	{
		// tooltip provided here, in this case none is provieded
		*ppszInfotip = NULL;
		return E_NOTIMPL;
	}

	IFACEMETHODIMP GetCanonicalName(GUID* pguidCommandName)
	{
		*pguidCommandName = __uuidof(this);
		return S_OK;
	}

	// compute the visibility of the verb here, respect "fOkToBeSlow" if this is slow (does IO for example)
	// when called with fOkToBeSlow == FALSE return E_PENDING and this object will be called
	// back on a background thread with fOkToBeSlow == TRUE
	IFACEMETHODIMP GetState(IShellItemArray* /* psiItemArray */, BOOL fOkToBeSlow, EXPCMDSTATE* pCmdState)
	{
		//HRESULT hr;
		//if (fOkToBeSlow)
		//{
		//    Sleep(4 * 1000);    // simulate expensive work
		//    *pCmdState = ECS_ENABLED;
		//    hr = S_OK;
		//}
		//else
		//{
		//    *pCmdState = ECS_DISABLED;
		//    // returning E_PENDING requests that a new instance of this object be called back
		//    // on a background thread so that it can do work that might be slow
		//    hr = E_PENDING;
		//}
		if (fOkToBeSlow)
		{
			
		}
		*pCmdState = ECS_ENABLED;
		HRESULT hr = S_OK;
		return hr;
	}

	IFACEMETHODIMP Invoke(IShellItemArray* psiItemArray, IBindCtx* pbc);

	IFACEMETHODIMP GetFlags(EXPCMDFLAGS* pFlags)
	{
		*pFlags = ECF_DEFAULT;
		return S_OK;
	}

	IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** ppEnum)
	{
		*ppEnum = NULL;
		return E_NOTIMPL;
	}

	// IInitializeCommand
	IFACEMETHODIMP Initialize(PCWSTR /* pszCommandName */, IPropertyBag* /* ppb */)
	{
		// the verb name is in pszCommandName, this handler can vary its behavior
		// based on the command name (implementing different verbs) or the
		// data stored under that verb in the registry can be read via ppb
		return S_OK;
	}

	// IObjectWithSite
	IFACEMETHODIMP SetSite(IUnknown* punkSite)
	{
		SetInterface(&_punkSite, punkSite);
		return S_OK;
	}

	IFACEMETHODIMP GetSite(REFIID riid, void** ppv)
	{
		*ppv = NULL;
		return _punkSite ? _punkSite->QueryInterface(riid, ppv) : E_FAIL;
	}

private:
	~CExplorerCommandVerb()
	{
		SafeRelease(&_punkSite);
		SafeRelease(&_pstmShellItemArray);
		DllRelease();
	}

	DWORD _ThreadProc();

	static DWORD __stdcall s_ThreadProc(void* pv)
	{
		CExplorerCommandVerb* pecv = (CExplorerCommandVerb*)pv;
		const DWORD ret = pecv->_ThreadProc();
		pecv->Release();
		return ret;
	}

	long _cRef;
	IUnknown* _punkSite;
	HWND _hwnd;
	IStream* _pstmShellItemArray;
};
std::wstring GetPackagePath(HWND hwnd)
{
	HKEY hKey;
	LPCWSTR keyPath = L"SOFTWARE\\Classes\\Local Settings\\Software\\Microsoft\\Windows\\CurrentVersion\\AppModel\\PackageRepository\\Packages";

	// Open the key for reading
	if (RegOpenKeyEx(HKEY_LOCAL_MACHINE, keyPath, 0, KEY_READ, &hKey) != ERROR_SUCCESS) {
		std::cerr << "Failed to open registry key." << std::endl;
		MessageBox(hwnd, L"fail open", L"ExplorerCommand Sample Verb", MB_OK);
		return L"";
	}
	std::wstring subkeyNameToFind = L"App";
	std::wstring subkeyNameToExclude = L"neutral";
	std::wstring retValue = L"";
	WCHAR subkeyName[MAX_PATH];
	WCHAR value[MAX_PATH];
	DWORD index = 0;

	// Enumerate the subkeys to find the one starting with "abc_"
	while (RegEnumKey(hKey, index, subkeyName, MAX_PATH) != ERROR_NO_MORE_ITEMS) {
		std::wstring subkey(subkeyName);
		if ((subkey.find(subkeyNameToFind) == 0) && (subkey.find(subkeyNameToExclude) == std::wstring::npos))
		{
			HKEY hSubKey;
			if (RegOpenKeyEx(hKey, subkey.c_str(), 0, KEY_READ, &hSubKey) == ERROR_SUCCESS) {
				DWORD valueLength = MAX_PATH;
				// Query the value named "Path"
				if (RegQueryValueEx(hSubKey, L"Path", NULL, NULL, (LPBYTE)value, &valueLength) == ERROR_SUCCESS) {
					std::wcout << L"Value of 'Path' under subkey '" << subkey << L"' is: " << value << std::endl;
					//MessageBox(hwnd, value, L"ExplorerCommand Sample Verb", MB_OK);
					retValue = value;
					//MessageBox(hwnd, retValue.c_str(), L"ExplorerCommand Sample Verb", MB_OK);
				}
				RegCloseKey(hSubKey);
			}
		}

		index++;
	}

	RegCloseKey(hKey);
	return retValue;
}
DWORD CExplorerCommandVerb::_ThreadProc()
{
	STARTUPINFO si;
	PROCESS_INFORMATION pi;
	std::wstring regPath = GetPackagePath(_hwnd) + L"\\App.exe";
	WCHAR path[MAX_PATH] = { 0 };
	wmemcpy_s(path, MAX_PATH, regPath.c_str(), regPath.size());
	
	ZeroMemory(&si, sizeof(si));
	si.cb = sizeof(si);
	ZeroMemory(&pi, sizeof(pi));
	bool ret = CreateProcess(
		NULL,           // No module name (use command line)
		path,        // Command line
		NULL,           // Process handle not inheritable
		NULL,           // Thread handle not inheritable
		FALSE,          // Set handle inheritance to FALSE
		0,              // No creation flags
		NULL,           // Use parent's environment block
		NULL,           // Use parent's starting directory
		&si,            // Pointer to STARTUPINFO structure
		&pi             // Pointer to PROCESS_INFORMATION structure
	);
	
	// Start the application
	if (!ret)
	{
		std::cout << "CreateProcess failed: " << GetLastError() << std::endl;
		return 1;
	}
	
	// Close process and thread handles
	CloseHandle(pi.hProcess);
	CloseHandle(pi.hThread);
	
	return 0;
}
IFACEMETHODIMP CExplorerCommandVerb::Invoke(IShellItemArray* psia, IBindCtx* /* pbc */)
{
	IUnknown_GetWindow(_punkSite, &_hwnd);

	HRESULT hr = CoMarshalInterThreadInterfaceInStream(__uuidof(psia), psia, &_pstmShellItemArray);
	if (SUCCEEDED(hr))
	{
		AddRef();
		if (!SHCreateThread(s_ThreadProc, this, CTF_COINIT_STA | CTF_PROCESS_REF, NULL))
		{
			Release();
		}
	}
	return S_OK;
}

static WCHAR const c_szProgID[] = L"txtfile";

HRESULT CExplorerCommandVerb_RegisterUnRegister(bool fRegister)
{
	CRegisterExtension re(__uuidof(CExplorerCommandVerb));
	HRESULT hr;
	if (fRegister)
	{
		hr = re.RegisterInProcServer(c_szVerbDisplayName, L"Apartment");
		if (SUCCEEDED(hr))
		{
			// register this verb on .txt files ProgID
			hr = re.RegisterExplorerCommandVerb(c_szProgID, c_szVerbName, c_szVerbDisplayName);
			if (SUCCEEDED(hr))
			{
				hr = re.RegisterVerbAttribute(c_szProgID, c_szVerbName, L"NeverDefault");
			}
			if (SUCCEEDED(hr))
			{
				hr = re.RegisterVerbAttribute(c_szProgID, c_szVerbName, L"Position", L"Bottom");
			}
		}
	}
	else
	{
		// best effort
		hr = re.UnRegisterVerb(c_szProgID, c_szVerbName);
		hr = re.UnRegisterObject();
	}
	return hr;
}

HRESULT CExplorerCommandVerb_CreateInstance(REFIID riid, void** ppv)
{
	*ppv = NULL;
	CExplorerCommandVerb* pVerb = new (std::nothrow) CExplorerCommandVerb();
	HRESULT hr = pVerb ? S_OK : E_OUTOFMEMORY;
	if (SUCCEEDED(hr))
	{
		pVerb->QueryInterface(riid, ppv);
		pVerb->Release();
	}
	return hr;
}
