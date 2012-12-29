// HelloClrHost.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"

#pragma comment(lib, "mscoree.lib")

// Import mscorlib.tlb (Microsoft Common Language Runtime Class Library).
#import "mscorlib.tlb" raw_interfaces_only \
	high_property_prefixes("_get","_put","_putref") \
	rename("ReportEvent", "InteropServices_ReportEvent")
using namespace mscorlib;

#define CLR_V10 L"v1.0.3705"
#define CLR_V11 L"v1.1.4322"
#define CLR_V20 L"v2.0.50727"
#define CLR_V30 L"v3.0"
#define CLR_V35 L"v3.5"
#define CLR_V40 L"v4.0.30319"

int _tmain(int argc, _TCHAR* argv[])
{
	ICLRMetaHost *pMetaHost = NULL;
    ICLRRuntimeInfo *pRuntimeInfo = NULL;
	IUnknownPtr spAppDomainThunk = NULL;
	_AppDomainPtr spDefaultAppDomain = NULL;

	PCWSTR pszAssemblyName = L"mscorlib";
	bstr_t bstrAssemblyName(pszAssemblyName);
	_AssemblyPtr spAssembly = NULL;
	
	bstr_t bstrClassName(L"System.String");
	_TypePtr spType = NULL;
	
	bstr_t bstrStaticMethodName(L"Format");
	

	HRESULT result;
	BOOL fLoadable;

	if (FAILED(result = CLRCreateInstance(CLSID_CLRMetaHost, IID_PPV_ARGS(&pMetaHost))))
    {
        fwprintf(stderr, L"CLRCreateInstance failed w/hr 0x%08lx\n", result);
    }
	else if (FAILED(result = pMetaHost->GetRuntime(CLR_V40, IID_PPV_ARGS(&pRuntimeInfo)))) {
		fwprintf(stderr, L"ICLRMetaHost::GetRuntime failed w/hr 0x%08lx\n", result);
	}
	else if (FAILED(result = pRuntimeInfo->IsLoadable(&fLoadable)))
	{
		fwprintf(stderr, L"ICLRRuntimeInfo::IsLoadable failed w/hr 0x%08lx\n", result);
	}
	else {
		LPWSTR pwzVersion;
		DWORD versionStrSize;
		HANDLE hProcHeap = GetProcessHeap();
		pRuntimeInfo->GetVersionString(NULL, &versionStrSize);
		pwzVersion = (LPWSTR) HeapAlloc(hProcHeap, HEAP_GENERATE_EXCEPTIONS | HEAP_ZERO_MEMORY, versionStrSize * sizeof(TCHAR));
		pRuntimeInfo->GetVersionString(pwzVersion, &versionStrSize);
		if (!fLoadable)
		{
			fwprintf(stderr, L".NET runtime %s cannot be loaded\n\r", pwzVersion);
		}
		else
		{
			wprintf(L".NET runtime %s loaded successfully\n\r", pwzVersion);

			ICorRuntimeHost *pCorRuntimeHost = NULL;
			if (FAILED(result = pRuntimeInfo->GetInterface(CLSID_CorRuntimeHost,  IID_PPV_ARGS(&pCorRuntimeHost))))
			{
				wprintf(L"ICLRRuntimeInfo::GetInterface failed w/hr 0x%08lx\n", result);
			}
			else if (FAILED(result = pCorRuntimeHost->Start()))
			{
				wprintf(L"CLR failed to start w/hr 0x%08lx\n", result);
			}
			else if(result = pCorRuntimeHost->GetDefaultDomain(&spAppDomainThunk))
			{
				wprintf(L"ICorRuntimeHost::GetDefaultDomain failed w/hr 0x%08lx\n", result);
			}
			else if (FAILED(result = spAppDomainThunk->QueryInterface(IID_PPV_ARGS(&spDefaultAppDomain))))
			{
				wprintf(L"Failed to get default AppDomain w/hr 0x%08lx\n", result);
			}
			else if (FAILED(result = spDefaultAppDomain->Load_2(bstrAssemblyName, &spAssembly)))
			{
				wprintf(L"Failed to load the assembly w/hr 0x%08lx\n", result);
			}
			else if (FAILED(result = spAssembly->GetType_2(bstrClassName, &spType)) || spType == NULL)
			{
				wprintf(L"Failed to get the Type interface w/hr 0x%08lx\n", result);

			}
			else {
				wprintf(L"Loaded the assembly %s\n", pszAssemblyName);

				{
					//TODO: replace with _bstr_t
					BSTR fullName;
					spType->get_FullName(&fullName);
					wprintf(L"Got type %s\n", fullName);
					SysFreeString(fullName);
				}

				SAFEARRAY *psaWriteLineArgs = SafeArrayCreateVector(VT_VARIANT, 0, 2);
				variant_t vtEmpty;
				variant_t vtLengthRet;
				{
					LONG index = 0;
					LONG age = 31;
					SafeArrayPutElement(psaWriteLineArgs, &index, L"My Name is {0}. My Age is {1}");
					index = 1;
					SafeArrayPutElement(psaWriteLineArgs, &index, L"Justin Dearing");
					index = 2;
					SafeArrayPutElement(psaWriteLineArgs, &index, &age);
				}

				result = spType->InvokeMember_3(bstrStaticMethodName, static_cast<BindingFlags>(
					BindingFlags_InvokeMethod | BindingFlags_Static | BindingFlags_Public), 
					NULL, vtEmpty, psaWriteLineArgs, &vtLengthRet);
				if (FAILED(result))
				{
					wprintf(L"Failed to invoke System.Console.WriteLine() w/hr 0x%08lx\n", result);
				}
			}
		}
		HeapFree(hProcHeap, 0, pwzVersion);

		
	}

	return 0;
}

