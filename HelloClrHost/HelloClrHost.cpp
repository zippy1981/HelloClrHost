
#include <tchar.h>
#include<windows.h>
#include <metahost.h>

#pragma comment(lib, "mscoree.lib")

// Import mscorlib.tlb (Microsoft Common Language Runtime Class Library).
#import "mscorlib.tlb" raw_interfaces_only \
	high_property_prefixes("_get","_put","_putref") \
	rename("ReportEvent", "InteropServices_ReportEvent")
using namespace mscorlib;

#define CLR_V40 L"v4.0.30319"

int _tmain(int argc, _TCHAR* argv[])
{
	ICLRMetaHost *pMetaHost = NULL;
    ICLRRuntimeInfo *pRuntimeInfo = NULL;
	IUnknownPtr spAppDomainThunk = NULL;
	_AppDomainPtr spDefaultAppDomain = NULL;

	bstr_t bstrAssemblyName(L"mscorlib");
	_AssemblyPtr spAssembly = NULL;
	
	bstr_t bstrClassName(L"System.String");
	_TypePtr spType = NULL;
	
	bstr_t bstrStaticMethodName(L"Format");

	CLRCreateInstance(CLSID_CLRMetaHost, IID_PPV_ARGS(&pMetaHost));
    pMetaHost->GetRuntime(CLR_V40, IID_PPV_ARGS(&pRuntimeInfo));

	ICorRuntimeHost *pCorRuntimeHost = NULL;
	pRuntimeInfo->GetInterface(CLSID_CorRuntimeHost,  IID_PPV_ARGS(&pCorRuntimeHost));
	pCorRuntimeHost->Start();
	pCorRuntimeHost->GetDefaultDomain(&spAppDomainThunk);
	spAppDomainThunk->QueryInterface(IID_PPV_ARGS(&spDefaultAppDomain));
	spDefaultAppDomain->Load_2(bstrAssemblyName, &spAssembly);
	spAssembly->GetType_2(bstrClassName, &spType);
	wprintf(L"Loaded the assembly %s\n", bstrAssemblyName.GetBSTR());
	{
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

	HRESULT result = spType->InvokeMember_3(bstrStaticMethodName, static_cast<BindingFlags>(
		BindingFlags_InvokeMethod | BindingFlags_Static | BindingFlags_Public), 
		NULL, vtEmpty, psaWriteLineArgs, &vtLengthRet);
	if (FAILED(result))
	{
		wprintf(L"Failed to invoke System.String.Format() w/hr 0x%08lx\n", result);
	}
	return 0;
}

