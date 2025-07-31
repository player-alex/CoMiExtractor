#include <Windows.h>
#include <atlbase.h>
#include <algorithm>
#include <functional>
#include <iostream>
#include <iomanip>
#include <map>
#include <string>
#include <vector>
#include "CoMiExtractor.h"
#include "Utils.h"

#define RETURN_ON_FAIL(hr) \
    if (FAILED(hr)) {      \
        return hr;         \
    }

CoMiExtractor::CoMiExtractor(std::string clsidStr) {
    _clsidStr = clsidStr;
    _clsid = Utils::StringToCLSID(_clsidStr.c_str());
}

CoMiExtractor::~CoMiExtractor() { }

void CoMiExtractor::_GetMethodOffsets(const char* iidStr) {
    HRESULT hr;
    GUID iid;

    hr = CLSIDFromString(
        Utils::MultiByteToWideCharString(iidStr).c_str(), &iid);

    if (FAILED(hr)) {
        return;
    }

    CComPtr<IUnknown> pInterface;
    hr = _pUnk->QueryInterface(iid, reinterpret_cast<void**>(&pInterface));

    if (FAILED(hr) || !pInterface) {
        return;
    }

    void** vtable = *reinterpret_cast<void***>(pInterface.p);

    for (size_t i = 0; i < _interfaceInfos[iidStr].methodInfos.size(); ++i) {
        int offset = _interfaceInfos[iidStr].methodInfos[i].vTableOffset / sizeof(void*);

        _interfaceInfos[iidStr].methodInfos[i].rvaOffset =
            Utils::CalculateRVAFromAddress(
                reinterpret_cast<LPVOID>(vtable[offset]));
    }
}

void CoMiExtractor::_ExtractParams(
    FUNCDESC* pFuncDesc, 
    BSTR* argNames, 
    UINT nameCount, 
    std::vector<ParamInfo>& paramInfos) 
{
    for (SHORT j = 0; j < pFuncDesc->cParams; ++j) {
        ParamInfo paramInfo = {};
        PARAMDESC paramDesc = pFuncDesc->lprgelemdescParam[j].paramdesc;
        TYPEDESC& typeDesc = pFuncDesc->lprgelemdescParam[j].tdesc;

        paramInfo.name = (j + 1 < static_cast<SHORT>(nameCount)) ?
            Utils::WideCharToMultiByteString(argNames[j + 1]) : "arg";
        paramInfo.direction = paramDesc.wParamFlags;
        paramInfo.paramType = typeDesc.vt;

        paramInfos.push_back(paramInfo);
    }
}

HRESULT CoMiExtractor::_ExtractMethodInfo(ITypeInfo* pTypeInfo, UINT index, MethodInfo* pMethodInfo) {
    HRESULT hr;
    FUNCDESC* pFuncDesc = nullptr;

    hr = pTypeInfo->GetFuncDesc(index, &pFuncDesc);
    RETURN_ON_FAIL(hr);

    auto funcDescReleaser = [&](FUNCDESC* pFuncDesc) {
        if (pFuncDesc) {
            pTypeInfo->ReleaseFuncDesc(pFuncDesc);
        }
        };
    std::unique_ptr<FUNCDESC, decltype(funcDescReleaser)> funcDescGuard(
        pFuncDesc, funcDescReleaser);

    UINT nameCount = pFuncDesc->cParams + 1;
    std::unique_ptr<BSTR[], BSTRArrayDeleter> argNames(
        new BSTR[nameCount](),
        BSTRArrayDeleter(nameCount));
    
    hr = pTypeInfo->GetNames(pFuncDesc->memid, argNames.get(), nameCount, &nameCount);

    if (SUCCEEDED(hr)) {
        pMethodInfo->name = Utils::WideCharToMultiByteString(argNames[0]);
        pMethodInfo->invokeKind = pFuncDesc->invkind;
        pMethodInfo->retType = pFuncDesc->elemdescFunc.tdesc.vt;
        pMethodInfo->vTableOffset = pFuncDesc->oVft;
        pMethodInfo->dispId = pFuncDesc->memid;
        _ExtractParams(pFuncDesc, argNames.get(), nameCount, pMethodInfo->paramInfos);
    }

    return hr;
}

HRESULT CoMiExtractor::_InspectCoClass(ITypeInfo* pTypeInfo, TYPEATTR* pTypeAttr) {
    std::string guid;

    for (UINT i = 0; i < pTypeAttr->cImplTypes; ++i) {
        HREFTYPE href = 0;
        HRESULT hr = pTypeInfo->GetRefTypeOfImplType(i, &href);
        RETURN_ON_FAIL(hr);

        CComPtr<ITypeInfo> pImplTypeInfo;
        hr = pTypeInfo->GetRefTypeInfo(href, &pImplTypeInfo);
        RETURN_ON_FAIL(hr);

        TYPEATTR* pImplAttr = nullptr;
        hr = pImplTypeInfo->GetTypeAttr(&pImplAttr);
        RETURN_ON_FAIL(hr);

        auto implAttrReleaser = [&](TYPEATTR* attr) {
            if (pImplTypeInfo && attr) {
                pImplTypeInfo->ReleaseTypeAttr(attr);
            }
        };
        std::unique_ptr<TYPEATTR, decltype(implAttrReleaser)> 
            implAttrGuard(pImplAttr, implAttrReleaser);

        hr = _InspectTypeInfo(pImplTypeInfo, guid);

        RETURN_ON_FAIL(hr);
    }

    return S_OK;
}

HRESULT CoMiExtractor::_InspectTypeInfo(ITypeInfo* pTypeInfo, std::string& guid) {
    TYPEATTR* pTypeAttr = nullptr;
    HRESULT hr = pTypeInfo->GetTypeAttr(&pTypeAttr);

    if (FAILED(hr) || !pTypeAttr) {
        return hr;
    }

    auto typeAttrReleaser = [&](TYPEATTR* x) {
        if (x) {
            pTypeInfo->ReleaseTypeAttr(x);
        }
        };
    std::unique_ptr<TYPEATTR, decltype(typeAttrReleaser)> typeAttrGuard(
        pTypeAttr, typeAttrReleaser);

    if (pTypeAttr->typekind == TKIND_COCLASS) {
        return _InspectCoClass(pTypeInfo, pTypeAttr);
    }
    else {
        guid = Utils::GetGuidStringFromTypeAttr(pTypeAttr);
        _interfaceInfos.insert({ guid, {} });
        _interfaceInfos[guid].name = Utils::GetTypeName(pTypeInfo);

        for (UINT i = 0; i < pTypeAttr->cFuncs; ++i) {
            MethodInfo methodInfo = {};
            hr = _ExtractMethodInfo(pTypeInfo, i, &methodInfo);

            if (SUCCEEDED(hr)) {
                _interfaceInfos[guid].methodInfos.push_back(methodInfo);
            }
        }

        _GetMethodOffsets(guid.c_str());
    }

    return hr;
}

HRESULT CoMiExtractor::Extract() {
    HRESULT hr = S_OK;
    MULTI_QI mqi = { &IID_IUnknown, nullptr, S_OK };
    COSERVERINFO csi = {};
    std::string guid = _clsidStr;

    hr = CoCreateInstanceEx(_clsid, nullptr, CLSCTX_INPROC_SERVER, &csi, 1, &mqi);
    RETURN_ON_FAIL(hr);

    _pUnk = static_cast<IUnknown*>(mqi.pItf);
    using QueryFunc = std::function<HRESULT(ITypeInfo**)>;
    std::vector<QueryFunc> queryFuncs = {
        ([this](ITypeInfo** pTypeInfo)
        { return this->_QueryToIProvideClassInfo(pTypeInfo);  }),
        ([this](ITypeInfo** pTypeInfo)
        { return this->_QueryToIDispatch(pTypeInfo);  })
    };

    for (const auto& f : queryFuncs) {
        CComPtr<ITypeInfo> pTypeInfo;
        hr = f(&pTypeInfo);
        
        if (SUCCEEDED(hr)) {
            hr = _InspectTypeInfo(pTypeInfo, guid);
            _PrintInterfaceInfos();
        }
        else {
            std::cout << "Failed to _InspectTypeInfo: " 
                << Utils::HResultToComErrorMessage(hr) << std::endl;
        }

        _interfaceInfos.clear();
    }

    return hr;
}

HRESULT CoMiExtractor::_QueryToIProvideClassInfo(ITypeInfo** pTypeInfo) {
    HRESULT hr = S_OK;
    CComPtr<IProvideClassInfo> pci;
    std::string guid = _clsidStr;

    hr = _pUnk->QueryInterface(
        IID_IProvideClassInfo, 
        reinterpret_cast<void**>(&pci));
    RETURN_ON_FAIL(hr);

    return pci->GetClassInfoW(pTypeInfo);
}

HRESULT CoMiExtractor::_QueryToIDispatch(ITypeInfo** pTypeInfo) {
    HRESULT hr = S_OK;
    CComPtr<IDispatch> pDisp;
    std::string guid = _clsidStr;

    hr = _pUnk->QueryInterface(
        IID_IDispatch, 
        reinterpret_cast<void**>(&pDisp));
    RETURN_ON_FAIL(hr);

    UINT count = 0;
    hr = pDisp->GetTypeInfoCount(&count);

    if (count == 0 || FAILED(hr)) {
        return hr;
    }

    return pDisp->GetTypeInfo(0, LOCALE_USER_DEFAULT, pTypeInfo);
}

void CoMiExtractor::_PrintParams(const std::vector<ParamInfo>& paramInfos) {
    for (size_t i = 0; i < paramInfos.size(); ++i) {
        if (paramInfos.size() > 1) {
            std::cout << "\t\t";
        }

        if (paramInfos[i].direction & PARAMFLAG_FIN) {
            std::cout << "/* [in] */ ";
        }
        if (paramInfos[i].direction & PARAMFLAG_FOUT) {
            std::cout << "/* [out] */ ";
        }

        std::cout << Utils::GetVarTypeName(paramInfos[i].paramType) << " "
            << paramInfos[i].name;

        if (i < paramInfos.size() - 1) {
            std::cout << ", ";
            if (paramInfos.size() > 1) {
                std::cout << "\n";
            }
        }
    }
}

void CoMiExtractor::_PrintInterfaceInfos() {
    for (const auto& interfaceInfo : _interfaceInfos) {
        std::string interfaceId = Utils::RemoveChars(interfaceInfo.first, "{}");

        std::cout << "MIDL_INTERFACE(\"" << interfaceId << "\")\n";
        std::cout << interfaceInfo.second.name << " : public IUnknown\n";
        std::cout << "{\n";
        std::cout << "public:\n";

        for (size_t i = 0; i < interfaceInfo.second.methodInfos.size(); ++i) {
            const auto& methodInfo = interfaceInfo.second.methodInfos[i];

            std::cout << "\t// Display Id: 0x"
                << std::hex
                << methodInfo.dispId
                << ", RVA Offset: 0x"
                << std::hex
                << methodInfo.rvaOffset
                << "\n";

            std::cout << "\tvirtual "
                << Utils::GetVarTypeName(methodInfo.retType) << " "
                << "STDMETHODCALLTYPE "
                << Utils::GetInvokeKindName(methodInfo.invokeKind)
                << methodInfo.name << "(";

            if (methodInfo.paramInfos.size() > 1) {
                std::cout << "\n";
            }

            _PrintParams(methodInfo.paramInfos);

            std::cout << ") = 0;\n";

            if (i < interfaceInfo.second.methodInfos.size() - 1) {
                std::cout << "\n";
            }
        }
        std::cout << "}\n\n";
    }
}