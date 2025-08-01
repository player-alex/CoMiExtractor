#pragma once

#include <wtypes.h>
#include <atlbase.h>
#include <string>
#include <vector>
#include <map>

typedef struct ParamInfo {
    std::string name;
    unsigned short direction;
    VARTYPE paramType;
} ParamInfo_s;

typedef struct MethodInfo {
    std::string name;
    uintptr_t vTableOffset;
    uintptr_t rvaOffset;
    INVOKEKIND invokeKind;
    DISPID dispId;
    VARTYPE retType;
    std::vector<ParamInfo> paramInfos;
} MethodInfo_s;

typedef struct InterfaceInfo {
    std::string name;
    std::vector<MethodInfo> methodInfos;
} InterfaceInfo_s;

class CoMiExtractor {
public:
    CoMiExtractor(std::string clsid);
    ~CoMiExtractor();

    HRESULT Extract();

private:
    std::string _clsidStr;
    CLSID _clsid;
    CComPtr<IUnknown> _pUnk = nullptr;
    std::map<std::string, InterfaceInfo> _interfaceInfos;

    HRESULT _QueryToIProvideClassInfo(ITypeInfo** pTypeInfo);
    HRESULT _QueryToIDispatch(ITypeInfo** pTypeInfo);

    void _PrintParams(const std::vector<ParamInfo>& paramInfos);
    void _PrintDisplayId(const MethodInfo methodInfo);
    void _PrintRvaOffset(const MethodInfo methodInfo);
    void _PrintInterfaceInfos();

    void _GetMethodOffsets(const char* iidStr);
    HRESULT _InspectCoClass(ITypeInfo* pTypeInfo, TYPEATTR* pTypeAttr);
    HRESULT _InspectTypeInfo(ITypeInfo* pTypeInfo, std::string& guid);
    HRESULT _ExtractMethodInfo(ITypeInfo* pTypeInfo, UINT index, MethodInfo* pMethodInfo);
    void _ExtractParams(FUNCDESC* pFuncDesc, BSTR* argNames, UINT nameCount, std::vector<ParamInfo>& paramInfos);
};