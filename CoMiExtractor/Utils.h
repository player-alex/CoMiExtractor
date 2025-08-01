#pragma once

#include <Windows.h>
#include <string>
#include <comdef.h>
#include <Psapi.h>

struct BSTRArrayDeleter {
    size_t count;

    BSTRArrayDeleter(size_t c) : count(c) {}

    void operator()(BSTR* arr) const {
        if (arr) {
            for (size_t i = 0; i < count; ++i) {
                #pragma warning(suppress: 6001)
                if (arr[i]) {
                    SysFreeString(arr[i]);
                }
            }

            delete[] arr;
        }
    }
};

const std::map<VARTYPE, std::string> INVOKE_KIND_NAMES = {
    { INVOKEKIND::INVOKE_FUNC, "" },
    { INVOKEKIND::INVOKE_PROPERTYGET, "Get" },
    { INVOKEKIND::INVOKE_PROPERTYPUT, "Set" },
    { INVOKEKIND::INVOKE_PROPERTYPUTREF, "SetRef" },
};

const std::map<VARTYPE, std::string> VARTYPE_NAMES = {
    { VT_EMPTY, "void" },                  // Empty / no data
    { VT_NULL, "NULL" },                   // Null pointer or value
    { VT_I2, "short" },                    // 2-byte signed integer
    { VT_I4, "long" },                     // 4-byte signed integer
    { VT_R4, "float" },                    // 4-byte floating-point
    { VT_R8, "double" },                   // 8-byte floating-point
    { VT_CY, "CY" },                       // Currency (64-bit fixed-point)
    { VT_DATE, "DATE" },                   // Date and time (double)
    { VT_BSTR, "BSTR" },                   // Basic string (OLECHAR*)
    { VT_DISPATCH, "IDispatch*" },         // Pointer to an IDispatch interface
    { VT_ERROR, "SCODE" },                 // Scode (HRESULT)
    { VT_BOOL, "VARIANT_BOOL" },           // Boolean (true is -1, false is 0)
    { VT_VARIANT, "VARIANT" },             // A VARIANT itself
    { VT_UNKNOWN, "IUnknown*" },           // Pointer to an IUnknown interface
    { VT_DECIMAL, "DECIMAL" },             // Decimal (128-bit)
    { VT_I1, "char" },                     // 1-byte signed integer
    { VT_UI1, "BYTE" },                    // 1-byte unsigned integer
    { VT_UI2, "WORD" },                    // 2-byte unsigned integer
    { VT_UI4, "DWORD" },                   // 4-byte unsigned integer
    { VT_I8, "long long" },                // 8-byte signed integer (INT64)
    { VT_UI8, "unsigned long long" },      // 8-byte unsigned integer (UINT64)
    { VT_INT, "int" },                     // Default integer size for the platform
    { VT_UINT, "unsigned int" },           // Default unsigned integer size for the platform
    { VT_VOID, "void" },                   // No return value / void pointer (often mapped to HRESULT in actual methods)
    { VT_HRESULT, "HRESULT" },             // OLE status code
    { VT_PTR, "void*" },                   // Generic pointer (specific type depends on context)
    { VT_SAFEARRAY, "SAFEARRAY*" },        // Safe array of data
    { VT_CARRAY, "C-style array" },        // C-style array (less common in modern COM)
    { VT_USERDEFINED, "IUnknown*" },       // User-defined type (typically an interface pointer)
    { VT_LPSTR, "LPSTR" },                 // Long pointer to a null-terminated ANSI string
    { VT_LPWSTR, "LPWSTR" },               // Long pointer to a null-terminated Unicode string
    { VT_RECORD, "void*" },                // User-defined record (struct/class) - often handled via IRecordInfo
    { VT_INT_PTR, "INT_PTR" },             // Signed integer type for pointers (machine-specific size)
    { VT_UINT_PTR, "UINT_PTR" },           // Unsigned integer type for pointers (machine-specific size)
    { VT_FILETIME, "FILETIME" },           // 64-bit file time
    { VT_BLOB, "BYTE*" },                  // Binary Large Object (pointer to raw bytes)
    { VT_STREAM, "IStream*" },             // Pointer to an IStream interface
    { VT_STORAGE, "IStorage*" },           // Pointer to an IStorage interface
    { VT_STREAMED_OBJECT, "IStream*" },    // Object serialized into a stream
    { VT_STORED_OBJECT, "IStorage*" },     // Object stored in a storage
    { VT_BLOB_OBJECT, "IUnknown*" },       // Object stored as a BLOB
    { VT_CF, "CLIPFORMAT" },               // Clipboard format
    { VT_CLSID, "CLSID" },                 // Class ID (GUID)
    { VT_VERSIONED_STREAM, "IStream*" },   // Versioned stream
    { VT_BSTR_BLOB, "BYTE*" },             // Binary string BLOB
    { VT_VECTOR, "void*" },                // Vector (fixed-size array)
    { VT_ARRAY, "SAFEARRAY*" },            // Array (dynamic-size array, often SAFEARRAY)
    { VT_BYREF, "void*" },                 // By reference (pointer to data)
    { VT_RESERVED, "RESERVED" },           // Reserved for future use
    { VT_ILLEGAL, "ILLEGAL" },             // Illegal VARTYPE
    { VT_ILLEGALMASKED, "ILLEGALMASKED" }, // Illegal VARTYPE masked
    { VT_TYPEMASK, "TYPEMASK" }            // Mask for VARTYPE base type
};

namespace Utils {
    static std::string WideCharToMultiByteString(const std::wstring& s) {
        if (s.empty()) {
            return "";
        }

        int sizeNeeded = WideCharToMultiByte(CP_UTF8, 0, s.c_str(), -1, nullptr, 0, nullptr, nullptr);
        if (sizeNeeded <= 0) {
            return "";
        }

        std::string result(sizeNeeded - 1, '\0');

        WideCharToMultiByte(CP_UTF8, 0, s.c_str(), -1, &result[0], sizeNeeded, nullptr, nullptr);
        return result;
    }

    static std::wstring MultiByteToWideCharString(const std::string& s) {
        if (s.empty()) {
            return L"";
        }

        int sizeNeeded = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, nullptr, 0);
        if (sizeNeeded <= 0) {
            return L"";
        }

        std::wstring result(sizeNeeded - 1, L'\0');

        MultiByteToWideChar(CP_UTF8, 0, s.c_str(), -1, &result[0], sizeNeeded);
        return result;
    }

    static CLSID StringToCLSID(const char* s) {
        CLSID result = {};
        HRESULT hr = CLSIDFromString(MultiByteToWideCharString(s).c_str(), &result);

        if (FAILED(hr)) {
            throw std::invalid_argument("Invalid CLSID string.");
        }

        return result;
    }

    static std::string GetTypeName(ITypeInfo* pTypeInfo) {
        std::string result;
        CComBSTR nameBuf = nullptr;
            
        if (FAILED(pTypeInfo->GetDocumentation(
            MEMBERID_NIL, &nameBuf, 
            nullptr, nullptr, nullptr))) {
            return result;
        }

        result = WideCharToMultiByteString(std::wstring(nameBuf));
        return result;
    }

    static std::string GetVarTypeName(VARTYPE vt) {
        auto it = VARTYPE_NAMES.find(vt);
        return it == VARTYPE_NAMES.end() ? "?" : it->second;
    }

    static std::string GetInvokeKindName(INVOKEKIND iv) {
        auto it = INVOKE_KIND_NAMES.find(iv);
        return it == INVOKE_KIND_NAMES.end() ? "?" : it->second;
    }

    static std::string GetGuidStringFromTypeAttr(const TYPEATTR* pTypeAttr) {
        if (!pTypeAttr) {
            return "";
        }

        wchar_t szGuid[40] = {};
        int len = StringFromGUID2(pTypeAttr->guid, szGuid, _countof(szGuid));

        if (len == 0) {
            return "";
        }

        return WideCharToMultiByteString(std::wstring(szGuid));
    }

    static uintptr_t CalculateRVAFromAddress(LPVOID pAddress) {
        MEMORY_BASIC_INFORMATION mbi;

        if (!VirtualQuery(pAddress, &mbi, sizeof(mbi))) {
            return 0;
        }

        return reinterpret_cast<uintptr_t>(pAddress) - reinterpret_cast<uintptr_t>(mbi.AllocationBase);
    }
    
    static std::string RemoveChars(const std::string& s, std::string chars) {
        std::string result = s;

        for (const auto c : chars) {
            result.erase(
                std::remove(result.begin(), result.end(), c),
                result.end());
        }

        return result;
    }

    static std::string HResultToComErrorMessage(HRESULT hr) {
        try {
            _com_error err(hr);
            return WideCharToMultiByteString(err.ErrorMessage());
        }
        catch (...) {
            return "0x" + std::to_string(hr);
        }
    }
};