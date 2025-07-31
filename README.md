# CoMiExtractor
Windows COM(Component Object Model) Interfaces Extractor  
The tool will be inspect *IProviderClassInfo*, *IDispatch* and extract interface informations.

Extracted informations can be used to development, reverse engineering.

## Extraction Table
**Interface**

|Target|Description|
|:-:|:-:|
|Name|Interface name|
|IID|Interface id|

**Method**

|Target|Description|
|:-:|:-:|
|Name|Method name|
|Return Type|Method return value type|
|Invoke Kind| [Get, Set, SetRef]|
|Display Id|Method display id. You can guess information using this id.|
|RVA Offset|RVA Offset. Calculated by *VTable[MethodOffset] - ModuleBase*|

**Parameter**

|Target|Description|
|:-:|:-:|
|Name|Parameter name|
|Direction|Parameter direction. [ In, Out ]|
|Type| Parameter type. [Ref](https://learn.microsoft.com/en-us/windows/win32/api/wtypes/ne-wtypes-varenum)|

## Usage
Two options provided.

1. Execute program and enter clsid manually
2. CoMiExtractor.exe "{CLSID}"

## Example Output
```bash
Enter CLSID: {0D43FE01-F093-11CF-8940-00A0C9054228}
MIDL_INTERFACE("2A0B9D10-4B87-11D3-A97A-00104B365C9F")
IFileSystem3 : public IUnknown
{
public:
        // Display Id: 0x60000000, RVA Offset: 0x3c80
        virtual void STDMETHODCALLTYPE QueryInterface(
                /* [in] */ void* riid,
                /* [out] */ void* ppvObj) = 0;

        // Display Id: 0x60000001, RVA Offset: 0xfec0
        virtual DWORD STDMETHODCALLTYPE AddRef() = 0;

        // Display Id: 0x60000002, RVA Offset: 0x5600
        virtual DWORD STDMETHODCALLTYPE Release() = 0;

        // Display Id: 0x60010000, RVA Offset: 0x5a60
        virtual void STDMETHODCALLTYPE GetTypeInfoCount(/* [out] */ void* pctinfo) = 0;

        // Display Id: 0x60010001, RVA Offset: 0x5790
        virtual void STDMETHODCALLTYPE GetTypeInfo(
                /* [in] */ unsigned int itinfo,
                /* [in] */ DWORD lcid,
                /* [out] */ void* pptinfo) = 0;

        // Display Id: 0x60010002, RVA Offset: 0x58e0
        virtual void STDMETHODCALLTYPE GetIDsOfNames(
                /* [in] */ void* riid,
                /* [in] */ void* rgszNames,
                /* [in] */ unsigned int cNames,
                /* [in] */ DWORD lcid,
                /* [out] */ void* rgdispid) = 0;

        // Display Id: 0x60010003, RVA Offset: 0x57f0
        virtual void STDMETHODCALLTYPE Invoke(
                /* [in] */ long dispidMember,
                /* [in] */ void* riid,
                /* [in] */ DWORD lcid,
                /* [in] */ WORD wFlags,
                /* [in] */ void* pdispparams,
                /* [out] */ void* pvarResult,
                /* [out] */ void* pexcepinfo,
                /* [out] */ void* puArgErr) = 0;

        // Display Id: 0x271a, RVA Offset: 0x15d60
        virtual void* STDMETHODCALLTYPE GetDrives() = 0;

        // Display Id: 0x2710, RVA Offset: 0x14870
        virtual BSTR STDMETHODCALLTYPE BuildPath(
                /* [in] */ BSTR Path,
                /* [in] */ BSTR Name) = 0;

        // Display Id: 0x2714, RVA Offset: 0x154a0
        virtual BSTR STDMETHODCALLTYPE GetDriveName(/* [in] */ BSTR Path) = 0;

        // Display Id: 0x2715, RVA Offset: 0x15780
        virtual BSTR STDMETHODCALLTYPE GetParentFolderName(/* [in] */ BSTR Path) = 0;

        // Display Id: 0x2716, RVA Offset: 0x15540
        virtual BSTR STDMETHODCALLTYPE GetFileName(/* [in] */ BSTR Path) = 0;

        // Display Id: 0x2717, RVA Offset: 0x153c0
        virtual BSTR STDMETHODCALLTYPE GetBaseName(/* [in] */ BSTR Path) = 0;

        // Display Id: 0x2718, RVA Offset: 0x154f0
        virtual BSTR STDMETHODCALLTYPE GetExtensionName(/* [in] */ BSTR Path) = 0;

        // Display Id: 0x2712, RVA Offset: 0x15200
        virtual BSTR STDMETHODCALLTYPE GetAbsolutePathName(/* [in] */ BSTR Path) = 0;

        // Display Id: 0x2713, RVA Offset: 0x15920
        virtual BSTR STDMETHODCALLTYPE GetTempName() = 0;

        // Display Id: 0x271f, RVA Offset: 0x14eb0
        virtual VARIANT_BOOL STDMETHODCALLTYPE DriveExists(/* [in] */ BSTR DriveSpec) = 0;

        // Display Id: 0x2720, RVA Offset: 0x15040
        virtual VARIANT_BOOL STDMETHODCALLTYPE FileExists(/* [in] */ BSTR FileSpec) = 0;

        // Display Id: 0x2721, RVA Offset: 0x15120
        virtual VARIANT_BOOL STDMETHODCALLTYPE FolderExists(/* [in] */ BSTR FolderSpec) = 0;

        // Display Id: 0x271b, RVA Offset: 0x15460
        virtual void* STDMETHODCALLTYPE GetDrive(/* [in] */ BSTR DriveSpec) = 0;

        // Display Id: 0x271c, RVA Offset: 0x3670
        virtual void* STDMETHODCALLTYPE GetFile(/* [in] */ BSTR FilePath) = 0;

        // Display Id: 0x271d, RVA Offset: 0x15740
        virtual void* STDMETHODCALLTYPE GetFolder(/* [in] */ BSTR FolderPath) = 0;

        // Display Id: 0x271e, RVA Offset: 0x157d0
        virtual void* STDMETHODCALLTYPE GetSpecialFolder(/* [in] */ IUnknown* SpecialFolder) = 0;

        // Display Id: 0x4b0, RVA Offset: 0x14d50
        virtual void STDMETHODCALLTYPE DeleteFile(
                /* [in] */ BSTR FileSpec,
                /* [in] */ VARIANT_BOOL Force) = 0;

        // Display Id: 0x4b1, RVA Offset: 0x14e00
        virtual void STDMETHODCALLTYPE DeleteFolder(
                /* [in] */ BSTR FolderSpec,
                /* [in] */ VARIANT_BOOL Force) = 0;

        // Display Id: 0x4b4, RVA Offset: 0x15a40
        virtual void STDMETHODCALLTYPE MoveFile(
                /* [in] */ BSTR Source,
                /* [in] */ BSTR Destination) = 0;

        // Display Id: 0x4b5, RVA Offset: 0x15b60
        virtual void STDMETHODCALLTYPE MoveFolder(
                /* [in] */ BSTR Source,
                /* [in] */ BSTR Destination) = 0;

        // Display Id: 0x4b2, RVA Offset: 0x14940
        virtual void STDMETHODCALLTYPE CopyFile(
                /* [in] */ BSTR Source,
                /* [in] */ BSTR Destination,
                /* [in] */ VARIANT_BOOL OverWriteFiles) = 0;

        // Display Id: 0x4b3, RVA Offset: 0x14a80
        virtual void STDMETHODCALLTYPE CopyFolder(
                /* [in] */ BSTR Source,
                /* [in] */ BSTR Destination,
                /* [in] */ VARIANT_BOOL OverWriteFiles) = 0;

        // Display Id: 0x460, RVA Offset: 0x14bc0
        virtual void* STDMETHODCALLTYPE CreateFolder(/* [in] */ BSTR Path) = 0;

        // Display Id: 0x44d, RVA Offset: 0x14c80
        virtual void* STDMETHODCALLTYPE CreateTextFile(
                /* [in] */ BSTR FileName,
                /* [in] */ VARIANT_BOOL Overwrite,
                /* [in] */ VARIANT_BOOL Unicode) = 0;

        // Display Id: 0x44c, RVA Offset: 0x15c90
        virtual void* STDMETHODCALLTYPE OpenTextFile(
                /* [in] */ BSTR FileName,
                /* [in] */ IUnknown* IOMode,
                /* [in] */ VARIANT_BOOL Create,
                /* [in] */ IUnknown* Format) = 0;

        // Display Id: 0x4e20, RVA Offset: 0x158c0
        virtual void* STDMETHODCALLTYPE GetStandardStream(
                /* [in] */ IUnknown* StandardStreamType,
                /* [in] */ VARIANT_BOOL Unicode) = 0;

        // Display Id: 0x4e2a, RVA Offset: 0x15590
        virtual BSTR STDMETHODCALLTYPE GetFileVersion(/* [in] */ BSTR FileName) = 0;
}
...
```