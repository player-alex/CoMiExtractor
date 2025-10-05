# 🔍 CoMiExtractor

> **A powerful Windows COM (Component Object Model) Interface Extractor for developers and reverse engineers**

[![Platform](https://img.shields.io/badge/platform-Windows-blue.svg)](https://www.microsoft.com/windows)
[![Language](https://img.shields.io/badge/language-C%2B%2B-orange.svg)](https://isocpp.org/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

## 📋 Overview

CoMiExtractor is a specialized tool designed to inspect and extract interface information from Windows COM objects. It leverages `IProvideClassInfo` and `IDispatch` interfaces to provide comprehensive details about COM interfaces, methods, and parameters - essential for development, debugging, and reverse engineering.

## ✨ Features

- 🎯 **Interface Extraction** - Automatically discovers and extracts COM interface definitions
- 🔬 **Method Analysis** - Retrieves detailed method information including parameters and return types
- 📊 **RVA Offset Calculation** - Computes Relative Virtual Address offsets for each method
- 🧬 **Type Information** - Provides complete type information for parameters and return values
- 🛠️ **Development Ready** - Outputs C++-compatible interface definitions ready for use

## 📦 What Gets Extracted

### Interface Information

| Target | Description |
|:------:|:-----------|
| **Name** | Interface name |
| **IID** | Interface identifier (GUID) |

### Method Information

| Target | Description |
|:------:|:-----------|
| **Name** | Method name |
| **Return Type** | Method return value type |
| **Invoke Kind** | Property access type [Get, Set, SetRef] |
| **Display Id** | Method display ID (DISPID) |
| **RVA Offset** | Relative Virtual Address offset |

### Parameter Information

| Target | Description |
|:------:|:-----------|
| **Name** | Parameter name |
| **Direction** | Parameter direction [In, Out] |
| **Type** | Parameter type ([VARENUM reference](https://learn.microsoft.com/en-us/windows/win32/api/wtypes/ne-wtypes-varenum)) |

## 🚀 Usage

### Option 1: Interactive Mode
```bash
CoMiExtractor.exe
# Enter CLSID when prompted
```

### Option 2: Command Line
```bash
CoMiExtractor.exe "{CLSID}"
```

### Example
```bash
CoMiExtractor.exe "{2A0B9D10-4B87-11D3-A97A-00104B365C9F}"
```

## 📄 Output Example

```cpp
MIDL_INTERFACE("2A0B9D10-4B87-11D3-A97A-00104B365C9F")
IFileSystem3 : public IUnknown
{
public:
        // Display Id: 0x60000000
        static const uintptr_t QueryInterfaceRvaOffset = 0x3c80;
        virtual void STDMETHODCALLTYPE QueryInterface(
                /* [in] */ void* riid,
                /* [out] */ void* ppvObj) = 0;

        // Display Id: 0x60000001
        static const uintptr_t AddRefRvaOffset = 0xfec0;
        virtual DWORD STDMETHODCALLTYPE AddRef() = 0;

        // Display Id: 0x60000002
        static const uintptr_t ReleaseRvaOffset = 0x5600;
        virtual DWORD STDMETHODCALLTYPE Release() = 0;

        // Display Id: 0x2710
        static const uintptr_t BuildPathRvaOffset = 0x14870;
        virtual BSTR STDMETHODCALLTYPE BuildPath(
                /* [in] */ BSTR Path,
                /* [in] */ BSTR Name) = 0;

        // Display Id: 0x460
        static const uintptr_t CreateFolderRvaOffset = 0x14bc0;
        virtual void* STDMETHODCALLTYPE CreateFolder(/* [in] */ BSTR Path) = 0;
};
```

## 🏗️ Building from Source

### Prerequisites
- Visual Studio 2022 (or later)
- Windows SDK 10.0
- Platform Toolset v143

### Build Steps

1. **Clone the repository**
   ```bash
   git clone https://github.com/player-alex/CoMiExtractor.git
   cd CoMiExtractor
   ```

2. **Open the solution**
   ```bash
   cd CoMiExtractor
   start CoMiExtractor.sln
   ```

3. **Build the project**
   - Select configuration (Debug/Release)
   - Select platform (Win32/x64)
   - Build → Build Solution (Ctrl+Shift+B)

4. **Run the executable**
   ```bash
   cd Release
   CoMiExtractor.exe
   ```

## 🔧 Project Structure

```
CoMiExtractor/
├── CoMiExtractor/
│   ├── inc/
│   │   ├── CoMiExtractor.h    # Main extractor class header
│   │   └── Utils.h             # Utility functions and type mappings
│   ├── src/
│   │   ├── main.cpp            # Entry point
│   │   └── CoMiExtractor.cpp   # Core extraction implementation
│   └── CoMiExtractor.vcxproj   # Visual Studio project file
├── LICENSE                      # MIT License
└── README.md                    # This file
```

## 🛡️ Use Cases

### Development
- 📚 Generate interface definitions for undocumented COM objects
- 🔄 Understand COM object capabilities and method signatures
- 🧪 Create test harnesses for COM components

### Reverse Engineering
- 🔍 Analyze proprietary COM components
- 🗺️ Map interface structures and method offsets
- 🔐 Understand system component internals

### Security Research
- 🛡️ Audit COM object attack surfaces
- 🔬 Identify exploitable method signatures
- 📊 Document interface capabilities for threat modeling

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## ⚠️ Disclaimer

This tool is intended for legitimate development, debugging, and security research purposes only. Users are responsible for ensuring their use of this tool complies with all applicable laws and regulations.

## 🙏 Acknowledgments

- Built with Windows COM/OLE APIs
- Uses ATL (Active Template Library) for COM support
- Type information extracted via ITypeInfo interfaces

---

<div align="center">

**Made with ❤️ for the Windows development community**

⭐ Star this repository if you find it useful!

</div>
