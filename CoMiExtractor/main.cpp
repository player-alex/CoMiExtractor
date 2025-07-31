#include <Windows.h>
#include <atlcomcli.h>
#include <iostream>
#include "CoMiExtractor.h"
#include "Utils.h"

static void init() {
    HRESULT hr;

    hr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

    if (FAILED(hr)) {
        throw std::runtime_error(
            "CoInitializationEx failed: " + 
            Utils::HResultToComErrorMessage(hr));
    }
}

int main(int argc, char** argv) {
    HRESULT hr;
    std::string clsidStr;

    SetConsoleOutputCP(CP_UTF8);

    try {
        init();

        if (argc == 1) {
            std::cout << "Enter CLSID: ";
            std::getline(std::cin, clsidStr);
        }
        else {
            clsidStr = argv[argc - 1];
        }
        
        CoMiExtractor extractor(clsidStr);
        hr = extractor.Extract();

        if (FAILED(hr)) {
            throw hr;
        }
    }
    catch (HRESULT hr) {
        std::cout << Utils::HResultToComErrorMessage(hr) << std::endl;
    }
    catch (const std::exception& e) {
        std::cout << e.what() << std::endl;
    }
    catch (...) {
        std::cout << "Unknown exception caught" << std::endl;
    }

    CoUninitialize();
    return 0;
}