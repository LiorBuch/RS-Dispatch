#include <windows.h>
#include <objbase.h>
#include <iostream>
#include <comdef.h>

extern "C" {
    HRESULT init_dcom();
    void uninit_dcom();
    HRESULT create_object(
        const wchar_t* remote_machine,
        const wchar_t* clsid_string,
        const wchar_t* iid_string,
        void** out_interface
    );
}

int main() {
    HRESULT hr = init_dcom();
        if (FAILED(hr)) {
        std::cerr << "Failed to initialize COM" << std::endl;
        return 1;
    }
    // Remote machine name and COM object details
    const wchar_t* remote_machine = L"name";  //machine's name
    const wchar_t* clsid_string = L"{00024500-0000-0000-C000-000000000046}";  //CLSID
    const wchar_t* iid_string = L"{00020400-0000-0000-C000-000000000046}";    //IID

    void* pInterface = nullptr;
    hr = create_object(remote_machine,clsid_string,iid_string,&pInterface);
        if (FAILED(hr)) {
        std::cerr << "Failed to create DCOM object. HRESULT = " << std::hex << hr << std::endl;
        uninit_dcom();
        return 1;
    }
        std::cout << "Successfully created DCOM object." << std::endl;

    // Clean up: Release the interface and uninitialize COM
    if (pInterface != nullptr) {
        IUnknown* pUnknown = static_cast<IUnknown*>(pInterface);
        pUnknown->Release();
    }

    uninit_dcom();
    return 0;
}