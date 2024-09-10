#include <windows.h>
#include <objbase.h>
#include <iostream>
#include <comdef.h>

extern "C" {
    HRESULT init_dcom() {
        return CoInitializeEx(NULL,COINIT_APARTMENTTHREADED);
    }
    void uninit_dcom() {
        CoUninitialize();
    }
    HRESULT create_object(const wchar_t* remote_machine,const wchar_t* clsid_str,const wchar_t* iid_str,void** out_interface) {
        // Define variables.
        CLSID clsid;
        IID iid;
        HRESULT hr;
        
        // Setup the variables.
        hr = CLSIDFromString(clsid_str,&clsid);
        if(FAILED(hr)){
            std::wcerr << L"Failed to convert CLSID from string: " << clsid_str << std::endl;
            return hr;
        }
        hr = IIDFromString(iid_str,&iid);
        if(FAILED(hr)){
            std::wcerr << L"Failed to convert IID from string: " << iid_str << std::endl;
            return hr;
        }

        // Define server info.
        COSERVERINFO server_info = {0};
        server_info.pwszName = (wchar_t*)remote_machine;
        MULTI_QI multi_qi = {0};
        multi_qi.pIID = &iid;

        hr = CoCreateInstanceEx(
            clsid,
            NULL,
            CLSCTX_REMOTE_SERVER,
            &server_info,
            1,
            &multi_qi
        );
        if(SUCCEEDED(hr)){
            *out_interface = multi_qi.pItf;
        } else {
            std::wcerr << L"Failed to create DCOM object on remote machine" << std::endl;
        }
        return hr;
    }
}

