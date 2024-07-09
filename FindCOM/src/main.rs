extern crate winapi;

use std::ffi::OsStr;
use std::fs::File;
use std::io::{BufRead, BufReader};
use std::iter::once;
use std::mem::zeroed;
use std::os::windows::ffi::OsStrExt;
use std::ptr::null_mut;
use winapi::ctypes::c_void;
use winapi::shared::{
    guiddef::GUID,
    wtypesbase::CLSCTX_INPROC_SERVER,
};
use winapi::um::{
    combaseapi::{CoCreateInstance, CoInitializeEx, CoUninitialize, CLSIDFromString},
    oaidl::{IDispatch, DISPPARAMS, VARIANT},
    objbase::COINIT_APARTMENTTHREADED,
    oleauto::{VariantClear, DISPATCH_METHOD, SysAllocString},
    winnt::HRESULT,
};
use winapi::shared::wtypes::BSTR;
use winapi::Interface;

fn main() {
    unsafe {

        let hr: HRESULT = CoInitializeEx(null_mut(), COINIT_APARTMENTTHREADED);
        if hr != winapi::shared::winerror::S_OK && hr != winapi::shared::winerror::S_FALSE {
            eprintln!("CoInitializeEx failed: HRESULT = 0x{:X}", hr);
            return;
        }


        let file = match File::open("clsids.txt") {
            Ok(file) => file,
            Err(e) => {
                eprintln!("Error opening clsids.txt: {:?}", e);
                CoUninitialize();
                return;
            }
        };
        let reader = BufReader::new(file);
        let clsid_list: Vec<String> = reader.lines().filter_map(Result::ok).collect();

        for clsid_str in clsid_list {

            let clsid_wide: Vec<u16> = OsStr::new(&clsid_str)
                .encode_wide()
                .chain(once(0))
                .collect();
            let mut clsid: GUID = zeroed();
            let hr: HRESULT = CLSIDFromString(clsid_wide.as_ptr(), &mut clsid);

            if hr != winapi::shared::winerror::S_OK {
                eprintln!("CLSIDFromString failed for {}: HRESULT = 0x{:X}", clsid_str, hr);
                continue;
            }

            println!("Attempting to create COM instance for CLSID: {}", clsid_str);


            let mut unknown: *mut c_void = null_mut();
            let hr = CoCreateInstance(
                &clsid,
                null_mut(),
                CLSCTX_INPROC_SERVER,
                &IDispatch::uuidof(),
                &mut unknown,
            );

            if hr != winapi::shared::winerror::S_OK {
                eprintln!("CoCreateInstance failed for {}: HRESULT = 0x{:X}", clsid_str, hr);
                continue;
            }

            let dispatch: *mut IDispatch = unknown as *mut IDispatch;

            // 获取 DispID
            let mut method_name: Vec<u16> = OsStr::new("CreateTextFile").encode_wide().chain(once(0)).collect();
            let mut dispid: i32 = 0;
            let hr = (*dispatch).GetIDsOfNames(
                &mut zeroed::<GUID>(),
                &mut method_name.as_mut_ptr(),
                1,
                0,
                &mut dispid,
            );

            if hr != winapi::shared::winerror::S_OK {
                eprintln!("GetIDsOfNames failed for {}: HRESULT = 0x{:X}", clsid_str, hr);
                VariantClear(&mut zeroed::<VARIANT>());
                continue;
            }

            println!("DispID for CreateTextFile in CLSID {}: {}", clsid_str, dispid);


            let filename: BSTR = SysAllocString(OsStr::new("text.txt").encode_wide().chain(once(0)).collect::<Vec<u16>>().as_ptr());
            let mut filename_variant: VARIANT = zeroed();
            filename_variant.n1.n2_mut().vt = winapi::shared::wtypes::VT_BSTR as u16;
            *filename_variant.n1.n2_mut().n3.bstrVal_mut() = filename;

            let mut overwrite_variant: VARIANT = zeroed();
            overwrite_variant.n1.n2_mut().vt = winapi::shared::wtypes::VT_BOOL as u16;
            *overwrite_variant.n1.n2_mut().n3.boolVal_mut() = winapi::shared::wtypes::VARIANT_TRUE;

            let mut args = [overwrite_variant, filename_variant];
            let mut disp_params = DISPPARAMS {
                cArgs: args.len() as u32,
                rgvarg: args.as_mut_ptr(),
                rgdispidNamedArgs: null_mut(),
                cNamedArgs: 0,
            };
            let mut result = zeroed::<VARIANT>();


            let hr = (*dispatch).Invoke(
                dispid,
                &mut zeroed::<GUID>(),
                0,
                DISPATCH_METHOD,
                &mut disp_params,
                &mut result,
                null_mut(),
                null_mut(),
            );

            if hr == winapi::shared::winerror::S_OK {
                println!("COM Object {} Call-Success", clsid_str);
            } else {
                eprintln!("COM Object {} Call-Fail: HRESULT = 0x{:X}", clsid_str, hr);
            }


            VariantClear(&mut result);
            VariantClear(&mut args[0]);
            VariantClear(&mut args[1]);
        }

        CoUninitialize();
    }
}
