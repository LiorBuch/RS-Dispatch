#![allow(dead_code)]

use anyhow::{anyhow, Context};
use std::env;
use windows::core::*;
use windows::Win32::System::Com::*;
use windows::Win32::System::Ole::*;

const LOCALE_USER_DEFAULT: u32 = 0x0400;
const LOCALE_SYSTEM_DEFAULT: u32 = 0x0800;

pub struct Variant(VARIANT);

impl From<bool> for Variant {
    fn from(value: bool) -> Self {
        Self(value.into())
    }
}

impl From<i32> for Variant {
    fn from(value: i32) -> Self {
        Self(value.into())
    }
}

impl From<i64> for Variant {
    fn from(value: i64) -> Self {
        Self(value.into())
    }
}

impl From<f64> for Variant {
    fn from(value: f64) -> Self {
        Self(value.into())
    }
}

impl From<&str> for Variant {
    fn from(value: &str) -> Self {
        Self(BSTR::from(value).into())
    }
}

impl From<&String> for Variant {
    fn from(value: &String) -> Self {
        Self(BSTR::from(value).into())
    }
}

impl Variant {
    pub fn bool(&self) -> anyhow::Result<bool> {
        Ok(bool::try_from(&self.0)?)
    }

    pub fn int(&self) -> anyhow::Result<i32> {
        Ok(i32::try_from(&self.0)?)
    }
    pub fn long(&self) -> anyhow::Result<i64> {
        Ok(i64::try_from(&self.0)?)
    }
    pub fn float(&self) -> anyhow::Result<f64> {
        Ok(f64::try_from(&self.0)?)
    }

    pub fn string(&self) -> anyhow::Result<String> {
        Ok(BSTR::try_from(&self.0)?.to_string())
    }

    pub fn idispatch(&self) -> anyhow::Result<IDispatchW> {
        Ok(IDispatchW(IDispatch::try_from(&self.0)?))
    }

    pub fn vt(&self) -> u16 {
        unsafe { self.0.as_raw().Anonymous.Anonymous.vt }
    }
}
/// The IDispatchW is a wrapper structure for the [`windows::Win32::System::Com::IDispatch`] interface.   
/// It will be used for **all** of the IDispatch related actions and should always be called.   
/// @0:[IDispatch] -> Pass the raw IDispatch.
pub struct IDispatchW(pub IDispatch);
impl IDispatchW {
    /// This is the general invoke method for the idispatch.   
    /// Instead of using this one you can use the [IDispatchW::get()] [IDispatchW::put()]   
    /// or use the relative data type for more convenient cast.   
    /// Use vec![] for empty args and .into() to convert it to Variant.
    pub fn invoke(
        &self,
        flags: DISPATCH_FLAGS,
        name: &str,
        mut args: Vec<Variant>,
    ) -> anyhow::Result<Variant> {
        unsafe {
            let mut dispatch_id = 0;
            self.0
                .GetIDsOfNames(
                    &GUID::default(),
                    &PCWSTR::from_raw(HSTRING::from(name).as_ptr()),
                    1,
                    LOCALE_USER_DEFAULT,
                    &mut dispatch_id,
                ).with_context(|| "GetIDsOfNames Failiure!")?;
            let mut dispatch_param = DISPPARAMS::default();
            let mut dispatch_named = DISPID_PROPERTYPUT;
            if !args.is_empty() {
                args.reverse();
                dispatch_param.cArgs = args.len() as u32;
                dispatch_param.rgvarg = args.as_mut_ptr() as *mut VARIANT;
                if (flags & DISPATCH_PROPERTYPUT) != DISPATCH_FLAGS(0) {
                    dispatch_param.cNamedArgs = 1;
                    dispatch_param.rgdispidNamedArgs = &mut dispatch_named;
                }
            }
            let mut result = VARIANT::default();
            self.0
                .Invoke(
                    dispatch_id,
                    &GUID::default(),
                    LOCALE_SYSTEM_DEFAULT,
                    flags,
                    &dispatch_param,
                    Some(&mut result),
                    None,
                    None,
                ).with_context(|| "Invokation Failure!")?;
            Ok(Variant(result))
        }
    }

    pub fn get(&self, name: &str,args:Vec<Variant>) -> anyhow::Result<Variant> {
        self.invoke(DISPATCH_PROPERTYGET, name, args)
    }

    pub fn int(&self, name: &str,args:Vec<Variant>) -> anyhow::Result<i32> {
        let result = self.get(name,args)?;
        result.int()
    }

    pub fn bool(&self, name: &str,args:Vec<Variant>) -> anyhow::Result<bool> {
        let result = self.get(name,args)?;
        result.bool()
    }
    pub fn float(&self, name: &str,args:Vec<Variant>) -> anyhow::Result<f64> {
        let result = self.get(name,args)?;
        result.float()
    }

    pub fn string(&self, name: &str,args:Vec<Variant>) -> anyhow::Result<String> {
        let result = self.get(name,args)?;
        result.string()
    }

    pub fn put(&self, name: &str, args: Vec<Variant>) -> anyhow::Result<Variant> {
        self.invoke(DISPATCH_PROPERTYPUT, name, args)
    }

    pub fn call(&self, name: &str, args: Vec<Variant>) -> anyhow::Result<Variant> {
        self.invoke(DISPATCH_METHOD, name, args)
    }
}

/// Automatic uninitilize for the com object i case the memory was droped.   
/// can also be done manually with [RSCom::close_api()]
pub struct DeferCoUninitialize;
impl Drop for DeferCoUninitialize {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}
/// This structure hold the pointer for the COM IDispatch interface.
///
/// @api:[IDispatchW] -> The main interface for the COM.
pub struct RSCom {
    pub api: IDispatchW,
}

impl RSCom {
    /// Initializing the structure [`crate::com_module::RSCom`] and append for each pointer its interface.   
    /// @com_name:[&str] -> The com application name you want to use, for example:"Excel.Application".
    pub fn init(com_name:&str) -> anyhow::Result<RSCom> {
        let mut args = env::args();
        let _ = args.next();
        unsafe {
            // API Data
            // Init the thread
            let res = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            if res.is_err() {
                return Err(anyhow!("error: {}", res.message()));
            }
            //let _com = DeferCoUninitialize;
            // Get CLSID of the com
            let clsid = CLSIDFromProgID(PCWSTR::from_raw(
                HSTRING::from(com_name).as_ptr(),
            ))
            .with_context(|| "wrong clsid of api")?;
            println!("printing api id {:?}", clsid);
            // Create the instance of the COM
            let _api_dispatch = CoCreateInstance(&clsid, None, CLSCTX_LOCAL_SERVER)
                .with_context(|| "CoCreateInstance of api")?;
            // Cast from IDispatch to the IDispatchWrapper
            let api_dispatch = IDispatchW(_api_dispatch);
            Ok(RSCom { api: api_dispatch })
        }
    }
    /// Method for Uninitialize the Com Object.
    pub fn close_api(&self) -> () {
        unsafe { CoUninitialize() }
    }
}
unsafe impl Send for RSCom {}
