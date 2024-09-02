use anyhow::Context;
use std::collections::HashMap;
use std::env;
use std::result::Result;
use std::thread;
use std::time::Duration;
use windows::core::*;
use windows::Win32::System::Com::*;
use windows::Win32::System::Ole::*;

use crate::com_errors::ComError;

const LOCALE_USER_DEFAULT: u32 = 0x0400;
const LOCALE_SYSTEM_DEFAULT: u32 = 0x0800;

/// A Wrapper Struct for VARIANT
pub struct RSVariant(VARIANT);

impl From<bool> for RSVariant {
    fn from(value: bool) -> Self {
        Self(value.into())
    }
}

impl From<i32> for RSVariant {
    fn from(value: i32) -> Self {
        Self(value.into())
    }
}

impl From<i64> for RSVariant {
    fn from(value: i64) -> Self {
        Self(value.into())
    }
}

impl From<f64> for RSVariant {
    fn from(value: f64) -> Self {
        Self(value.into())
    }
}

impl From<&str> for RSVariant {
    fn from(value: &str) -> Self {
        Self(BSTR::from(value).into())
    }
}

impl From<&String> for RSVariant {
    fn from(value: &String) -> Self {
        Self(BSTR::from(value).into())
    }
}

impl RSVariant {
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
///
/// It will be used for **all** IDispatch related actions and should always be called.
pub struct IDispatchW(pub IDispatch);
impl IDispatchW {
    /// This is the general method for the invoke, it will call the `name` of the function with the `args` of the function.
    fn invoke(
        &self,
        flags: DISPATCH_FLAGS,
        name: &str,
        mut args: Vec<RSVariant>,
    ) -> anyhow::Result<RSVariant> {
        unsafe {
            let mut dispatch_id = 0;
            self.0
                .GetIDsOfNames(
                    &GUID::default(),
                    &PCWSTR::from_raw(HSTRING::from(name).as_ptr()),
                    1,
                    LOCALE_USER_DEFAULT,
                    &mut dispatch_id,
                )
                .with_context(|| "Failed to get IDs of names")?;
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
                )
                .with_context(|| "Failed to invoke")?;
            Ok(RSVariant(result))
        }
    }
    pub fn get(&self, name: &str) -> anyhow::Result<RSVariant> {
        self.invoke(DISPATCH_PROPERTYGET, name, vec![])
    }
    pub fn get_args(&self, name: &str,args:Vec<RSVariant>) -> anyhow::Result<RSVariant> {
        self.invoke(DISPATCH_PROPERTYGET, name, args)
    }

    pub fn get_idispatch(&self, name: &str) -> anyhow::Result<IDispatchW> {
        let result = self.get(name)?;
        result.idispatch()
    }

    pub fn get_idispatch_args(&self, name: &str,args:Vec<RSVariant>) -> anyhow::Result<IDispatchW> {
        let result = self.get_args(name,args)?;
        result.idispatch()
    }

    pub fn int(&self, name: &str) -> anyhow::Result<i32> {
        let result = self.get(name)?;
        result.int()
    }

    pub fn bool(&self, name: &str) -> anyhow::Result<bool> {
        let result = self.get(name)?;
        result.bool()
    }
    pub fn float(&self, name: &str) -> anyhow::Result<f64> {
        let result = self.get(name)?;
        result.float()
    }

    pub fn string(&self, name: &str) -> anyhow::Result<String> {
        let result = self.get(name)?;
        result.string()
    }

    pub fn put(&self, name: &str, args: Vec<RSVariant>) -> anyhow::Result<RSVariant> {
        self.invoke(DISPATCH_PROPERTYPUT, name, args)
    }

    pub fn call(&self, name: &str, args: Vec<RSVariant>) -> anyhow::Result<RSVariant> {
        self.invoke(DISPATCH_METHOD, name, args)
    }
}

pub struct DeferCoUninitialize;
impl Drop for DeferCoUninitialize {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}
/// This structure hold the pointers for each key IDispatch interface
///
/// api:`IDIspatchW` -> The main interface for the Instrument COM.
/// api_map:`HashMap` -> Can save [`IDispatch`] pointers for later usage, not a must.
pub struct RSCom {
    pub api: IDispatchW,
    pub api_map: HashMap<String, IDispatchW>,
}

impl RSCom {
    /// Initializing the structure [`crate::com_module::InstrumentCom`] and append for each pointer its interface.
    pub fn new(prog_id: &str) -> std::result::Result<RSCom, ComError> {
        let mut args = env::args();
        let _ = args.next();
        unsafe {
            // API Data
            // Init the thread
            let res = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            if res.is_err() {
                return Err(ComError::NotInitialize());
            }
            let _com = DeferCoUninitialize;
            // Get CLSID of the com
            
            let clsid = CLSIDFromProgID(PCWSTR::from_raw(HSTRING::from(prog_id).as_ptr()))
                .map_err(|_| ComError::ComNotFound())?;
            println!("printing api id {:?}", clsid);
            // Create the instance of the COM
            let _api_dispatch = CoCreateInstance(&clsid, None, CLSCTX_INPROC_SERVER)
                .map_err(|e| ComError::ComInstance(e.message()))?;
            thread::sleep(Duration::from_millis(1000));
            // Cast from IDispatch to the IDispatchWrapper
            let api_dispatch = IDispatchW(_api_dispatch);
            // Getter in the api for the Display IDispatch
            let com = RSCom {
                api: api_dispatch,
                api_map: HashMap::new(),
            };
            return Ok(com);
        }
    }
    pub fn save_idispatch(&mut self, key: String, idispatch: IDispatchW) -> Result<(), ComError> {
        let sucess = self.api_map.insert(key, idispatch);
        if sucess.is_some() {
            return Err(ComError::PointerAlreadyMapped());
        }
        Ok(())
    }
    pub fn get_map_idispatch(&mut self, key: String) -> Result<&IDispatchW, ComError>{
        let sucess = self.api_map.get(&key);
        match sucess {
            Some(idis) => Ok(idis),
            None => Err(ComError::NoInterface("Interface is not mapped!".to_string(),))
        }
        
    }
    pub fn remove_idispatch(&mut self, key: String) -> Result<(), ComError> {
        let sucess = self.api_map.remove(&key);
        if sucess.is_none() {
            return Err(ComError::NoInterface(
                "Interface is not mapped!".to_string(),
            ));
        }
        Ok(())
    }
}
impl Drop for RSCom {
    fn drop(&mut self) {
        unsafe { CoUninitialize()}
    }
}
