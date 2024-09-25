RS Dispatch
======

Crate to provide a simplified method to use windows COM objects with IDispatch interface.

## Example

```rust

let sdm_result:SafeDeviceMap = SafeDeviceMap::init(None);
match sdm_result {
    Ok(mapper) => {
        mapper.connect_device("address_01".to_string());
        let data = mapper.query_from_device("name_01".to_string(),"cool funcation with args").unwrap();
        println!("got {} from the device",data);
        mapper.disconnect_device("name_01".to_string());
    }
    Err(e) => {/*print codes or anything */}
}
```


## Change Log 0.1.0

- Initial version.
- Basic functionality for IDispatch.