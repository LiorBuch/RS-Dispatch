RS Dispatch
======
[![Com CI](https://github.com/LiorBuch/RS-Dispatch/actions/workflows/com.yml/badge.svg)](https://github.com/LiorBuch/RS-Dispatch/actions/workflows/com.yml)
======

Crate to provide a simplified method to use windows COM objects with IDispatch interface.

## Example

```rust
    let com = RSCom::init("InternetExplorer.Application")?;
    let vis_r = com.api.get("Visible",vec![]);
    com.put("Visible",vec![true.into()]);
```


## Change Log 0.1.0

- Initial version.
- Basic functionality for IDispatch.