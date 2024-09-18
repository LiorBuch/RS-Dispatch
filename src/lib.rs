pub mod idispatchw;
pub mod com_errors;
#[cfg(test)]
mod tests {
    use std::{thread, time::Duration};

    use crate::idispatchw;

    #[test]
    fn it_works_adv() {
        let com = idispatchw::RSComMap::new("InstrumentV3.InstrumentAPI");
        match com {
            Ok(obj) => {
                println!("Ok on Com!");
                let vis_r = obj.api.get("isRemote");
                match vis_r {
                    Ok(o) => {
                        thread::sleep(Duration::from_secs(1));
                        println!("Worked");
                        assert_eq!(1, 1);
                    },
                    Err(e) => {
                        println!("{}", e);
                        assert_eq!(1, -2);
                    },
                }
            },
            Err(e) => {
                println!("{}", e);
                assert_eq!(1, -1);
            },
        }
    }
    
    #[test]
    fn it_works() {
        let com = idispatchw::RSComMap::new("Excel.Application");
        match com {
            Ok(_) => {
                println!("Ok on Com!");
                assert_eq!(1, 1);
            },
            Err(e) => {
                println!("{}", e);
                assert_eq!(1, -1);
            },
        }
    }
}
