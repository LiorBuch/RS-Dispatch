pub mod com;
pub mod common;
#[cfg(test)]
mod tests {
    use std::{thread, time::Duration};

    use crate::com::com_module::RSCom;



    #[test]
    fn it_works() {
        let com = RSCom::init("InternetExplorer.Application");
        match com {
            Ok(obj) => {
                println!("Ok on Com!");
                let vis_r = obj.api.get("Visible",vec![]);
                match vis_r {
                    Ok(_o) => {
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
}
