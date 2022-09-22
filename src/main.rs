extern crate core;

mod sheet;

use std::fs::File;
use std::io::{Error, Write};
use std::time::{Duration, SystemTime};
use crate::sheet::sheet::{get_sheet, get_rows_by_id};


fn main()
{

    let begin_sys_time : SystemTime = SystemTime::now();
    // opens a new workbook
    let path = "test.xlsx";

    get_sheet(path, 0,1);

    let elapsed = match begin_sys_time.elapsed() {
        Err(e) => panic!("{}", e),
        Ok(t) => t,
    };

    let current_elapsed = match SystemTime::now().elapsed() {
        Err(e) => panic!("{}", e),
        Ok(t) => t
    };

    println!("total time : {}", elapsed.as_millis() - current_elapsed.as_millis());

    let find_ch00 = SystemTime::now();
    let find_id = "CH01_001_1";
    let result = get_rows_by_id(path, 0, 1, find_id);
    let find_ch00_elapsed = match find_ch00.elapsed() {
        Err(e) => panic!("{}", e),
        Ok(t) => t,
    };

    match result {
        Some(result) => {
            println!("Found ID {}", find_id);
            println!("{}", result);

            let path = "sample.json";
            let mut output =  File::create(path).unwrap();
            write!(output, "{}", result.as_str()).expect("FILE CANT WRITE.");
        },
        None => println!("Not found ID {}", find_id)
    };

    let find_ch00_current_elapsed = match SystemTime::now().elapsed() {
        Err(e) => panic!("{}", e),
        Ok(t) => t
    };


    println!("find ch00_000 total time : {}", find_ch00_elapsed.as_millis() - find_ch00_current_elapsed.as_millis());

}