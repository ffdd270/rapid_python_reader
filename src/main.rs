extern crate core;

mod sheet;

use std::env;
use std::fs::File;
use std::io::{Write};
use std::time::{SystemTime};
use crate::sheet::sheet::{get_sheet, get_rows_by_id};


fn main()
{
    let mut args_iter = env::args().enumerate();

    if env::args().count() != 4
    {
        println!("use this usage : rapid_python_reader <xlsx_path> <output_path> <target_id>");
        return;
    };

    let begin_sys_time : SystemTime = SystemTime::now();
    // 실행 파일

    args_iter.next();
    // 여기서 as_str하면 임시값의 생명 주기가 박살나서 안 됨.
    let path = args_iter.next().unwrap().1;
    let output_path = args_iter.next().unwrap().1;
    let find_id = args_iter.next().unwrap().1;

    let find_ch00 = SystemTime::now();
    let result = get_rows_by_id(path.as_str(), 0, 1, find_id.as_str());
    let find_ch00_elapsed = match find_ch00.elapsed() {
        Err(e) => panic!("{}", e),
        Ok(t) => t,
    };

    match result {
        Some(result) => {
            println!("Found ID {}", find_id);
            //println!("{}", result);

            let mut output =  File::create(output_path).unwrap();
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