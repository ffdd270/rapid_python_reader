extern crate core;

mod sheet;

use std::env;
use std::fs::File;
use std::io::{Write};
use std::time::{SystemTime};
use crate::sheet::sheet::{get_sheet, get_rows_by_id};
use std::time::Instant;

fn main()
{
    let mut args_iter = env::args().enumerate();

    if env::args().count() != 4
    {
        println!("use this usage : rapid_python_reader <xlsx_path> <output_path> <target_id>");
        return;
    };

    // 실행 파일

    args_iter.next();
    // 여기서 as_str하면 임시값의 생명 주기가 박살나서 안 됨.

    let path = args_iter.next().unwrap().1;
    let output_path = args_iter.next().unwrap().1;
    let find_id = args_iter.next().unwrap().1;

    let instant = Instant::now();

    let result_option : Option<String>;

    if find_id != "sheet_export" {
        result_option = get_rows_by_id(path.as_str(), 0, 1, find_id.as_str());
    }
    else {
        result_option = get_sheet(path.as_str(), 0, 1 );
    }

    match result_option {
        Some(result) => {
            println!("Found ID {}", find_id);
            let mut output =  File::create(output_path).unwrap();
            write!(output, "{}", result.as_str()).expect("FILE CANT WRITE.");
        },
        None => println!("Not found ID {}", find_id)
    };

    println!("find id total time : {:.2?}", instant.elapsed());

}