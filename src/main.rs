extern crate core;
extern crate regex;

mod sheet;

use std::env;
use std::fs::File;
use std::io::{Write};
use std::time::{SystemTime};
use std::env::Args;
use crate::sheet::sheet::{get_sheet, get_rows_by_id};
use std::time::Instant;
use std::fs;
use std::io;
use std::io::{BufReader, Error};
use std::path::Path;
use regex::Regex;
use calamine::{Reader, open_workbook, Xlsx, DataType, Range};
use std::thread;
use std::sync::mpsc;

fn export_sheet( path : &str, mut args_iter : &mut std::iter::Enumerate<std::env::Args> )
{
    let output_path = args_iter.next().unwrap().1;
    let find_id = args_iter.next().unwrap().1;


    let result_option : Option<String>;

    if find_id != "sheet_export" {
        result_option = get_rows_by_id(path, 0, 1, find_id.as_str());
    }
    else {
        result_option = get_sheet(path, 0, 1 );
    }

    match result_option {
        Some(result) => {
            println!("Found ID {}", find_id);
            let mut output =  File::create(output_path).unwrap();
            write!(output, "{}", result.as_str()).expect("FILE CANT WRITE.");
        },
        None => println!("Not found ID {}", find_id)
    };

}

fn list_excel_files_in_directory(path: &str) -> io::Result<Vec<String>> {
    let mut excel_files = Vec::new();
    let entries = fs::read_dir(path)?;

    for entry in entries {
        let entry = entry?;
        let path = entry.path();
        if path.is_file() {
            if let Some(filename) = path.file_name() {
                if filename.to_string_lossy().starts_with("~") {
                    continue;
                }
            }
            if let Some(extension) = path.extension() {
                if extension == "xlsx" || extension == "xlsm" {
                    excel_files.push(path.display().to_string());
                }
            }
        }
    }

    Ok(excel_files)
}

fn export_sheet_name(mut args_iter : &mut std::iter::Enumerate<std::env::Args>)
{
    let dir_path = args_iter.next().unwrap().1;
    let output_path = args_iter.next().unwrap().1;
    match list_excel_files_in_directory(dir_path.as_str()) {
        Ok(files) => {
            let num_threads =12;
            let chunk_size = (files.len() + num_threads - 1) / num_threads; // Ceiling division

            let (tx, rx) = mpsc::channel();

            for chunk in files.chunks(chunk_size) {
                let tx_clone = tx.clone();
                let chunk = chunk.to_owned();

                thread::spawn(move || {
                    let re = Regex::new(r"^[a-z0-9_]+$").unwrap();
                    let mut processed_files = Vec::new();
                    for file in chunk {
                        let mut str = String::new();
                        str = file.clone() + ":";
                        match open_workbook::<Xlsx<_>, &str>(file.as_str()) {
                            Ok(xlsx) => {
                                let mut sheet_names = xlsx.sheet_names();
                                let mut sheet_name = sheet_names.into_iter().filter(|x| re.is_match(x)).collect::<Vec<_>>();
                                for sheet in sheet_name {
                                    str += (sheet.as_str());
                                    str += ",";
                                }
                            },
                            Err(e) => { println!("err in file : {} err : {}",file, e); }
                        }

                        processed_files.push(str);
                    }
                    tx_clone.send(processed_files).unwrap();
                });
            }
            let mut results = Vec::new();
            for _ in 0..num_threads {
                let processed_files = rx.recv().unwrap();
                results.extend(processed_files);
            }

            fs::write(output_path, results.join("\n")).unwrap();
        },
        Err(e) => {
            eprintln!("Error occurred: {}", e);
        }
    }
}

fn main()
{
    let mut args_iter = env::args().enumerate();

    if env::args().count() == 4
    {
        // 실행 파일

        let instant = Instant::now();

        args_iter.next();
        // 여기서 as_str하면 임시값의 생명 주기가 박살나서 안 됨.

        let path = args_iter.next().unwrap().1;

        if path != "export_sheet_name"
        {
            export_sheet(&path, &mut args_iter);
        }
        else
        {
            export_sheet_name(&mut args_iter);
        }


        println!("find id total time : {:.2?}", instant.elapsed());
    }
    else
    {
        println!("use this usage : rapid_python_reader <xlsx_path> <output_path> <target_id>");
        return;
    }
}