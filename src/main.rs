extern crate core;

use std::string::String;
use calamine::{Reader, open_workbook, Xlsx, DataType, Range};
use std::time::{Duration, SystemTime};


fn get_sheets_and_index_to_id(xlsx_path : &str, sheet_index : u32, id_row_index : u32) -> (Vec<String>, Vec<(String, Range<DataType>)>)
{
    let error_string = format!("Failed to open workbook {}",  xlsx_path);
    //open workbook
    let mut workbook : Xlsx<_> = open_workbook(xlsx_path).expect(error_string.as_str());

    if workbook.worksheets().len() <= sheet_index as usize {
        panic!("{}",error_string.as_str());
    }

    let worksheet = &workbook.worksheets()[sheet_index as usize];

    let range = &worksheet.1;
    println!("worksheet {} reading.", worksheet.0);
    //println!("cell count : {}", range.get_size().0 * range.get_size().1);
    let mut index_to_id : Vec<String>  = Vec::new();

    // row id index searching.
    let id_row_range = range.range((id_row_index, 0), (id_row_index, range.get_size().1 as u32));

    id_row_range.cells().for_each(| cell| {
        match cell.2.get_string() {
            Some(str) => index_to_id.push(str.to_string()),
            None => {}
        };
    });

    return (index_to_id, workbook.worksheets());
}

fn row_to_push_string( ref_string : &mut String, index_to_id : &Vec<String>, row : &[DataType]) {
    ref_string.clear();
    ref_string.push_str("{");

    row.iter().enumerate().for_each(|(index, cell)| {
        match cell.get_string() {
            Some(str) => {
                ref_string.push_str( format!("\"{}\": {},", index_to_id[index], str).as_str() );
            },
            None => {}
        };
    });

    ref_string.push_str( "}," );
}

fn get_sheet( xlsx_path : &str, sheet_index : u32, id_row_index : u32) -> String {
    let (index_to_id, sheets) = get_sheets_and_index_to_id(xlsx_path, sheet_index, id_row_index);
    println!("index to id index : {:?}", index_to_id);;

    let range = &sheets[sheet_index as usize].1;

    let mut result_string = String::new();
    for row in range.range( (id_row_index + 1, 0), (range.get_size().0 as u32, range.get_size().1 as u32) ).rows() {
        row_to_push_string(&mut result_string, &index_to_id, row);
    }

    println!("{}", result_string);

    return result_string;
}

fn get_rows_by_id( xlsx_path : &str, sheet_index : u32, id_row_index : u32, find_id : &str) -> String {
    let (index_to_id, sheets) = get_sheets_and_index_to_id(xlsx_path, sheet_index, id_row_index);
    println!("index to id index : {:?}", index_to_id);;

    let range = &sheets[sheet_index as usize].1;
    let mut id_find = false;
    let mut result_string = String::new();
    for row in range.range( (id_row_index + 1, 0), (range.get_size().0 as u32, range.get_size().1 as u32) ).rows() {
        if row[0].to_string() != find_id || !(id_find && row[0].is_empty()) {
            if id_find {
                break;
            }
            else {
                continue;
            }
        }
        else {
            id_find = true;
            row_to_push_string(&mut result_string, &index_to_id, row);
        }
    }

    println!("{}", result_string);

    return result_string;
}

fn default_test(path : &str)
{
    let mut workbook: Xlsx<_> = open_workbook(path).expect("Cannot open file");

    // workbook.worksheets()[0] => 첫번째 거 가져다 쓴다.
    // Read whole worksheet data and provide some statistics
    if let worksheet_tuple = &workbook.worksheets()[0] {
        let range = &worksheet_tuple.1;
        let total_cells = range.get_size().0 * range.get_size().1;
        let non_empty_cells: usize = range.used_cells().count();
        println!("Found {} cells in 'Sheet1', including {} non empty cells",
                 total_cells, non_empty_cells);

        for col in range.rows() {
            let col_0 = &col[0];
            match col_0.get_string() { //이건 Result가 아니라 Option이고, 그래서  Some None 반환이다.
                Some(val) => {
                    match val {
                        "" => {},
                        _ => {
                            println!("{}", val);
                            break;
                        }
                    }
                },
                None => {}
            }
        }

        fn test() -> bool { println!("range"); true}
        assert!(test());
        /*
                // alternatively, we can manually filter rows
                assert_eq!(non_empty_cells, range.rows()
                    .flat_map(|r| r.iter().filter(|&c| c != &DataType::Empty)).count());

         */
    } else {
        panic!("Cannot find 'Sheet1' in 'test.xlsx'");
    }

}

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
    get_rows_by_id(path, 0, 1, "CH01_003_5");
    let find_ch00_elapsed = match find_ch00.elapsed() {
        Err(e) => panic!("{}", e),
        Ok(t) => t,
    };

    let find_ch00_current_elapsed = match SystemTime::now().elapsed() {
        Err(e) => panic!("{}", e),
        Ok(t) => t
    };


    println!("find ch00_000 total time : {}", find_ch00_elapsed.as_millis() - find_ch00_current_elapsed.as_millis());

}