
pub mod sheet
{
    use std::collections::HashMap;
    use std::fs::File;
    use std::hash::Hash;
    use std::io::{BufReader, Error};
    use std::string::String;
    use calamine::{Reader, open_workbook, Xlsx, DataType, Range};
    use serde_json::{Value};

    type SheetAndIndex = (HashMap<usize, String>, Vec<(String, Range<DataType>)>);
    pub fn get_sheets_and_index_to_id(xlsx_path : &str, sheet_index : u32, id_row_index : u32) -> Option<SheetAndIndex>
    {
        let error_string = format!("Failed to open workbook {}",  xlsx_path);
        //open workbook
        let mut workbook_option : Option<Xlsx<BufReader<File>>> = None;

        match open_workbook::<Xlsx<_>, &str>(xlsx_path) {
            Ok(xlsx) => { workbook_option = Some(xlsx); },
            Err(e) => { println!("{}",e); return None; }
        }

        let mut workbook = workbook_option.unwrap();

        if workbook.worksheets().len() <= sheet_index as usize {
            panic!("{}",error_string.as_str());
        }

        let worksheet = &workbook.worksheets()[sheet_index as usize];

        let range = &worksheet.1;
        //println!("worksheet {} reading.", worksheet.0);
        //println!("cell count : {}", range.get_size().0 * range.get_size().1);
        let mut index_to_id : HashMap<usize, String>  = HashMap::new();

        // row id index searching.
        let id_row_range = range.range((id_row_index, 0), (id_row_index, range.get_size().1 as u32));

        id_row_range.cells().for_each(| cell| {
            match cell.2.get_string() {
                Some(str) => index_to_id.insert(cell.1,str.to_string()),
                None => { None }
            };
        });

        return Some((index_to_id, workbook.worksheets()));
    }

    pub fn row_to_push_string( ref_string : &mut String, index_to_id : &HashMap<usize, String>, row : &[DataType]) {
        //ref_string.clear();
        ref_string.push_str("{");

        let mut iter_string = String::new();


        row.iter().enumerate().for_each(|(index, cell)| {
            match cell.get_string() {
                Some(str) => {
                    match index_to_id.get(&index)  {
                        Some(id) => {
                            iter_string.push_str(format!("\"{}\": \"{}\",", id, str).as_str());
                        },
                        None => {}
                    }
                },
                None => {}
            };
        });

        iter_string.pop();
        ref_string.push_str(iter_string.as_str());
        ref_string.push_str( "}," );
    }



    pub fn row_to_push_json( ref_json : &mut Vec<serde_json::Value>, index_to_id : &HashMap<usize, String>, row : &[DataType]) {
        let mut json = serde_json::json!({});

        row.iter().enumerate().for_each(|(index, cell)| {
            match cell.get_string() {
                Some(str) => {
                    match index_to_id.get(&index)  {
                        Some(id) => {
                            json[id] = str.into();
                        },
                        None => {}
                    }
                },
                None => {}
            };
        });

        ref_json.push(json );
    }

    pub fn get_sheet( xlsx_path : &str, sheet_index : u32, id_row_index : u32) -> Option<String> {
        let sheet_and_index_option : _ = get_sheets_and_index_to_id(xlsx_path, sheet_index, id_row_index);

        let mut option_data : Option<SheetAndIndex> = None;
        match sheet_and_index_option {
            Some((index_to_id, sheets)) => {
                option_data = Some((index_to_id, sheets));
            },
            None => {return None;}
        }

        let (index_to_id, sheets) = option_data.unwrap();
        //println!("index to id index : {:?}", index_to_id);

        let range = &sheets[sheet_index as usize].1;

        let mut value_vector : Vec<serde_json::Value> = Vec::new();

        for row in range.range( (id_row_index + 1, 0), (range.get_size().0 as u32, range.get_size().1 as u32) ).rows() {
            row_to_push_json(&mut value_vector, &index_to_id, row);
        }

        let json_value = Value::Array(value_vector);
        return Some(serde_json::to_string_pretty(&json_value).unwrap());
    }

    pub fn get_main_id_list( xlsx_path : &str, sheet_index : u32, id_row_index : u32) -> Option<String> {
        let sheet_and_index_option : _ = get_sheets_and_index_to_id(xlsx_path, sheet_index, id_row_index);

        let mut option_data : Option<SheetAndIndex> = None;
        match sheet_and_index_option {
            Some((index_to_id, sheets)) => {
                option_data = Some((index_to_id, sheets));
            },
            None => {return None;}
        }

        let (index_to_id, sheets) = option_data.unwrap();
        //println!("index to id index : {:?}", index_to_id);

        let range = &sheets[sheet_index as usize].1;


        let mut result_str = String::new();

        for row in range.range( (id_row_index + 1, 0), (range.get_size().0 as u32, range.get_size().1 as u32) ).rows() {
            match row[0].get_string() {
                Some(str) => {
                    match index_to_id.get(&0)  {
                        Some(id) => {
                            // String => Object
                            // &str => View용 객체 (Read만)
                            result_str.push_str((str.to_owned() + ",").as_str());
                        },
                        None => {}
                    }
                },
                None => {}
            };
        }

        result_str.remove(result_str.len() - 1); // 마지막에 쉼표 제거

        return Some(result_str);
    }

    pub fn get_rows_by_id( xlsx_path : &str, sheet_index : u32, id_row_index : u32, find_id : &str) -> Option<String> {
        let sheet_and_index_option : _ = get_sheets_and_index_to_id(xlsx_path, sheet_index, id_row_index);

        let mut option_data : Option<SheetAndIndex> = None;
        match sheet_and_index_option {
            Some((index_to_id, sheets)) => {
                option_data = Some((index_to_id, sheets));
            },
            None => {return None;}
        }

        let (index_to_id, sheets) = option_data.unwrap();
        //println!("index to id index : {:?}", index_to_id);;

        let range = &sheets[sheet_index as usize].1;
        let mut id_find = false;
        let mut result_string = String::new();
        result_string.push_str("[" );
        for row in range.range( (id_row_index + 1, 0), (range.get_size().0 as u32, range.get_size().1 as u32) ).rows() {

            //println!("row[0].to_string() {}", row[0].to_string());
            if row[0].to_string() != find_id && !(id_find && row[0].is_empty()) {
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

        if !id_find {
            return None;
        }
        result_string.pop(); // 마지막에 쉼표 제거
        result_string.push_str("]" );

        return Some(result_string);
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
        }

    }
}