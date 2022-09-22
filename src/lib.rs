mod sheet;

pub fn add(left: usize, right: usize) -> usize {
    left + right
}

use crate::sheet::sheet::{get_sheet, get_rows_by_id};
use pyo3::prelude::*;

#[pyfunction]
pub fn get_sheet_from_rust( path : &str, sheet_index : u32, id_row_index :  u32) -> PyResult<String> {
    return Ok(get_sheet(path, sheet_index, id_row_index));
}

#[cfg(test)]
mod tests {
    use super::*;

    use std::time::{Duration, SystemTime};
    #[test]
    fn test_get_sheet() {
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

    #[test]
    fn it_works() {
        let result = add(2, 2);
        assert_eq!(result, 4);
    }
}
