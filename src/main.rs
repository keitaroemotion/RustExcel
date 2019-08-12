extern crate simple_excel_writer;
use simple_excel_writer as excel;
use std::io;
use excel::*;

fn main() {
    let mut wb    = Workbook::create("tmp/b.xlsx");
    let mut sheet = wb.create_sheet("SheetName");

    sheet.add_column(Column { width: 30.0 });
    sheet.add_column(Column { width: 30.0 });
    sheet.add_column(Column { width: 80.0 });
    sheet.add_column(Column { width: 60.0 });

    wb.write_sheet(
        &mut sheet,
        |sheet_writer| {
            let sw = sheet_writer;
            sw.append_row(row!["Name", "Title", "Success", "XML Remark"])?;
            sw.append_row(row!["Amy",  (),      true,      "AAA"       ])?;
            sw.append_blank_rows(2);
            sw.append_row(row!["Tony", blank!(30), "retired"           ])
        }
    )
    .expect("write excel error!");

    wb.close().expect("close excel error!");
    println!("Hello, world!");
}
