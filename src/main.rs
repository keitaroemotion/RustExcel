extern crate simple_excel_writer;
use simple_excel_writer as excel;
use excel::*;

fn create_xlsx(records: Vec<Vec<&'static str>>) {
    let mut wb    = Workbook::create("tmp/b.xlsx");
    let mut sheet = wb.create_sheet ("SheetName");

    let columns = records[0]
                      .iter()
                      .map (|_| Column {width: 30.0});

    for column in columns {
        sheet.add_column(column);
    }

    wb.write_sheet(
        &mut sheet,
        |sheet_writer| {
            let sw = sheet_writer;
            sw.append_row(row!["Name", "Title", "Success", "Else"])?;
            for record in records {
                sw.append_row(
                    row![
                        record[0],
                        record[1],
                        record[2],
                        record[3]
                    ]
                ).expect("adding row failed");
            }
            Ok(())
        }
    )
    .expect("write excel error!");

    wb.close().expect("close excel error!");
}

fn main() {
    create_xlsx(vec![vec!["1","1","1","1"]]);
    println!("Hello, world!");
}
