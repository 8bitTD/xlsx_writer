use xlsx_writer::*;
fn main(){
    let xlsx_path = format!("{}{}{}",std::env::current_dir().unwrap().display().to_string(),"/","test.xlsx").replace("\\","/");
    let mut cells = Vec::new();
    let c = Cell::default().set_pos(1,1).set_content("_test");
    cells.push(c);
    let sheet1 =Sheet::default()
        .set_name("test_sheet")
        .set_cells(cells);
    let mut xlsx = Xlsx::default().add_sheet(sheet1).set_path(xlsx_path);
    xlsx.write_xlsx();
    while xlsx.is_write(){
        xlsx.update();
    }
    xlsx.open_xlsx();
}