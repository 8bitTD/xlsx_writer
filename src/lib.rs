
use std::io::prelude::*;
use std::os::windows::process::CommandExt;

#[derive(Debug, Clone)]
pub struct Cell{
    px: usize,
    py: usize,
    font_col_index: Option<usize>,
    bg_col_index: Option<usize>,
    content: String,
    hyperlink: Option<String>,
    validation: Option<String>,
}
impl Default for Cell{
    fn default() -> Self{
        Self { px: 1, py: 1, font_col_index: None, bg_col_index: None, content: String::from(""), hyperlink: None, validation: None }
    }
}
impl Cell{
    pub fn set_pos(&mut self, y: usize, x: usize) -> Self{
        if (x < 1) || (y < 1){return self;}
        self.px = x;
        self.py = y;
        self
    }
    pub fn set_content(&mut self, content: &str) -> Self{
        self.content = content.to_string();
        self
    }
    pub fn set_font_col_index(mut self, index: usize) -> Self{
        self.font_col_index = Some(index);
        self
    }
    pub fn set_bg_col_index(mut self, index: usize) -> Self{
        self.bg_col_index = Some(index);
        self
    }
    pub fn set_hyperlink(mut self, path: &str) -> Self{
        self.hyperlink = Some(path.to_string());
        self
    }
    pub fn set_valication(mut self, validation: &str) -> Self{
        self.validation = Some(String::from(validation));
        self
    }
}

#[derive(Debug, Clone)]
pub struct Width{
    pub px: usize,
    pub width: usize,
}
#[derive(Debug, Clone)]
pub struct Line{
    range: String,
    number: usize,
}

#[derive(Debug, Clone)]
pub struct Sheet{
    pub name: String,
    pub cells: Vec<Cell>,
    pub widths: Vec<Width>,
    pub sort: Option<String>,
    pub lines: Vec<Line>
}
impl Default for Sheet{
    fn default() -> Self{
        Self { 
            name: String::from(""),
            cells: Vec::new(),
            widths: Vec::new(),
            sort: None,
            lines: Vec::new(),
        }
    }
}
impl Sheet{
    pub fn set_name(mut self, name: &str) -> Self{
        self.name = String::from(name);
        self
    }
    pub fn set_cells(mut self, cells: Vec<Cell>) -> Self{
        self.cells = cells;
        self
    }
    pub fn add_cell(mut self, cell: Cell) -> Self{
        self.cells.push(cell);
        self
    }
    pub fn set_widths(mut self, widths: Vec<Width>) -> Self{
        self.widths = widths;
        self
    }
    pub fn add_width(mut self, px: usize, width: usize) -> Self{
        self.widths.push(Width{px, width});
        self
    }
    pub fn set_sort(mut self, y: usize, sx: usize, ex:usize) -> Self{
        let st = format!("{}{}",get_alpabet_from_num(sx),y);
        let ed = format!("{}{}",get_alpabet_from_num(ex),y);
        self.sort = Some(format!("{}{}{}",st,":",ed));
        self
    }
    pub fn add_line(mut self, sy:usize, ey:usize, sx:usize, ex:usize, num:usize) -> Self{
        let line = Line{
            range: format!("{}{}{}{}{}", get_alpabet_from_num(sx),sy,":",get_alpabet_from_num(ex),ey),
            number: num,
        };
        self.lines.push(line);
        self
    }
}

#[derive(Debug, PartialEq)]
pub enum XlsxState{
    Idle,
    Write,
}

#[derive(Debug)]
pub struct Xlsx{
    pub xlsx_path: String,
    pub rxs_write: Option<std::sync::mpsc::Receiver<bool>>,
    pub state: XlsxState,
    pub sheets: Vec<Sheet>,
    pub rc: std::sync::Arc<std::sync::Mutex<String>>,
}

impl Default for Xlsx{
    fn default() -> Self{
        let dt_path = dirs::desktop_dir().unwrap().as_os_str().to_str().unwrap().to_string().replace("\\","/");
        let datetime = chrono::Utc::now().with_timezone(&chrono::FixedOffset::east_opt(9 * 3600).unwrap()).naive_local();
        let dt_path = format!("{}{}{}",dt_path,"/xlsx_",datetime.format("%Y%m%d%H%M%S").to_string());
        let xlsx_path = format!("{}{}",dt_path,".xlsx");
        Self { 
            xlsx_path: xlsx_path,
            rxs_write: None,
            state: XlsxState::Idle,
            sheets: Vec::new(),
            rc: std::sync::Arc::new(std::sync::Mutex::new(String::from(""))),
        }
    }
}

impl Xlsx{
    pub fn set_path(mut self, path: String) -> Self{
        self.xlsx_path = path;
        self
    }
    pub fn add_sheet(mut self, sheet: Sheet) -> Self {
        self.sheets.push(sheet.clone());
        self
    }
    pub fn update(&mut self){
        if self.rxs_write.is_none(){return;}
        match self.rxs_write.as_ref().unwrap().try_recv(){
            Ok(_c)=>{ self.state = XlsxState::Idle; },
            Err(_e) =>{ }
        };
    }
    pub fn is_write(&mut self) -> bool{
        self.state == XlsxState::Write
    }

    pub fn write_xlsx(&mut self) {
        self.state = XlsxState::Write;
        self.rxs_write = write_xlsx(self.xlsx_path.clone(), self.sheets.clone(), &self.rc);
    }

    pub fn open_xlsx(&self){
        std::process::Command::new("cmd")
            .args(&["/C", &self.xlsx_path])
            .creation_flags(0x08000000)
            .current_dir("C:\\Users")
            .spawn().unwrap();
            //*crc.lock().unwrap() = format!("{}","エクセルを起動しています");
    }
}

fn get_alpabet_from_num(num:usize) -> String{
    match num {
        1  => {String::from("A")},
        2  => {String::from("B")},
        3  => {String::from("C")},
        4  => {String::from("D")},
        5  => {String::from("E")},
        6  => {String::from("F")},
        7  => {String::from("G")},
        8  => {String::from("H")},
        9  => {String::from("I")},
        10 => {String::from("J")},
        11 => {String::from("K")},
        12 => {String::from("L")},
        13 => {String::from("M")},
        14 => {String::from("N")},
        15 => {String::from("O")},
        16 => {String::from("P")},
        17 => {String::from("Q")},
        18 => {String::from("R")},
        19 => {String::from("S")},
        20 => {String::from("T")},
        21 => {String::from("U")},
        22 => {String::from("V")},
        23 => {String::from("W")},
        24 => {String::from("X")},
        25 => {String::from("Y")},
        26 => {String::from("Z")},
        _ => {String::from("Z")},
    }
}


fn write_xlsx(xlsx_path: String, sheets: Vec<Sheet>, rc: &std::sync::Arc<std::sync::Mutex<String>>) -> Option<std::sync::mpsc::Receiver<bool>>{
    let (tx, rx) = std::sync::mpsc::channel();
    let crc= std::sync::Arc::clone(&rc);
    std::thread::spawn(move || {
        let ps_path = write_ps1(&xlsx_path, &sheets, &crc);
        let mut child = std::process::Command::new("cmd")
            .args(&["/C", "powershell -NoProfile -ExecutionPolicy Unrestricted",&ps_path])
            .creation_flags(0x08000000)
            .current_dir("C:\\Users")
            .spawn().unwrap();
        *crc.lock().unwrap() = format!("{}","パワーシェルを実行しています");
        let _result = child.wait().unwrap();
        rm_rf::remove(&ps_path).unwrap();//psファイルを削除する処理
        *crc.lock().unwrap() = format!("{}","パワーシェルファイルを削除しています");
        /*
        let cld = std::process::Command::new("cmd")
            .args(&["/C", &xlsx_path])
            .creation_flags(0x08000000)
            .current_dir("C:\\Users")
            .spawn().unwrap();
            *crc.lock().unwrap() = format!("{}","エクセルを起動しています");
        */
        tx.send(true).unwrap();
    });
    Some(rx)
}

fn write_ps1(xlsx_path: &String, sheets: &Vec<Sheet>, rc: &std::sync::Arc<std::sync::Mutex<String>>) -> String{
    let ps_path = xlsx_path.replace(".xlsx", ".ps1");
    let mut cmd = Vec::new();
    cmd.push(format!("$excel = New-Object -ComObject Excel.Application;\n"));
    cmd.push(format!("$excel.Visible = $false;\n"));
    cmd.push(format!("$book = $excel.Workbooks.Add();\n"));
    cmd.push(format!("$excel.DisplayAlerts = $false;\n"));
    
    for (_i, s) in sheets.iter().enumerate(){
        cmd.push(format!("$sheet = $book.Worksheets.Add();\n"));
        cmd.push(format!("$sheet.Name = {}{}{};{}", "\"",s.name, "\"", "\n"));
        *rc.lock().unwrap() = format!("{}","ラインを設定をしています");
        for (_,c) in s.cells.iter().enumerate(){
            if c.hyperlink.is_some(){
                let hl = c.hyperlink.as_ref().unwrap().replace("/","\\");
                cmd.push(format!(r#"$sheet.Hyperlinks.Add($sheet.Cells.Item({},{}),"{}","","{}","{}") | Out-Null;{}"#, c.py, c.px, hl,"", c.content,"\n"));
            }else{
                cmd.push(format!("$sheet.Cells.Item({},{}) = {}{}{}{}", c.py, c.px,"\"",c.content,"\"",";\n"));
            }
            if c.font_col_index.is_some(){
                cmd.push(format!("$sheet.Cells.Item({},{}).Font.ColorIndex = {};{}",c.py, c.px, c.font_col_index.unwrap(),"\n"));
            }
            if c.bg_col_index.is_some(){
                cmd.push(format!("$sheet.Cells.Item({},{}).Interior.ColorIndex = {};{}",c.py, c.px, c.bg_col_index.unwrap(),"\n"));
            }
            if c.validation.is_some(){
                cmd.push(format!("$sheet.Cells.Item({},{}).Validation.Delete();{}",c.py, c.px, "\n"));
                cmd.push(format!(r#"$sheet.Cells.Item({},{}).Validation.Add(3, 1, 1, "{}");{}"#,c.py, c.px, c.validation.as_ref().unwrap(), "\n"));
            }
        }
        *rc.lock().unwrap() = format!("{}","セルの幅を設定をしています");
        for w in &s.widths{
            cmd.push(format!("$sheet.columns.item({}).columnWidth = {};\n", w.px, w.width));
        }
        *rc.lock().unwrap() = format!("{}","ラインを設定をしています");
        for l in &s.lines{
            cmd.push(format!(r#"$sheet.Range("{}").Borders.LineStyle = {};{}"#, l.range, l.number, "\n"));
        }
        *rc.lock().unwrap() = format!("{}","フィルターを設定をしています");
        if s.sort.is_some(){
            cmd.push(format!(r#"$sheet.Range("{}").AutoFilter() | Out-Null;{}"#, s.sort.as_ref().unwrap(), "\n"));
        }
        cmd.push(format!("{0}{1}{2}{3}{4}","$sheet.Cells.font.Name = ","\"","Meiryo UI","\"",";\n"));
    }

    cmd.push(format!("$excel.Worksheets.item({}{}{}).delete();{}","\"","Sheet1", "\"", "\n"));
    cmd.push(format!("{0}{1}{2}{3}{4}","$book.SaveAs(","\"",xlsx_path.replace("/","\\"),"\"",");\n"));
    cmd.push("$excel.Quit();\n".to_string());
    cmd.push("$excel = $null;\n".to_string());
    cmd.push("[GC]::Collect();\n".to_string());
    let mozi = cmd.iter().cloned().collect::<String>();
    let (cow, _, _) = encoding_rs::SHIFT_JIS.encode(&mozi);
    let mut file = std::fs::File::create(&ps_path).unwrap();
    file.write_all(&cow).unwrap();
    ps_path
}

#[cfg(test)]
mod tests {
    //use super::*;
    #[test]
    fn it_works() {
        
    }
}
