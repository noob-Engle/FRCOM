use std::fs::File;
use std::io::Write;
use winreg::enums::*;
use winreg::RegKey;

fn main() {
    let hkcr = RegKey::predef(HKEY_CLASSES_ROOT);
    let clsid_key = hkcr.open_subkey_with_flags("CLSID", KEY_READ).unwrap();

    let mut file = File::create("clsids.txt").expect("Unable to create file");

    for clsid in clsid_key.enum_keys() {
        match clsid {
            Ok(clsid) => writeln!(file, "{}", clsid).expect("Unable to write data"),
            Err(e) => eprintln!("Error reading CLSID: {:?}", e),
        }
    }

    println!("CLSID keys have been written to clsids.txt");
}