# FRCOM
Rust Windows COM Code Sinppet

> build
```bash
cargo build
```
## INFO 

- FindCLSID.exe  Get CLSID
- FindCom.exe    CLSID Call Com 

<img width="1208" alt="image" src="https://github.com/noob-Engle/FRCOM/assets/82130997/afb90940-efd5-4477-b8c7-d2fb9413e573">

> Print CLSID Object Func  or Attribute

```ps1
New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR -ErrorAction SilentlyContinue
$clsid = "0D43FE01-F093-11CF-8940-00A0C9054228"
$clsidPath = "HKCR:\CLSID\{$clsid}"
$clsidProperties = Get-ItemProperty -Path $clsidPath
Write-Host "Properties for CLSID {$clsid}:"
$clsidProperties | Format-List
Write-Host "Methods for CLSID {$clsid}:"
$clsidProperties | Get-Member -MemberType Method
```



<img width="1316" alt="image" src="https://github.com/noob-Engle/FRCOM/assets/82130997/6547bf8c-bf7a-4cca-bcc9-f9304856e280">




## TodoList
- [*] FindCLSID
- [*] Call COM Success 
