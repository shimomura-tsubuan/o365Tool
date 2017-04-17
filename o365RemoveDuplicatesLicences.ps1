#$lastLoggedUser = (Get-Item -Path Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI).GetValue("LastLoggedOnUser")
#$lastLoggeeUserSamName = $lastLoggedUser.split("\")[1]

# Officeのプロセスの存在を確認し、存在する場合はスクリプトを終了する。
$msOfficeCheck = (Get-Process).ProcessName | Select-String "POWERPNT|WINWORD|MSACCESS|EXCEL" -quiet
if($msOfficeCheck){ exit }

# プロダクトキーが２つインストールされていないか確認。
$getDstatus = cscript "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" /dstatus
$resultCount = ($getDstatus | Select-String "Last 5 characters of installed product key:").Count
if(!(Test-Path "C:\temp")) {New-Item -Type Directory "C:\temp"}
$getDstatus | Out-File "C:\temp\o365RemoveDuplicatesLicences.log" -Encoding Default

# プロダクトキーが２つあるときの処理。(一度では削除しきれない場合があるのでループする。)
if(2 -le $resultCount){
  While($True){
    $getDstatus = cscript "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" /dstatus
    $resultBool = $getDstatus | Select-String "Last 5 characters of installed product key:" -quiet
    IF($resultBool){
      $productKey = (($getDstatus | Select-String "Last 5 characters of installed product key:") -split ": ")[1]
      cscript "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" /unpkey:$productKey
    } else {
      break
    }
      Sleep 5
  }
}
