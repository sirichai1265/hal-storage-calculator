param($Html,$Attach,$Sub,$To)
# สร้างดราฟใหม่ (ไม่ส่ง) · ถ้ามีดราฟหัวข้อเดียวกันอยู่แล้ว ให้อัปเดตเนื้อหา/ผู้รับ/ไฟล์แนบของดราฟนั้นแทนการสร้างซ้ำ
$ol=New-Object -ComObject Outlook.Application
$acc=$ol.Session.Accounts | Where-Object { $_.SmtpAddress -ieq 'sirichai@heungaline.co.th' } | Select-Object -First 1
$m=$null
foreach($i in @($ol.Session.GetDefaultFolder(16).Items)){ if($i.Subject -eq $Sub){ $m=$i; break } }
$new = ($m -eq $null)
if($new){ $m=$ol.CreateItem(0) } else { for($k=$m.Attachments.Count;$k -ge 1;$k--){ $m.Attachments.Remove($k) } }
$m.SendUsingAccount=$acc
$m.To=$To
$m.Subject=$Sub
$m.HTMLBody=[IO.File]::ReadAllText($Html,[Text.Encoding]::UTF8)
[void]$m.Attachments.Add($Attach)
$m.Save()
"$(if($new){'created'}else{'updated'}): $Sub | att=$($m.Attachments.Count) | from $($acc.SmtpAddress)"
