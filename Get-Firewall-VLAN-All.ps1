<#
.SYNOPSIS
    VGC的网关基本都在防火墙上，收集IP信息时，关联照网关接口地址更便于团队查询
    该脚本将提取防火墙接口数据

.DESCRIPTION
    读取网络组从SmartCenter上导出的VLAN接口数据，如果有多个，可手动合并为一个文件
    与豆腐坊生成的VLAN数据做匹配
    过滤不需要的数据后
    最终生成可提供给客户的信息

    重要：原始数据必须以制表线为第一行

.NOTES
    ===========================================================================
     Created by:   	Michael Gao
     Last Update:   April 28, 2019
     
    ===========================================================================
#>

param($Drive = "C")

ac $home\TimeCalc.txt "$(Get-Date), Firewall-VLAN started"

$rslt = "$Drive`:\Out\FW\YiZhuang-FireWall-VLAN_" + $(Get-Date -UFormat "%Y-%m-%d") + ".xlsx"
$inputCSV = (gci "$Drive`:\Out\FW\Firewall-VLAN_*.txt" | select -Last 1).FullName

$xl = New-Object -ComObject excel.application 
$xl.Visible = $true
$xl.DisplayAlerts = $false

[void]$xl.Workbooks.Add()
$xl.ActiveWindow.WindowState = -4137
$sh = $xl.Worksheets.Item(1)
$sh.Name = "RAWData"

$TxtConnector = ("TEXT;" + $inputCSV)
$Connector = $sh.QueryTables.add($TxtConnector,$sh.Range("A1"))
$query = $sh.QueryTables.item($Connector.name)

$query.TextFileOtherDelimiter = "|"
$query.TextFileParseType  = 1
$query.TextFileColumnDataTypes = ,2 * $sh.Cells.Columns.Count
$query.AdjustColumnWidth = 1
$query.Refresh()
rv TxtConnector,Connector,query

[void]$sh.Columns.Item(1).Delete()
[void]$sh.Rows.Item(1).Delete()

[void]$sh.usedrange.autofilter(1,"=",1)
[void]$sh.usedrange.specialcells(12).EntireRow.Select()
[void]$sh.usedrange.specialcells(12).EntireRow.delete()
for($i = 1;$i -le $sh.UsedRange.Rows.Count;$i++){[void]$sh.Range("e$i").Activate();$sh.Range("e$i").Value2 = $sh.Range("a$i").text.trim()}
for($i = 1;$i -le $sh.UsedRange.Rows.Count;$i++){[void]$sh.Range("f$i").Activate();$sh.Range("f$i").Value2 = $sh.Range("b$i").text.trim()}
for($i = 1;$i -le $sh.UsedRange.Rows.Count;$i++){[void]$sh.Range("g$i").Activate();$sh.Range("g$i").Value2 = $sh.Range("c$i").text.trim()}
for($i = 1;$i -le $sh.UsedRange.Rows.Count;$i++){[void]$sh.Range("h$i").Activate();$sh.Range("h$i").Value2 = $sh.Range("d$i").text.trim()}
[void]$sh.Columns.Range("a:d").Delete()

[void]$sh.Range("a1").Activate()
[void]$sh.Rows.Item(1).Insert(-4121)
$sh.Range('a1').value2 = "Type/Interface"
$sh.Range('b1').value2 = "Firewall"
$sh.Range('c1').value2 = "VSID"
$sh.Range('d1').value2 = "IP/Mask"
[void]$sh.range('a1:d1').entirecolumn.autofit()
$sh.Application.ActiveWindow.SplitColumn = 0
$sh.Application.ActiveWindow.SplitRow = 1
$sh.Application.ActiveWindow.FreezePanes = $True

[void]$sh.Columns.Item(2).Insert(-4161)

#$DataType = 1 #DataType As XlTextParsingType = xlDelimited
#$XlTextParsingType = 1 #使用分割符
#$xlTextQualifier = 1 #TextQualifier As XlTextQualifier = xlTextQualifierDoubleQuote
#$ConsecutiveDelimiter = $true
#$xlTextQualifierDoubleQuote = 1
#$TextFileTabDelimiter = $false
#$TextFileSemicolonDelimiter = $false
#$TextFileCommaDelimiter = $false
#$TextFileSpaceDelimiter = $true #使用空格做分割符
#$TextFileOtherDelimiter = $false
#$TrailingMinusNumbers = $true
#[void]$sh.Columns.Item(1).Select()
#[void]$sh.Columns.item(1).TextToColumns("A:A")
#for($i = 2;$i -le $sh.UsedRange.Rows.Count;$i++){[void]$sh.Range("a$i").Activate();$sh.Range("a$i").TextToColumns()}
[void]$sh.Columns.item(1).TextToColumns($sh.range('a1'),1,1,$true,$false,$false,$false,$true,$false,$true)
[void]$sh.Columns.item(5).TextToColumns($sh.range('e1'),1,1,$true,$false,$false,$false,$true,$false,$true)
[void]$sh.Columns.Item(3).Insert(-4161)
[void]$sh.Columns.item(2).TextToColumns($sh.range('b1'),1,1,$true,$false,$false,$false,$false,$true,".")
[void]$sh.Columns.item(7).TextToColumns($sh.range('g1'),1,1,$true,$false,$false,$false,$false,$true,"/")
$sh.Range('a1').value() = "Type"
$sh.Range('b1').value() = "Interface"
$sh.Range('c1').value() = "V-ID"
$sh.Range('f1').value() = "Ver"
$sh.Range('g1').value() = "GW"
$sh.Range('h1').value() = "Mask"
[void]$sh.Range('a1:h1').entirecolumn.autofit()

[void]$sh.usedrange.autofilter(1,"A",1)
[void]$sh.usedrange.specialcells(12).EntireRow.Select()
[void]$sh.usedrange.specialcells(12).EntireRow.delete()
[void]$sh.Rows.Item(1).Insert(-4121)
$sh.Range('a1').value() = "Type"
$sh.Range('b1').value() = "Interface"
$sh.Range('c1').value() = "V-ID"
$sh.Range('d1').value() = "Firewall"
$sh.Range('e1').value() = "VSID"
$sh.Range('f1').value() = "Ver"
$sh.Range('g1').value() = "GW"
$sh.Range('h1').value() = "Mask"
$sh.Application.ActiveWindow.SplitColumn = 0
$sh.Application.ActiveWindow.SplitRow = 1
$sh.Application.ActiveWindow.FreezePanes = $True

[void]$xl.worksheets.add([system.type]::missing,$xl.worksheets.item($xl.worksheets.count))
$sh2 = $xl.Worksheets.Item(2)
$sh2.Name = "VLAN"
[void]$sh2.Activate()
for($i = 2;$i -le $sh.UsedRange.Rows.Count;$i++){
    [void]$sh2.Cells.Item($i,1).Activate()
    $sh2.Cells.Item($i,1).value() = $sh.Cells.Item($i,1).value()
    $sh2.Cells.Item($i,2).value() = $sh.Cells.Item($i,2).value()
    $sh2.Cells.Item($i,3).value() = $sh.Cells.Item($i,3).value()
    $sh2.Cells.Item($i,4).value() = $sh.Cells.Item($i,7).value()
    $sh2.Cells.Item($i,5).value() = $sh.Cells.Item($i,8).value()
    $sh2.Cells.Item($i,6).value() = $sh.Cells.Item($i,4).value()
    $sh2.Cells.Item($i,7).value() = $sh.Cells.Item($i,5).value()
}
    
[void]$sh2.Range("a1").Activate()
$sh2.Range('a1').value() = "Type"
$sh2.Range('b1').value() = "Interface"
$sh2.Range('c1').value() = "VLAN"
$sh2.Range('d1').value() = "Gateway"
$sh2.Range('e1').value() = "Mask"
$sh2.Range('f1').value() = "Device"
$sh2.Range('g1').value() = "VSID"

[void]$sh2.range('a1:f1').entirecolumn.autofit()
$sh2.Application.ActiveWindow.SplitColumn = 0
$sh2.Application.ActiveWindow.SplitRow = 1
$sh2.Application.ActiveWindow.FreezePanes = $True

$sh = $xl.Worksheets.Item(2)
$sh.Range('i1').value2 = "SubNet"
$sh.Range('j1').value2 = "SubNetMask"
# $sh.Range('k1').value2 = "FirstIP"
# $sh.Range('l1').value2 = "LastIP"
# $sh.Range('m1').value2 = "Usable"

for($l = 2;$l -le $sh.UsedRange.Rows.Count;$l++) {
    If($sh.Cells.Item($l,4).value()){
        $IP = $sh.Cells.Item($l,4).value()
        # 查看IP地址及二进制串信息
        #     $ip.Split('.') | ForEach-Object {
        #         '{0,5} : {1}' -f $_, [System.Convert]::ToString($_,2).PadLeft(8,'0')
        #     }
        #将IP地址的每一段转换为二进制并组合为一个串
        $IPBin = -join ($ip.Split('.') | ForEach-Object {[System.Convert]::ToString($_,2).PadLeft(8,'0')})
        #获取掩码位
        $length = $sh.Cells.Item($l,5).value()

        $Mask = ""
        #将掩码长度的填制位全部置1
        for($i = 0;$i -lt $length;$i++){$Mask = $Mask + "1"}
        #补全32位完整掩码
        for($i = $length;$i -lt 32;$i++){$Mask = $Mask + "0"}

        #计算子网，将IP地址和掩码进行【与】操作，得到子网的二进制串
        $net = ""
        for($i = 0;$i -lt 32;$i++) {
            if(($Mask[$i] -eq $IPBin[$i]) -and  $Mask[$i] -eq "1") {
                $net = $net + "1"
                }
            else{$net = $net + "0"}
            }

        #将子网的二进制串转换为十进制并用【.】分割成为IP地址串
        $SubNet = -join ([Convert]::ToInt32($net.Substring(0,8),2),".",[Convert]::ToInt32($net.Substring(8,8),2),".",[Convert]::ToInt32($net.Substring(16,8),2),".",[Convert]::ToInt32($net.Substring(24,8),2))
        #将子网掩码的二进制串转换为十进制并用【.】分割
        $SubNetMask = -join ([Convert]::ToInt32($mask.Substring(0,8),2),".",[Convert]::ToInt32($mask.Substring(8,8),2),".",[Convert]::ToInt32($mask.Substring(16,8),2),".",[Convert]::ToInt32($mask.Substring(24,8),2))
        [void]$sh.Cells.Item($l,1).Activate()
        $sh.Cells.Item($l,9).value() = $SubNet
        $sh.Cells.Item($l,10).value() = $SubNetMask
        #计算第一个可用IP
        # $firstIP = -join ([Convert]::ToInt32($net.Substring(0,8),2),".",[Convert]::ToInt32($net.Substring(8,8),2),".",[Convert]::ToInt32($net.Substring(16,8),2),".",$([Convert]::ToInt32($net.Substring(24,8),2) + 1))
        #计算最后一个可用IP
        # $part1 = $IPBin.Substring(0,$length)
        # for($i = $part1.Length;$i -lt 32;$i++){$part1 = $part1 + "1"}
        # $lastIP = -join ([Convert]::ToInt32($part1.Substring(0,8),2),".",[Convert]::ToInt32($part1.Substring(8,8),2),".",[Convert]::ToInt32($part1.Substring(16,8),2),".",$([Convert]::ToInt32($part1.Substring(24,8),2) - 1))
        # $broad = -join ([Convert]::ToInt32($part1.Substring(0,8),2),".",[Convert]::ToInt32($part1.Substring(8,8),2),".",[Convert]::ToInt32($part1.Substring(16,8),2),".",[Convert]::ToInt32($part1.Substring(24,8),2))
        # $sh.Cells.Item($l,11).value() = $firstIP
        # $sh.Cells.Item($l,12).value() = $lastIP
        #$sh.Cells.Item($l,13).value() = $broad
        # switch($length){
        #     "20" {$sh.Cells.Item($l,13).value() = 4096;break}
        #     "21" {$sh.Cells.Item($l,13).value() = 2048;break}
        #     "22" {$sh.Cells.Item($l,13).value() = 1024;break}
        #     "23" {$sh.Cells.Item($l,13).value() = 512;break}
        #     "24" {$sh.Cells.Item($l,13).value() = 256;break}
        #     "25" {$sh.Cells.Item($l,13).value() = 128;break}
        #     "26" {$sh.Cells.Item($l,13).value() = 64;break}
        #     "27" {$sh.Cells.Item($l,13).value() = 32;break}
        #     "28" {$sh.Cells.Item($l,13).value() = 16;break}
        #     "29" {$sh.Cells.Item($l,13).value() = 8;break}
        #     "30" {$sh.Cells.Item($l,13).value() = 4;break}
        # }
    }
}

$sh.Range('i1:j1').entirecolumn.autofit()
$xl.Worksheets.Item(1).Visible = $false
[void]$sh.Columns.Range("H:H").Delete()
$sh.range('a1:b1').EntireColumn.Hidden = $true
#[void]$sh.Columns.Range("A:B").Delete()
[void]$sh.Range('C2').Activate()

#########################

$pg = (gci C:\Out\IP\YZ_VDPortgroup*.csv | select -Last 1).fullname
[void]$xl.Workbooks.Open($pg)
$xl.workbooks.item(2).worksheets.item(1).copy([system.type]::missing,$sh)
$xl.Workbooks.Item(2).Close()


$xl.workbooks.item(1).SaveAs($rslt,51)
$xl.Workbooks.Close()
$xl.Quit()
rv xl
ac $home\TimeCalc.txt "$(Get-Date), Firewall-VLAN finished"


# $from = "豆腐坊 <collect-data@monitor.com>"
# $To = "DL VGC-Support <DL.DL-VGC-Support@t-systems.com>"
# #$Cc = "xiaohong.xu@t-systems.com",""
# $body = @"

# 感谢网络组提供的原始数据，本附件列出了所有防火墙上的VLAN及IP地址范围信息，每周一发一次

# 大家可以通过VLAN号进行匹配，希望能为同事们的查询带来方便。

# 祝大家工作愉快！
# "@

# Send-MailMessage -To $To -Subject "YiZhuang Firewall VLAN and IP Range - CW$(get-date -u %W)" -SmtpServer 6.86.3.12 -From $from -Attachments $rslt -Body $body -Encoding UTF8
# #Send-MailMessage -To "michael.gao@t-systems.com" -Subject "Merge test" -SmtpServer 6.86.3.12 -From $from -Attachments $rslt -Body $body -Encoding utf8

# gci C:\Out\YiZhuang-FireWall-VLAN_*.xlsx | ? creationtime -lt (get-date).adddays(-5) | mv -dest C:\Out\Arc\