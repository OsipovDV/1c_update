## ��� ����������� � Event Log ���� ���������:
## New-EventLog �LogName Application �Source �1C Update script�
########################################################################################################################################################
##
##
## ���������� ���������
########################################################################################################################################################
$1c_exec_path = '"c:\Program Files (x86)\1cv8\8.3.6.2299\bin\1cv8.exe"'
$jobfile_path = "\\ra-fs04\1c_config_update$"
$updatelog_path = "D:\LOG1C\test.log"
$SMTPServer = "mx2.rusalco.com"
#


## ���� �����
########################################################################################################################################################
$jobfiles = Get-ChildItem ($jobfile_path + "\*.upd")


## ��������� �� �������
########################################################################################################################################################
foreach ($job in $jobfiles)
{

    ## ������ ����-������
    ########################################################################################################################################################
    # ���� �� ����
    # ���� �� ����� ����������
    # ���� � ����� ����������
    # ���� ����� ������
    # "�����������" ��� ���

    $content=Get-Content $job
    if ($content.Count -ge 4)
    {
    
        $1CDBPATH=$content[0]
        $1CUPDATEPATH=$content[1]
        $1CUPDATEDATETIME=$content[2]
        $RECIPIENTTO=$content[3]
        if ($content[4] -eq "�����������") { $IsDynamic=$True } else { $IsDynamic=$False}
    
    
    if ((Get-Date) -lt (Get-Date($1CUPDATEDATETIME)) )
    { 
        ## ����� ���������� ��� �� �������, ��������� � ����������
        break
    }
    
    
    Send-MailMessage -SmtpServer  $SMTPServer -Encoding ([System.Text.Encoding]::UTF8) -From "1C_Update_script@roust.com" -To $RECIPIENTTO -Subject "���������� 1� ���� $1CDBPATH ��������" -Body "��������� ���� $1CDBPATH ������ ������������ $1CUPDATEPATH ����� $1CUPDATEDATETIME �����������:$IsDynamic"

    ## ��������� ���� ��� �������������
    ########################################################################################################################################################
    ## "c:\Program Files (x86)\1cv8\8.3.6.2299\bin\1cv8.exe" enterprise /S "RA-1c02:1641\UPP_GST_D_Osipov" /C"����������������������������" /DisableStartupMessages /AU-

    if ($IsDynamic -eq $False) ## ���� ��������� �����������, ������������� �������� �� ����
    {
        Write-EventLog �LogName Application �Source �1C Update script� �EntryType Information �EventID 100 �Message ���������� ������������� � ���� $1CDBPATH.�
        $args =  " enterprise /S " + $1CDBPATH + ' /C"����������������������������"'  + " /DisableStartupMessages /AU-" 
               
        $Process = Start-Process $1c_exec_path -argumentlist $args -passthru

        Start-Sleep -s 60
        Stop-Process -Id $Process.Id
    }
    else
    {
        Write-EventLog �LogName Application �Source �1C Update script� �EntryType Information �EventID 100 �Message ����������� ���� $1CDBPATH ��������� ������������, ������������� �� ��������.�
    }
 

    ## ��������� ���������� ���� 1� (�����������, ��������� ������ 1�)
    ########################################################################################################################################################
    #"c:\Program Files (x86)\1cv8\8.3.6.2299\bin\1cv8.exe" config /S "RA-1c02:1641\UPP_GST_D_Osipov"  /UpdateCfg "d:\1Cv8-2015-12-03-14-59_10.cf" /UpdateDBCfg /Out"D:\LOG1C\log.txt"
    Write-EventLog �LogName Application �Source �1C Update script� �EntryType Information �EventID 200 �Message ���������� ���������� ���� $1CDBPATH ������ ������������ $1CUPDATEPATH. ���� �� ���-����� ���������� $updatelog_path�

    $args =  " config /S $1CDBPATH  /UpdateCfg $1CUPDATEPATH /UpdateDBCfg  /UC������������� /Out D:\LOG1C\test.log"

    $starttime = Get-Date
    Start-Process $1c_exec_path -argumentlist $args -wait
    $updatetime = (Get-Date) - $starttime
    
    Write-EventLog �LogName Application �Source �1C Update script� �EntryType Information �EventID 201 �Message ����������� ���� $1CDBPATH ��������� �� $updatetime�


    ## ��������� ������ � ���� ��� �������������
    ########################################################################################################################################################
    ## c:\Program Files (x86)\1cv8\8.3.6.2299\bin\1cv8.exe" enterprise /S "RA-1c02:1641\UPP_GST_D_Osipov" /C"����������������������������" /UC"�������������"

    if ($IsDynamic -eq $False)
    {
        Write-EventLog �LogName Application �Source �1C Update script� �EntryType Information �EventID 101 �Message ���������� ������������� ������ � ����� $1CDBPATH.�
        $args =  " enterprise /S " + $1CDBPATH + ' /C"����������������������������" /UC"�������������" '  + " /DisableStartupMessages /AU-"  
        $Process = Start-Process $1c_exec_path -argumentlist $args -passthru
        Start-Sleep -s 120
        Stop-Process -Id $Process.Id
    }
    
    Send-MailMessage -SmtpServer $SMTPServer -Encoding ([System.Text.Encoding]::UTF8) -From "1C_Update_script@roust.com" -To $RECIPIENTTO -Subject "���������� 1� ���� $1CDBPATH ���������" -Body "���������� ���� $1CDBPATH ���������, �� �������� ����� � ����������." -Attachments "$updatelog_path"
    
    If (Test-Path ($job.FullName + ".done"))
    { 
        Remove-Item -path ($job.FullName + ".done") -force
    }
    Rename-Item -path $job -newname ($job.FullName + ".done") -force
 }}

 