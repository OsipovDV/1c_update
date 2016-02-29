## ƒл€ логировани€ в Event Log надо выполнить:
## New-EventLog ЦLogName Application ЦSource У1C Update scriptФ
########################################################################################################################################################
##
##
## ќпредел€ем константы
########################################################################################################################################################
$1c_exec_path = '"c:\Program Files (x86)\1cv8\8.3.6.2299\bin\1cv8.exe"'
$jobfile_path = "\\ra-fs04\1c_config_update$"
$updatelog_path = "D:\LOG1C\test.log"
$SMTPServer = "mx2.rusalco.com"
#


## »щем файлы
########################################################################################################################################################
$jobfiles = Get-ChildItem ($jobfile_path + "\*.upd")


## ќбновл€ем по очереди
########################################################################################################################################################
foreach ($job in $jobfiles)
{

    ## ѕарсим файл-задачу
    ########################################################################################################################################################
    # путь до базы
    # путь до файла обновлени€
    # дата и врем€ обновлени€
    # кому слать отчеты
    # "динамически" или нет

    $content=Get-Content $job
    if ($content.Count -ge 4)
    {
    
        $1CDBPATH=$content[0]
        $1CUPDATEPATH=$content[1]
        $1CUPDATEDATETIME=$content[2]
        $RECIPIENTTO=$content[3]
        if ($content[4] -eq "динамически") { $IsDynamic=$True } else { $IsDynamic=$False}
    
    
    if ((Get-Date) -lt (Get-Date($1CUPDATEDATETIME)) )
    { 
        ## ¬рем€ обновлени€ еще не подошло, переходим к следующему
        break
    }
    
    
    Send-MailMessage -SmtpServer  $SMTPServer -Encoding ([System.Text.Encoding]::UTF8) -From "1C_Update_script@roust.com" -To $RECIPIENTTO -Subject "ќбновление 1— базы $1CDBPATH запущено" -Body "ќбновл€ем базу $1CDBPATH файлом конфигурации $1CUPDATEPATH после $1CUPDATEDATETIME динамически:$IsDynamic"

    ## Ѕлокируем базу дл€ пользователей
    ########################################################################################################################################################
    ## "c:\Program Files (x86)\1cv8\8.3.6.2299\bin\1cv8.exe" enterprise /S "RA-1c02:1641\UPP_GST_D_Osipov" /C"«авершить–аботуѕользователей" /DisableStartupMessages /AU-

    if ($IsDynamic -eq $False) ## ≈сли обновл€ем динамически, пользователей выгон€ть не надо
    {
        Write-EventLog ЦLogName Application ЦSource У1C Update scriptФ ЦEntryType Information ЦEventID 100 ЦMessage УЅлокируем пользователей в базе $1CDBPATH.Ф
        $args =  " enterprise /S " + $1CDBPATH + ' /C"«авершить–аботуѕользователей"'  + " /DisableStartupMessages /AU-" 
               
        $Process = Start-Process $1c_exec_path -argumentlist $args -passthru

        Start-Sleep -s 60
        Stop-Process -Id $Process.Id
    }
    else
    {
        Write-EventLog ЦLogName Application ЦSource У1C Update scriptФ ЦEntryType Information ЦEventID 100 ЦMessage Уќбновление базы $1CDBPATH объ€влено динамическим, пользователей не выгон€ем.Ф
    }
 

    ## «апускаем обновление базы 1с (опционально, рестартим сервис 1с)
    ########################################################################################################################################################
    #"c:\Program Files (x86)\1cv8\8.3.6.2299\bin\1cv8.exe" config /S "RA-1c02:1641\UPP_GST_D_Osipov"  /UpdateCfg "d:\1Cv8-2015-12-03-14-59_10.cf" /UpdateDBCfg /Out"D:\LOG1C\log.txt"
    Write-EventLog ЦLogName Application ЦSource У1C Update scriptФ ЦEntryType Information ЦEventID 200 ЦMessage У«апускаем обновление базы $1CDBPATH файлом конфигурации $1CUPDATEPATH. ѕуть до лог-файла обновлени€ $updatelog_pathФ

    $args =  " config /S $1CDBPATH  /UpdateCfg $1CUPDATEPATH /UpdateDBCfg  /UC од–азрешени€ /Out D:\LOG1C\test.log"

    $starttime = Get-Date
    Start-Process $1c_exec_path -argumentlist $args -wait
    $updatetime = (Get-Date) - $starttime
    
    Write-EventLog ЦLogName Application ЦSource У1C Update scriptФ ЦEntryType Information ЦEventID 201 ЦMessage Уќбновление базы $1CDBPATH завершено за $updatetimeФ


    ## –азрешаем доступ в базу дл€ пользователей
    ########################################################################################################################################################
    ## c:\Program Files (x86)\1cv8\8.3.6.2299\bin\1cv8.exe" enterprise /S "RA-1c02:1641\UPP_GST_D_Osipov" /C"–азрешить–аботуѕользователей" /UC" од–азрешени€"

    if ($IsDynamic -eq $False)
    {
        Write-EventLog ЦLogName Application ЦSource У1C Update scriptФ ЦEntryType Information ЦEventID 101 ЦMessage У–азрешаем пользовател€м работу с базой $1CDBPATH.Ф
        $args =  " enterprise /S " + $1CDBPATH + ' /C"–азрешить–аботуѕользователей" /UC" од–азрешени€" '  + " /DisableStartupMessages /AU-"  
        $Process = Start-Process $1c_exec_path -argumentlist $args -passthru
        Start-Sleep -s 120
        Stop-Process -Id $Process.Id
    }
    
    Send-MailMessage -SmtpServer $SMTPServer -Encoding ([System.Text.Encoding]::UTF8) -From "1C_Update_script@roust.com" -To $RECIPIENTTO -Subject "ќбновление 1— базы $1CDBPATH завершено" -Body "ќбновление базы $1CDBPATH завершено, во вложении отчет о обновлении." -Attachments "$updatelog_path"
    
    If (Test-Path ($job.FullName + ".done"))
    { 
        Remove-Item -path ($job.FullName + ".done") -force
    }
    Rename-Item -path $job -newname ($job.FullName + ".done") -force
 }}

 