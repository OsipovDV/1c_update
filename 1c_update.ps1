$Global:SchRegDen = $null

##########################################################################################################################################
#
#  ‘ункци€ Block1CDB 1 параметр строка подключени€ к базе в формате "ra-1c02:1641\some_db" и второй параметр пароль на подключение к базе 1с
#  ¬Ќ»ћјЌ»≈ - используетс€ глобальный параметр $global:SchReg дл€ определени€ была ли включена блокировка регламентных заданий до запуска
#

Function Block1CDB([string] $1c_conn_str,[string] $db_code)
{
    ##########################################################################################################################################
    #
    #  –азбираем строку подключени€ на им€ сервера, порт и им€ базы
    #

    
    $servername=$1c_conn_str.Substring(0,$1c_conn_str.IndexOf("\"))
    $dbname=$1c_conn_str.Substring($1c_conn_str.IndexOf("\")+1)
    
    if ($1c_conn_str.Contains(":") -eq $true)
    {
    $serverport=$servername.Substring($servername.IndexOf(":")+1)
    $servername=$servername.Substring(0,$servername.IndexOf(":"))
    } 
    else
    {
    $serverport="1541"
    }


    ##########################################################################################################################################
    #
    #  ѕодключаемс€ к кластеру 1с
    #

    $com_1c = New-Object -com v83.COMConnector
    $server_agent = $com_1c.ConnectAgent($servername+":"+([int32] $serverport-1).ToString())
    $clusters = $server_agent.GetClusters()
    foreach ($cluster in $clusters)
    {
        $server_agent.Authenticate($cluster,"","")
        $workingprocesses = $server_agent.GetWorkingProcesses($cluster)
        foreach ($workingprocess in $workingprocesses)
        {
            $conn_2wp = $com_1c.ConnectWorkingProcess("tcp://" + $workingprocess.HostName + ":" + $workingprocess.MainPort)
            $ibs = $conn_2wp.getinfobases()
            foreach($ib in $ibs)
            {

    ##########################################################################################################################################
    #
    #  ѕодключаемс€ к базе 1с, запрещаем подключение пользователей, блокируем регламентные задани€
    #
                 if ($ib.Name -eq $dbname)
                 {
                    if ($global:SchRegDen -eq $null) { $global:SchRegDen = [bool] $ib.ScheduledJobsDenied}
                    $ib.ConnectDenied = $true
                    $ib.PermissionCode = $db_code
                    $ib.ScheduledJobsDenied = $true
                    $conn_2wp.UpdateInfoBase($ib)
                 

    ##########################################################################################################################################
    #
    #  ¬ыгон€ем пользователей
    #
                     $user_connections = $conn_2wp.GetInfoBaseConnections($ib)
                        foreach( $uc in $user_connections)
                        {
                            if ($uc.AppID -ne "ComConsole")
                            {
                                $conn_2wp.Disconnect($uc)
                            }
                        }
                 }
            }
        }
    }
}


##########################################################################################################################################
#
#  ‘ункци€ UnBlock1CDB 1 параметр строка подключени€ к базе в формате "ra-1c02:1641\some_db"
#  ¬Ќ»ћјЌ»≈ - используетс€ глобальный параметр $global:SchReg дл€ определени€ была ли включена блокировка регламентных заданий до запуска
#

Function UnBlock1CDB([string] $1c_conn_str)
{

    ##########################################################################################################################################
    #
    #  –азбираем строку подключени€ на им€ сервера, порт и им€ базы
    #

    $servername=$1c_conn_str.Substring(0,$1c_conn_str.IndexOf("\"))
    $dbname=$1c_conn_str.Substring($1c_conn_str.IndexOf("\")+1)

    if ($1c_conn_str.Contains(":") -eq $true)
    {
    $serverport=$servername.Substring($servername.IndexOf(":")+1)
    $servername=$servername.Substring(0,$servername.IndexOf(":"))
    } 
    else
    {
    $serverport="1541"
    }

    ##########################################################################################################################################
    #
    #  ѕодключаемс€ к кластеру 1с
    #

    $com_1c = New-Object -com v83.COMConnector
    $server_agent = $com_1c.ConnectAgent($servername+":"+([int32] $serverport-1).ToString())
    $clusters = $server_agent.GetClusters()
    foreach ($cluster in $clusters)
    {
        $server_agent.Authenticate($cluster,"","")
        $workingprocesses = $server_agent.GetWorkingProcesses($cluster)
        foreach ($workingprocess in $workingprocesses)
        {
            $conn_2wp = $com_1c.ConnectWorkingProcess("tcp://" + $workingprocess.HostName + ":" + $workingprocess.MainPort)
            $ibs = $conn_2wp.getinfobases()
            foreach($ib in $ibs)
            {

    ##########################################################################################################################################
    #
    #  ѕодключаемс€ к базе 1с, запрещаем подключение пользователей, блокируем регламентные задани€
    #

                if ($ib.Name -eq $dbname)
                {
                    $ib.ScheduledJobsDenied = $Global:SchRegDen
                    $ib.ConnectDenied = $false
                    $ib.PermissionCode = ""
                    $conn_2wp.UpdateInfoBase($ib)

    ##########################################################################################################################################
    #
    #  ¬ыгон€ем пользователей
    #
                     $user_connections = $conn_2wp.GetInfoBaseConnections($ib)
                        foreach( $uc in $user_connections)
                        {
                            if ($uc.AppID -ne "ComConsole")
                            {
                                $conn_2wp.Disconnect($uc)
                            }
                        }
                 }
            }
        }
    }
}


## ƒл€ логировани€ в Event Log надо выполнить:
## New-EventLog ЦLogName Application ЦSource У1C Update scriptФ
########################################################################################################################################################
##
##
## ќпредел€ем константы
########################################################################################################################################################
$1c_exec_path = '"c:\Program Files (x86)\1cv8\8.3.6.2299\bin\1cv8.exe"'
$jobfile_path = "\\ra-fs04\1c_config_update$"
$updatelog_path = "\\ra-fs04\1c_config_update$\" + (Get-Date).ToShortDateString() + ".log"
$SMTPServer = "mx2.rusalco.com"
$1C_BlockCode = "L0CK3D"
[bool] $global:SchReg = $False
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
        continue
    }
    
    
    Send-MailMessage -SmtpServer  $SMTPServer -Encoding ([System.Text.Encoding]::UTF8) -From "1C_Update_script@roust.com" -To $RECIPIENTTO -Subject "ќбновление 1— базы $1CDBPATH запущено" -Body "ќбновл€ем базу $1CDBPATH файлом конфигурации $1CUPDATEPATH после $1CUPDATEDATETIME динамически:$IsDynamic"

    ## Ѕлокируем базу дл€ пользователей
    ########################################################################################################################################################
    ##

    if ($IsDynamic -eq $False) ## ≈сли обновл€ем динамически, пользователей выгон€ть не надо
    {
    ##  “аймаут на вс€кий пожарный случай, если накос€чили с файлом описанием обновлени€. 5 минут на отмену, достаточно переименовать\удалить файл задание
        Start-Sleep -s 300
        if ( (Test-Path -path $job.FullName)  -eq $false ) 
        {
            Send-MailMessage -SmtpServer  $SMTPServer -Encoding ([System.Text.Encoding]::UTF8) -From "1C_Update_script@roust.com" -To $RECIPIENTTO -Subject "ќбновление 1— базы $1CDBPATH ќ“ћ≈Ќ≈Ќќ" -Body "‘айл-задание был переименован или удален, обновление отменено."
            continue
        }
        Write-EventLog ЦLogName Application ЦSource У1C Update scriptФ ЦEntryType Information ЦEventID 100 ЦMessage УЅлокируем пользователей в базе $1CDBPATH.Ф
        $Global:SchRegDen = $null
        Block1CDB $1CDBPATH $1C_BlockCode

    }
    else
    {
        Write-EventLog ЦLogName Application ЦSource У1C Update scriptФ ЦEntryType Information ЦEventID 100 ЦMessage Уќбновление базы $1CDBPATH объ€влено динамическим, пользователей не выгон€ем.Ф
    }
 

    ## «апускаем обновление базы 1с (опционально, рестартим сервис 1с)
    ########################################################################################################################################################
    #"c:\Program Files (x86)\1cv8\8.3.6.2299\bin\1cv8.exe" config /S "RA-1c02:1641\UPP_GST_D_Osipov"  /UpdateCfg "d:\1Cv8-2015-12-03-14-59_10.cf" /UpdateDBCfg /Out"D:\LOG1C\log.txt"

    Write-EventLog ЦLogName Application ЦSource У1C Update scriptФ ЦEntryType Information ЦEventID 200 ЦMessage У«апускаем обновление базы $1CDBPATH файлом конфигурации $1CUPDATEPATH. ѕуть до лог-файла обновлени€ $updatelog_pathФ

    $args =  " config /S " + $1CDBPATH + ' /UpdateCfg "' + $1CUPDATEPATH + '" /UpdateDBCfg  /UC' + $1C_BlockCode + " /Out " + $updatelog_path 

    $starttime = Get-Date
    Start-Process $1c_exec_path -argumentlist $args -wait
    $updatetime = (Get-Date) - $starttime
    
    Write-EventLog ЦLogName Application ЦSource У1C Update scriptФ ЦEntryType Information ЦEventID 201 ЦMessage Уќбновление базы $1CDBPATH завершено за $updatetimeФ


    ## –азрешаем доступ в базу дл€ пользователей
    ########################################################################################################################################################
    ##

    if ($IsDynamic -eq $False)
    {
        Write-EventLog ЦLogName Application ЦSource У1C Update scriptФ ЦEntryType Information ЦEventID 101 ЦMessage У–азрешаем пользовател€м работу с базой $1CDBPATH.Ф
        UnBlock1CDB $1CDBPATH
    }
    
    Send-MailMessage -SmtpServer $SMTPServer -Encoding ([System.Text.Encoding]::UTF8) -From "1C_Update_script@roust.com" -To $RECIPIENTTO -Subject "ќбновление 1— базы $1CDBPATH завершено" -Body "ќбновление базы $1CDBPATH завершено, во вложении отчет о обновлении." -Attachments "$updatelog_path"
    
    If (Test-Path ($job.FullName + ".done"))
    { 
        Remove-Item -path ($job.FullName + ".done") -force
    }
    Rename-Item -path $job -newname ($job.FullName + ".done") -force
 }
}

 
