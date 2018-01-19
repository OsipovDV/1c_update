$Global:SchRegDen = $null

##########################################################################################################################################
#
#  Функция Block1CDB 1 параметр строка подключения к базе в формате "srv-1c02:1641\some_db" и второй параметр пароль на подключение к базе 1с
#  ВНИМАНИЕ - используется глобальный параметр $global:SchReg для определения была ли включена блокировка регламентных заданий до запуска
#

Function Block1CDB([string] $1c_conn_str,[string] $db_code)
{
    ##########################################################################################################################################
    #
    #  Разбираем строку подключения на имя сервера, порт и имя базы
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
    #  Подключаемся к кластеру 1с
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
    #  Подключаемся к базе 1с, запрещаем подключение пользователей, блокируем регламентные задания
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
    #  Выгоняем пользователей
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
#  Функция UnBlock1CDB 1 параметр строка подключения к базе в формате "srv-1c02:1641\some_db"
#  ВНИМАНИЕ - используется глобальный параметр $global:SchReg для определения была ли включена блокировка регламентных заданий до запуска
#

Function UnBlock1CDB([string] $1c_conn_str)
{

    ##########################################################################################################################################
    #
    #  Разбираем строку подключения на имя сервера, порт и имя базы
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
    #  Подключаемся к кластеру 1с
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
    #  Подключаемся к базе 1с, запрещаем подключение пользователей, блокируем регламентные задания
    #

                if ($ib.Name -eq $dbname)
                {
                    $ib.ScheduledJobsDenied = $Global:SchRegDen
                    $ib.ConnectDenied = $false
                    $ib.PermissionCode = ""
                    $conn_2wp.UpdateInfoBase($ib)

    ##########################################################################################################################################
    #
    #  Выгоняем пользователей
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


## Для логирования в Event Log надо выполнить:
## New-EventLog –LogName Application –Source “1C Update script”
########################################################################################################################################################
##
##
## Определяем константы
########################################################################################################################################################
$1c_exec_path = '"c:\Program Files (x86)\1cv8\8.3.6.2299\bin\1cv8.exe"'
$jobfile_path = "\\srv-fs04\1c_config_update$"
$updatelog_path = "\\srv-fs04\1c_config_update$\" + (Get-Date).ToShortDateString() + ".log"
$SMTPServer = "mx2.some.local"
$1C_BlockCode = "1CBLOCKED"
[bool] $global:SchReg = $False
#


## Ищем файлы
########################################################################################################################################################
$jobfiles = Get-ChildItem ($jobfile_path + "\*.upd")


## Обновляем по очереди
########################################################################################################################################################
foreach ($job in $jobfiles)
{

    ## Парсим файл-задачу
    ########################################################################################################################################################
    # путь до базы
    # путь до файла обновления
    # дата и время обновления
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
        ## Время обновления еще не подошло, переходим к следующему
        continue
    }
    
    
    Send-MailMessage -SmtpServer  $SMTPServer -Encoding ([System.Text.Encoding]::UTF8) -From "1C_Update_script@some.local" -To $RECIPIENTTO -Subject "Обновление 1С базы $1CDBPATH запущено" -Body "Обновляем базу $1CDBPATH файлом конфигурации $1CUPDATEPATH после $1CUPDATEDATETIME динамически:$IsDynamic"

    ## Блокируем базу для пользователей
    ########################################################################################################################################################
    ##

    if ($IsDynamic -eq $False) ## Если обновляем динамически, пользователей выгонять не надо
    {
    ##  Таймаут на всякий пожарный случай, если накосячили с файлом описанием обновления. 5 минут на отмену, достаточно переименовать\удалить файл задание
        Start-Sleep -s 300
        if ( (Test-Path -path $job.FullName)  -eq $false ) 
        {
            Send-MailMessage -SmtpServer  $SMTPServer -Encoding ([System.Text.Encoding]::UTF8) -From "1C_Update_script@some.local" -To $RECIPIENTTO -Subject "Обновление 1С базы $1CDBPATH ОТМЕНЕНО" -Body "Файл-задание был переименован или удален, обновление отменено."
            continue
        }
        Write-EventLog –LogName Application –Source “1C Update script” –EntryType Information –EventID 100 –Message “Блокируем пользователей в базе $1CDBPATH.”
        $Global:SchRegDen = $null
        Block1CDB $1CDBPATH $1C_BlockCode

    }
    else
    {
        Write-EventLog –LogName Application –Source “1C Update script” –EntryType Information –EventID 100 –Message “Îáíîâëåíèå áàçû $1CDBPATH îáúÿâëåíî äèíàìè÷åñêèì, ïîëüçîâàòåëåé íå âûãîíÿåì.”
    }
 

    ## Запускаем обновление базы 1с (опционально, рестартим сервис 1с)
    ########################################################################################################################################################
    #"c:\Program Files (x86)\1cv8\8.3.6.2299\bin\1cv8.exe" config /S "srv-1c02:1641\UPP_GST_D_Osipov"  /UpdateCfg "d:\1Cv8-2015-12-03-14-59_10.cf" /UpdateDBCfg /Out"D:\LOG1C\log.txt"

    Write-EventLog –LogName Application –Source “1C Update script” –EntryType Information –EventID 200 –Message “Запускаем обновление базы $1CDBPATH файлом конфигурации $1CUPDATEPATH. Путь до лог-файла обновления $updatelog_path”

    $args =  " config /S " + $1CDBPATH + ' /UpdateCfg "' + $1CUPDATEPATH + '" /UpdateDBCfg  /UC' + $1C_BlockCode + " /Out " + $updatelog_path 

    $starttime = Get-Date
    Start-Process $1c_exec_path -argumentlist $args -wait
    $updatetime = (Get-Date) - $starttime
    
    Write-EventLog –LogName Application –Source “1C Update script” –EntryType Information –EventID 201 –Message “Обновление базы $1CDBPATH завершено за $updatetime”


    ## Разрешаем доступ в базу для пользователей
    ########################################################################################################################################################
    ##

    if ($IsDynamic -eq $False)
    {
        Write-EventLog –LogName Application –Source “1C Update script” –EntryType Information –EventID 101 –Message “Разрешаем пользователям работу с базой $1CDBPATH.”
        UnBlock1CDB $1CDBPATH
    }
    
    Send-MailMessage -SmtpServer $SMTPServer -Encoding ([System.Text.Encoding]::UTF8) -From "1C_Update_script@some.local" -To $RECIPIENTTO -Subject "Обновление 1С базы $1CDBPATH завершено" -Body "Обновление базы $1CDBPATH завершено, во вложении отчет о обновлении." -Attachments "$updatelog_path"
    
    If (Test-Path ($job.FullName + ".done"))
    { 
        Remove-Item -path ($job.FullName + ".done") -force
    }
    Rename-Item -path $job -newname ($job.FullName + ".done") -force
 }
}

 
