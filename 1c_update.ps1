$Global:SchRegDen = $null

##########################################################################################################################################
#
#  Ôóíêöèÿ Block1CDB 1 ïàðàìåòð ñòðîêà ïîäêëþ÷åíèÿ ê áàçå â ôîðìàòå "srv-1c02:1641\some_db" è âòîðîé ïàðàìåòð ïàðîëü íà ïîäêëþ÷åíèå ê áàçå 1ñ
#  ÂÍÈÌÀÍÈÅ - èñïîëüçóåòñÿ ãëîáàëüíûé ïàðàìåòð $global:SchReg äëÿ îïðåäåëåíèÿ áûëà ëè âêëþ÷åíà áëîêèðîâêà ðåãëàìåíòíûõ çàäàíèé äî çàïóñêà
#

Function Block1CDB([string] $1c_conn_str,[string] $db_code)
{
    ##########################################################################################################################################
    #
    #  Ðàçáèðàåì ñòðîêó ïîäêëþ÷åíèÿ íà èìÿ ñåðâåðà, ïîðò è èìÿ áàçû
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
    #  Ïîäêëþ÷àåìñÿ ê êëàñòåðó 1ñ
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
    #  Ïîäêëþ÷àåìñÿ ê áàçå 1ñ, çàïðåùàåì ïîäêëþ÷åíèå ïîëüçîâàòåëåé, áëîêèðóåì ðåãëàìåíòíûå çàäàíèÿ
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
    #  Âûãîíÿåì ïîëüçîâàòåëåé
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
#  Ôóíêöèÿ UnBlock1CDB 1 ïàðàìåòð ñòðîêà ïîäêëþ÷åíèÿ ê áàçå â ôîðìàòå "srv-1c02:1641\some_db"
#  ÂÍÈÌÀÍÈÅ - èñïîëüçóåòñÿ ãëîáàëüíûé ïàðàìåòð $global:SchReg äëÿ îïðåäåëåíèÿ áûëà ëè âêëþ÷åíà áëîêèðîâêà ðåãëàìåíòíûõ çàäàíèé äî çàïóñêà
#

Function UnBlock1CDB([string] $1c_conn_str)
{

    ##########################################################################################################################################
    #
    #  Ðàçáèðàåì ñòðîêó ïîäêëþ÷åíèÿ íà èìÿ ñåðâåðà, ïîðò è èìÿ áàçû
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
    #  Ïîäêëþ÷àåìñÿ ê êëàñòåðó 1ñ
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
    #  Ïîäêëþ÷àåìñÿ ê áàçå 1ñ, çàïðåùàåì ïîäêëþ÷åíèå ïîëüçîâàòåëåé, áëîêèðóåì ðåãëàìåíòíûå çàäàíèÿ
    #

                if ($ib.Name -eq $dbname)
                {
                    $ib.ScheduledJobsDenied = $Global:SchRegDen
                    $ib.ConnectDenied = $false
                    $ib.PermissionCode = ""
                    $conn_2wp.UpdateInfoBase($ib)

    ##########################################################################################################################################
    #
    #  Âûãîíÿåì ïîëüçîâàòåëåé
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


## Äëÿ ëîãèðîâàíèÿ â Event Log íàäî âûïîëíèòü:
## New-EventLog –LogName Application –Source “1C Update script”
########################################################################################################################################################
##
##
## Îïðåäåëÿåì êîíñòàíòû
########################################################################################################################################################
$1c_exec_path = '"c:\Program Files (x86)\1cv8\8.3.6.2299\bin\1cv8.exe"'
$jobfile_path = "\\srv-fs04\1c_config_update$"
$updatelog_path = "\\srv-fs04\1c_config_update$\" + (Get-Date).ToShortDateString() + ".log"
$SMTPServer = "mx2.some.local"
$1C_BlockCode = "1CBLOCKED"
[bool] $global:SchReg = $False
#


## Èùåì ôàéëû
########################################################################################################################################################
$jobfiles = Get-ChildItem ($jobfile_path + "\*.upd")


## Îáíîâëÿåì ïî î÷åðåäè
########################################################################################################################################################
foreach ($job in $jobfiles)
{

    ## Ïàðñèì ôàéë-çàäà÷ó
    ########################################################################################################################################################
    # ïóòü äî áàçû
    # ïóòü äî ôàéëà îáíîâëåíèÿ
    # äàòà è âðåìÿ îáíîâëåíèÿ
    # êîìó ñëàòü îò÷åòû
    # "äèíàìè÷åñêè" èëè íåò

    $content=Get-Content $job
    if ($content.Count -ge 4)
    {
    
        $1CDBPATH=$content[0]
        $1CUPDATEPATH=$content[1]
        $1CUPDATEDATETIME=$content[2]
        $RECIPIENTTO=$content[3]
        if ($content[4] -eq "äèíàìè÷åñêè") { $IsDynamic=$True } else { $IsDynamic=$False}
    
    
    if ((Get-Date) -lt (Get-Date($1CUPDATEDATETIME)) )
    { 
        ## Âðåìÿ îáíîâëåíèÿ åùå íå ïîäîøëî, ïåðåõîäèì ê ñëåäóþùåìó
        continue
    }
    
    
    Send-MailMessage -SmtpServer  $SMTPServer -Encoding ([System.Text.Encoding]::UTF8) -From "1C_Update_script@some.local" -To $RECIPIENTTO -Subject "Îáíîâëåíèå 1Ñ áàçû $1CDBPATH çàïóùåíî" -Body "Îáíîâëÿåì áàçó $1CDBPATH ôàéëîì êîíôèãóðàöèè $1CUPDATEPATH ïîñëå $1CUPDATEDATETIME äèíàìè÷åñêè:$IsDynamic"

    ## Áëîêèðóåì áàçó äëÿ ïîëüçîâàòåëåé
    ########################################################################################################################################################
    ##

    if ($IsDynamic -eq $False) ## Åñëè îáíîâëÿåì äèíàìè÷åñêè, ïîëüçîâàòåëåé âûãîíÿòü íå íàäî
    {
    ##  Òàéìàóò íà âñÿêèé ïîæàðíûé ñëó÷àé, åñëè íàêîñÿ÷èëè ñ ôàéëîì îïèñàíèåì îáíîâëåíèÿ. 5 ìèíóò íà îòìåíó, äîñòàòî÷íî ïåðåèìåíîâàòü\óäàëèòü ôàéë çàäàíèå
        Start-Sleep -s 300
        if ( (Test-Path -path $job.FullName)  -eq $false ) 
        {
            Send-MailMessage -SmtpServer  $SMTPServer -Encoding ([System.Text.Encoding]::UTF8) -From "1C_Update_script@some.local" -To $RECIPIENTTO -Subject "Îáíîâëåíèå 1Ñ áàçû $1CDBPATH ÎÒÌÅÍÅÍÎ" -Body "Ôàéë-çàäàíèå áûë ïåðåèìåíîâàí èëè óäàëåí, îáíîâëåíèå îòìåíåíî."
            continue
        }
        Write-EventLog –LogName Application –Source “1C Update script” –EntryType Information –EventID 100 –Message “Áëîêèðóåì ïîëüçîâàòåëåé â áàçå $1CDBPATH.”
        $Global:SchRegDen = $null
        Block1CDB $1CDBPATH $1C_BlockCode

    }
    else
    {
        Write-EventLog –LogName Application –Source “1C Update script” –EntryType Information –EventID 100 –Message “Îáíîâëåíèå áàçû $1CDBPATH îáúÿâëåíî äèíàìè÷åñêèì, ïîëüçîâàòåëåé íå âûãîíÿåì.”
    }
 

    ## Çàïóñêàåì îáíîâëåíèå áàçû 1ñ (îïöèîíàëüíî, ðåñòàðòèì ñåðâèñ 1ñ)
    ########################################################################################################################################################
    #"c:\Program Files (x86)\1cv8\8.3.6.2299\bin\1cv8.exe" config /S "srv-1c02:1641\UPP_GST_D_Osipov"  /UpdateCfg "d:\1Cv8-2015-12-03-14-59_10.cf" /UpdateDBCfg /Out"D:\LOG1C\log.txt"

    Write-EventLog –LogName Application –Source “1C Update script” –EntryType Information –EventID 200 –Message “Çàïóñêàåì îáíîâëåíèå áàçû $1CDBPATH ôàéëîì êîíôèãóðàöèè $1CUPDATEPATH. Ïóòü äî ëîã-ôàéëà îáíîâëåíèÿ $updatelog_path”

    $args =  " config /S " + $1CDBPATH + ' /UpdateCfg "' + $1CUPDATEPATH + '" /UpdateDBCfg  /UC' + $1C_BlockCode + " /Out " + $updatelog_path 

    $starttime = Get-Date
    Start-Process $1c_exec_path -argumentlist $args -wait
    $updatetime = (Get-Date) - $starttime
    
    Write-EventLog –LogName Application –Source “1C Update script” –EntryType Information –EventID 201 –Message “Îáíîâëåíèå áàçû $1CDBPATH çàâåðøåíî çà $updatetime”


    ## Ðàçðåøàåì äîñòóï â áàçó äëÿ ïîëüçîâàòåëåé
    ########################################################################################################################################################
    ##

    if ($IsDynamic -eq $False)
    {
        Write-EventLog –LogName Application –Source “1C Update script” –EntryType Information –EventID 101 –Message “Ðàçðåøàåì ïîëüçîâàòåëÿì ðàáîòó ñ áàçîé $1CDBPATH.”
        UnBlock1CDB $1CDBPATH
    }
    
    Send-MailMessage -SmtpServer $SMTPServer -Encoding ([System.Text.Encoding]::UTF8) -From "1C_Update_script@some.local" -To $RECIPIENTTO -Subject "Îáíîâëåíèå 1Ñ áàçû $1CDBPATH çàâåðøåíî" -Body "Îáíîâëåíèå áàçû $1CDBPATH çàâåðøåíî, âî âëîæåíèè îò÷åò î îáíîâëåíèè." -Attachments "$updatelog_path"
    
    If (Test-Path ($job.FullName + ".done"))
    { 
        Remove-Item -path ($job.FullName + ".done") -force
    }
    Rename-Item -path $job -newname ($job.FullName + ".done") -force
 }
}

 
