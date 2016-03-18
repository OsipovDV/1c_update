$Global:SchRegDen = $null

##########################################################################################################################################
#
#  ������� Block1CDB 1 �������� ������ ����������� � ���� � ������� "ra-1c02:1641\some_db" � ������ �������� ������ �� ����������� � ���� 1�
#  �������� - ������������ ���������� �������� $global:SchReg ��� ����������� ���� �� �������� ���������� ������������ ������� �� �������
#

Function Block1CDB([string] $1c_conn_str,[string] $db_code)
{
    ##########################################################################################################################################
    #
    #  ��������� ������ ����������� �� ��� �������, ���� � ��� ����
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
    #  ������������ � �������� 1�
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
    #  ������������ � ���� 1�, ��������� ����������� �������������, ��������� ������������ �������
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
    #  �������� �������������
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
#  ������� UnBlock1CDB 1 �������� ������ ����������� � ���� � ������� "ra-1c02:1641\some_db"
#  �������� - ������������ ���������� �������� $global:SchReg ��� ����������� ���� �� �������� ���������� ������������ ������� �� �������
#

Function UnBlock1CDB([string] $1c_conn_str)
{

    ##########################################################################################################################################
    #
    #  ��������� ������ ����������� �� ��� �������, ���� � ��� ����
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
    #  ������������ � �������� 1�
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
    #  ������������ � ���� 1�, ��������� ����������� �������������, ��������� ������������ �������
    #

                if ($ib.Name -eq $dbname)
                {
                    $ib.ScheduledJobsDenied = $Global:SchRegDen
                    $ib.ConnectDenied = $false
                    $ib.PermissionCode = ""
                    $conn_2wp.UpdateInfoBase($ib)

    ##########################################################################################################################################
    #
    #  �������� �������������
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


## ��� ����������� � Event Log ���� ���������:
## New-EventLog �LogName Application �Source �1C Update script�
########################################################################################################################################################
##
##
## ���������� ���������
########################################################################################################################################################
$1c_exec_path = '"c:\Program Files (x86)\1cv8\8.3.6.2299\bin\1cv8.exe"'
$jobfile_path = "\\ra-fs04\1c_config_update$"
$updatelog_path = "\\ra-fs04\1c_config_update$\" + (Get-Date).ToShortDateString() + ".log"
$SMTPServer = "mx2.rusalco.com"
$1C_BlockCode = "L0CK3D"
[bool] $global:SchReg = $False
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
        continue
    }
    
    
    Send-MailMessage -SmtpServer  $SMTPServer -Encoding ([System.Text.Encoding]::UTF8) -From "1C_Update_script@roust.com" -To $RECIPIENTTO -Subject "���������� 1� ���� $1CDBPATH ��������" -Body "��������� ���� $1CDBPATH ������ ������������ $1CUPDATEPATH ����� $1CUPDATEDATETIME �����������:$IsDynamic"

    ## ��������� ���� ��� �������������
    ########################################################################################################################################################
    ##

    if ($IsDynamic -eq $False) ## ���� ��������� �����������, ������������� �������� �� ����
    {
    ##  ������� �� ������ �������� ������, ���� ���������� � ������ ��������� ����������. 5 ����� �� ������, ���������� �������������\������� ���� �������
        Start-Sleep -s 300
        if ( (Test-Path -path $job.FullName)  -eq $false ) 
        {
            Send-MailMessage -SmtpServer  $SMTPServer -Encoding ([System.Text.Encoding]::UTF8) -From "1C_Update_script@roust.com" -To $RECIPIENTTO -Subject "���������� 1� ���� $1CDBPATH ��������" -Body "����-������� ��� ������������ ��� ������, ���������� ��������."
            continue
        }
        Write-EventLog �LogName Application �Source �1C Update script� �EntryType Information �EventID 100 �Message ���������� ������������� � ���� $1CDBPATH.�
        $Global:SchRegDen = $null
        Block1CDB $1CDBPATH $1C_BlockCode

    }
    else
    {
        Write-EventLog �LogName Application �Source �1C Update script� �EntryType Information �EventID 100 �Message ����������� ���� $1CDBPATH ��������� ������������, ������������� �� ��������.�
    }
 

    ## ��������� ���������� ���� 1� (�����������, ��������� ������ 1�)
    ########################################################################################################################################################
    #"c:\Program Files (x86)\1cv8\8.3.6.2299\bin\1cv8.exe" config /S "RA-1c02:1641\UPP_GST_D_Osipov"  /UpdateCfg "d:\1Cv8-2015-12-03-14-59_10.cf" /UpdateDBCfg /Out"D:\LOG1C\log.txt"

    Write-EventLog �LogName Application �Source �1C Update script� �EntryType Information �EventID 200 �Message ���������� ���������� ���� $1CDBPATH ������ ������������ $1CUPDATEPATH. ���� �� ���-����� ���������� $updatelog_path�

    $args =  " config /S " + $1CDBPATH + ' /UpdateCfg "' + $1CUPDATEPATH + '" /UpdateDBCfg  /UC' + $1C_BlockCode + " /Out " + $updatelog_path 

    $starttime = Get-Date
    Start-Process $1c_exec_path -argumentlist $args -wait
    $updatetime = (Get-Date) - $starttime
    
    Write-EventLog �LogName Application �Source �1C Update script� �EntryType Information �EventID 201 �Message ����������� ���� $1CDBPATH ��������� �� $updatetime�


    ## ��������� ������ � ���� ��� �������������
    ########################################################################################################################################################
    ##

    if ($IsDynamic -eq $False)
    {
        Write-EventLog �LogName Application �Source �1C Update script� �EntryType Information �EventID 101 �Message ���������� ������������� ������ � ����� $1CDBPATH.�
        UnBlock1CDB $1CDBPATH
    }
    
    Send-MailMessage -SmtpServer $SMTPServer -Encoding ([System.Text.Encoding]::UTF8) -From "1C_Update_script@roust.com" -To $RECIPIENTTO -Subject "���������� 1� ���� $1CDBPATH ���������" -Body "���������� ���� $1CDBPATH ���������, �� �������� ����� � ����������." -Attachments "$updatelog_path"
    
    If (Test-Path ($job.FullName + ".done"))
    { 
        Remove-Item -path ($job.FullName + ".done") -force
    }
    Rename-Item -path $job -newname ($job.FullName + ".done") -force
 }
}

 
