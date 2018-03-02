clear-host

[int] $MSCurRow =3#$args[0]
$MSVersionNum = $args[1]
$MSBuildNum = $args[2]
$TestItemName = $args[3]

#$BaseFolder = "C:\TCResults\PSResults\QATools_Logs\"
$BaseFolder = "\\hv-autoscripts\AutomationLogs\TestComplete\PS_Logs\QATools_Logs\"
$slaveIP= $(ipconfig | where {$_ -match 'IPv4.+\s(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})' } | out-null; $Matches[1])

$logfilename = $(((get-date).ToUniversalTime()).ToString("MMddyyyyhhmmss")) #Get-Date -UFormat "%Y_%m_%d_%h_%m_%s"

$date = Get-Date;$TodayFolder = $BaseFolder+$date.ToString("MM-dd-yyyy");New-Item -ItemType Directory -Path $TodayFolder -ErrorAction SilentlyContinue

$VerBldFolder = $TodayFolder+"\"+"V"+$MSVersionNum+"_"+"B"+$MSBuildNum+"_"+"IP"+$slaveIP;New-Item -ItemType directory -Path $VerBldFolder -ErrorAction SilentlyContinue

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$excelFilePath = $scriptPath -replace "PowerShellScripts", "OutPutForPortal"

$excel = New-Object -ComObject Excel.Application;$excel.DisplayAlerts = $false;$excel.Visible = $false   

$excelFiles_Master = Get-ChildItem -Path $excelFilePath'\QAToolsPortalData.xlsx'
$wb_TestAutomation = $excel.workbooks.open($excelfiles_Master.fullname)

$ws_Distributed = $wb_TestAutomation.Sheets.Item('Distributed')
$Distributed_Range = $ws_Distributed.UsedRange
$DistributedRows = $Distributed_Range.Rows.Count

$logfilename = $VerBldFolder+"\"+$logfilename+"_"+$MSVersionNum+"_"+$MSBuildNum+"_"+$slaveIP+"_"+$MSCurRow+"_"+$TestItemName+".txt"
            
$(
    for ($ip=1; $ip -le $DistributedRows+1;$ip++)
    {
        $slaveActualIP = $ws_Distributed.Cells.Item($ip,2).text 
        
        #$slaveIP= $(ipconfig | where {$_ -match 'IPv4.+\s(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})' } | out-null; $Matches[1])

        if ($slaveActualIP -ceq $slaveIP)
        {   
        clear-host
        
        Write-Host "Pushing the Results to Portal From Slave IP: "+$slaveIP+";  Version:"+$MSVersionNum+";  Build Number:"+$MSBuildNum
        Write-Host "###################################################################################################################"
        $scriptPath = split-path -parent $MyInvocation.MyCommand.Definition        
        
        $ws = $wb_TestAutomation.Sheets.Item('PortalData')
        $wsRange = $ws.UsedRange
        $rows = $wsRange.Rows.Count

        for ($i = $MSCurRow; $i -le $rows+1; $i++) 

         {                               
            $ScenarioRunStatus = $ws.cells.item($i,3).value()  

            if ($ScenarioRunStatus -ceq "Yes")
            {      
                $ScenarioResultFlag = ""                
                $Category = $ws.cells.item($i,2).value();$ScenarioName = $ws.cells.item($i,4).value();$PreviousScenarioName = $ScenarioName;                               
                $ScenarioResult = $ws.cells.item($i,5).value();$TestCaseName = $ws.cells.item($i,7).value();$TestCaseErrorText = $ws.cells.item($i,9).value();                
                $total_time = $ws.cells.item($i,10).value();$productname = $ws.cells.item($i,11).value();$versionname = $ws.cells.item($i,12).value();                
                $buildname = $ws.cells.item($i,13).value();$hostname = $ws.cells.item($i,14).value();$logs_location = $ws.cells.item($i,15).value();
                $DataType = $ws.cells.item($i,17).value();$test_StartTime = $ws.cells.item($i,18).value();$test_EndTime = $ws.cells.item($i,19).value();
                $Date1 = $test_StartTime;$Date2 = $test_EndTime;$TimeDiff = New-TimeSpan $Date1 $Date2;$Difference = "{0:g}" -f $TimeDiff;$total_time = $Difference;
                $PE_TestCase = $ScenarioName+"("+$TestCaseName+")"

                #Read Scenario Count
                $ss_Range = @();

                do {
                    $CurScenarioName = $ws.Cells.Item($i,4).value();
                    if ($CurScenarioName -eq $PreviousScenarioName)
                        {
                            $TCResult = $ws.cells.item($i,8).value();
                            if (($TCResult -eq "failed") -or ($TCResult -eq $null))
                                {$ScenarioResultFlag = "failed";  $screenshotpath = $ws.cells.item($i,16).value()}
                                #elseif ($TCResult -eq "passed"); {$SenarioResultFlag = $true};else{$SenarioResultFlag = $false}                           
                                $i=$i+1;$CurScenarioName = $ws.cells.item($i,4).value()
                        }
                    else{$i=$i+1}
                } until ( ($CurScenarioName -ne $PreviousScenarioName) -or  $i -eq $rows)
                Write-Host "I = "$i"; Current Scenario Name = "$CurScenarioName"; Previous Scenrio Name = "$PreviousScenarioName"; Scenario Flag = "$ScenarioResultFlag
                if ($ScenarioResultFlag -eq "failed"){$ScenarioResult = "failed"}
                else {$ScenarioResult = "passed"}                
                if ( $i -ne $rows){$i=$i-1;}

                [string]$postParams1 =  "{"                
                $postParams1 += "`n""test_case"":""" + $ScenarioName + """," ;$postParams1 += "`n""category"":""" + $Category + """,";
                $postParams1 += "`n""result"":""" +  $ScenarioResult  + """,";$postParams1 += "`n""error"":""" + $TestCaseErrorText + ""","; 
                $postParams1 += "`n""total_time"":""" + $total_time + """,";$postParams1 += "`n""product"":""" + $productname  + """,";
                $postParams1 += "`n""version"":""" + $versionname + """,";$postParams1 += "`n""build"":""" + $buildname + """," ;
                $postParams1 += "`n""host"":""" + $hostname + """,";$postParams1 += "`n""logs_location"":""" + $logs_location + """,";
                #$screenshotpath = $ws.cells.item($i,16).value()                
                if ($screenshotpath  -eq $null){$postParams1 += "`n""screenshot"":"+"""" + $screenshotpath + ""","}
                else{$screenshot = [convert]::ToBase64String((Get-Content "$screenshotpath" -Encoding byte));$postParams1 += "`n""screenshot"":""data:image/jpeg;base64,"+ $screenshot + """," }# convert screenshot to base64 image                                                                 
                $postParams1 += "`n""custom_params"":{"   
                $postParams1 += "`n""data_type"":""" + $DataType + """," # testcomplete as per the Template specified in portal                
                $postParams1 += "`n""steps"":{"          
                $postParams1 += "`n},"  
                $postParams1 += "`n""test_StartTime"":""" + $test_StartTime + """,";$postParams1 += "`n""test_EndTime"":""" + $test_EndTime + """";
                $postParams1 += "`n}";$postParams1 += "}";$response1 = Invoke-RestMethod "http://10.7.92.178:3000/products/test-complete/report" -ContentType "application/json" -Method Post -Body $postParams1 -TimeoutSec 600 #reporting path, send JSON as POST request
                Start-Sleep -s 3
                $response1
                Write-Host $postParams1
                
                #pause
                ##$wb_TestAutomation.Close($false);
                #$excel.quit();           

            }
            }
        }               

    }
    ) *>&1 >  $logfilename
    Stop-Process -processname Excel*

