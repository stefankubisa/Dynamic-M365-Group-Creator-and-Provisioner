# Execution policy options
# ----------------------------------------------------------------------------------------------------
# Set-ExecutionPolicy Bypass -Scope Process
# Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope LocalMachine


# MASTER PLAN
# ----------------------------------------------------------------------------------------------------
<#
1. dynamic group creator 
    dynaic group creation - done
    dynamic membership criterion creation - done 
2. provisioner to google workspace application
    adder to the enterprise application - done
    security permissions changer 
    owner and manager assigner
#>


Clear-Host

$IsItDone_Pos = 1
$DisplayName_Pos = 2
$Email_Pos = 3
$MembershipRule_Pos = 4

$Path = $args[0]
$CurrentWorksheet = $args[1]

Write-Host "Creating Excel App Object" 
$excel = new-object -comobject Excel.Application 
Start-Sleep -Seconds 2
$excel.visible = $true 
$excel.DisplayAlerts = $false 
$excel.WindowState -4137
Write-Host "Opening Workbook"
Write-Host "____________________"

try {
    $workbook = $excel.workbooks.open($path)
}
catch {
    Write-Host $_.Exception.Message
    Write-Host "Closing Excel"
    Start-Sleep -Seconds 5
    $excel.Quit()
    throw $_
}
try {
    $Worksheet = $workbook.Worksheets.item($CurrentWorksheet)
}
catch {
    Write-Output $_.Exception.Message
    Write-Output "Closing Workbook"
    $workbook.Close()
    Write-Output "Closing Excel"
    Start-Sleep -Seconds 5
    $excel.Quit()
    throw $_
}

$verticalCount = (($Worksheet.UsedRange.Rows).count - 1 )
$horizontalCount = ($Worksheet.UsedRange.Columns).count - 3
$mailboxCount = $verticalCount
Write-Host -ForegroundColor DarkGreen "Mailbox Count: $mailboxCount"
$proceed = Read-Host "
If this makes sense and you ran this script previously/have the MgGraph module installed press 1 
to install the module press 2 
if you have a session going press 3
to abort press 4"
switch ($proceed) {
    1 {
        Import-Module Microsoft.Graph
        Connect-MgGraph -UseDeviceAuthentication 
        continue
    }
    2 {
        Install-Module PowerShellGet
        Install-Module -Name Microsoft.Graph -Scope AllUsers 
        Import-Module Microsoft.Graph
        Connect-MgGraph -Scopes "Directory.Read.All", "Application.Read.All" -UseDeviceAuthentication
        continue
    }
    3 {
        continue
    }
    4 {
        $excel.Quit()
        exit
    }
}

Write-Host "
What would you like to do:
--------------------------------------------------
1. Create Dynamic M365 Group with Display Name, Email and Dynamic Membership Querry 
"

$case = Read-Host "It's time to choose" 

$keepGoing = $true
while ($keepGoing) {
    for ($i = 1; $i -lt $verticalCount + 1; $i++) { 
        if (($Worksheet.Cells.Item($i + 1, $IsItDone_Pos)).Text -eq "OK" -or ($Worksheet.Cells.Item($i + 1, $IsItDone_Pos)).Text -eq "SKIP") {
            continue
        }
        $Worksheet.Cells.Item($i + 1, $IsItDone_Pos) = "In Progress"
        $Worksheet.Cells.Item($i + 1, $IsItDone_Pos).Interior.ColorIndex = 44
        $DisplayName = $Worksheet.Cells.Item($i + 1, $DisplayName_Pos).Text.trim() 
        $Email = $Worksheet.Cells.Item($i + 1, $Email_Pos).Text.trim() 
        $Querry = $Worksheet.Cells.Item($i + 1, $MembershipRule_Pos).Text.trim() 
        Write-Host "$Querry"
        $EmailSplit = ($Email).split("@")[0].trim() 
        Write-Host "$EmailSplit"
        $TextInfo = (Get-Culture).TextInfo 
        $EmailFront = $TextInfo.ToTitleCase($EmailSplit[0]) 

        Write-Host -ForegroundColor DarkGreen "_____________________________" 
        Write-Host -ForegroundColor DarkGreen "$DisplayName" 
        Write-Host -ForegroundColor DarkGreen "$Email" 
        Write-Host -ForegroundColor DarkGreen "-----------------------------" 

        if($case -eq 1) { 
            try {
                $group = Get-MgGroup -Filter "mail eq '$Email'"
                $groupId = $group.Id
                Remove-MgGroup -GroupId $groupId
                Write-Host -ForegroundColor Red "_____________________________" 
                Write-Output "Group with email '$email' has been deleted."
                Write-Host -ForegroundColor Red "-----------------------------" 
            }
            catch {
                Write-Output "Group with email '$email' not found."
            }

            Start-Sleep -Seconds 5 

            New-MgGroup -DisplayName $DisplayName `
                        -MailNickname $EmailSplit `
                        -MailEnabled:$true `
                        -SecurityEnabled:$false `
                        -GroupTypes @("Unified","DynamicMembership") `
                        -MembershipRule $Querry `
                        -MembershipRuleProcessingState "On"
            
            Write-Host -ForegroundColor Red "GIVING MGGRAPH TIME TO PROCESS THE CHANGES"
            $Worksheet.Cells.Item($i + 1, $j + 4).Interior.ColorIndex = 43
            Start-Sleep -Seconds 10 
        
            $group = Get-MgGroup -Filter "displayName eq '$displayName'"
            $groupId = $group.Id

            $appRoleId = "" 

            # Assign the group to the service principal
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $servicePrincipalId `
                                                    -PrincipalId $groupId `
                                                    -ResourceId $servicePrincipalId `
                                                    -AppRoleId $appRoleId
            }
            $Worksheet.Cells.Item($i + 1, $IsItDone_Pos) = "OK"
            $Worksheet.Cells.Item($i + 1, $IsItDone_Pos).Interior.ColorIndex = 43
        }         
        $keepGoing = $false
}

$workbook.Save()
Write-Host "All done, closing workbook"
Start-Sleep -Seconds 2
$excel.Quit()