<#
    Script to pull MS Patch states via WMI calls
    Created By - Leonard Hopkins
    Modified By - Kristopher Roy
    Modified On - 29 Jan 21 
#>


# Variables
$searchbase = "DC=Indyad,DC=Local"
$domain = "IndyAD"


$ErrorActionPreference = "SilentlyContinue"

If (!(Test-Path -Path 'C:\temp\SCCM')) {New-Item 'C:\temp\SCCM' -type directory}


Function AuditPatches{

$i = 0
$StartTime1 = Get-Date
$NumberofComputers = $Servers.Count

    Foreach ($Server in $Servers) { 

    Clear-Variable NewGroup
    Clear-Variable SCCMGroup
    Clear-Variable Group
      

    $i++
    $SecondsElapsed = ((Get-Date) - $StartTime1).TotalSeconds
    $SecondsRemaining = ($SecondsElapsed / ($i / $NumberofComputers)) - $SecondsElapsed
    Write-Progress -Activity "Getting Patches installed on $($Server): Server #$i of $($NumberofComputers)" -Id 1 -PercentComplete (($i/$($NumberofComputers)) * 100) -CurrentOperation "$("{0:N2}" -f ((($i/$($NumberofComputers)) * 100),2))% Complete" -SecondsRemaining $SecondsRemaining
    
    $PingResults = Test-Connection -Count 2 -Quiet -ComputerName $Server

            IF ($PingResults -ne "TRUE") {

                Write-Host "" 
                Write-Host $($Server).ToUpper()"is NOT Online." -ForegroundColor Red
                Write-Host ""
            
                $Offline = $Server.ToUpper() | Out-File "C:\temp\SCCM\$CurrentDate ServersNotOnLine.txt" -Append
                       
            }
    
            ElseIf($PingResults -eq "TRUE"){
             
                $n = 0
                $StartTime2 = Get-Date        
                $StartTime = Get-date -Format F

                $ComputerDetails = Get-ADComputer $Server -Properties * 
                $OSDetails = Get-WmiObject win32_operatingsystem -ComputerName ($Server) 
                $LastdateTime = [System.Management.ManagementDateTimeConverter]::ToDateTime($OSDetails.LastBootUpTime)
                $Date = $LastdateTime.ToShortDateString()
                $Time = $LastdateTime.ToShortTimeString()
                $ServerGroupMemberShipList = Get-ADPrincipalGroupMembership $ComputerDetails.SAMAccountName | Sort SAMAccountName | Where-Object {$_.Name -like "SCCM*"}
                $SCCMGroup = $ServerGroupMemberShipList.Name

                Write-Host "Patching report for " -ForegroundColor Cyan -NoNewline
                Write-host $Server.ToUpper() -ForegroundColor Yellow -NoNewline
                Write-Host " - " -ForegroundColor Cyan -NoNewline
                Write-Host "started on $StartTime"-ForegroundColor Cyan

                Write-Host "SCCM Group - " -ForegroundColor Cyan -NoNewline

                    Foreach ($Group in $SCCMGroup){$NewGroup += $Group + ", "}

                $SCCMGroup = $NewGroup.Substring(0,$NewGroup.Length-2)

                Write-Host "$SCCMGroup" -ForegroundColor Yellow

                Write-Host "Operating System" -NoNewline -ForegroundColor Cyan
                Write-Host " - " -ForegroundColor Cyan -NoNewline                                                      
                Write-Host $ComputerDetails.OperatingSystem -ForegroundColor Yellow 
                Write-Host "Last Restart" -ForegroundColor Cyan -NoNewline
                Write-Host " - " -ForegroundColor Cyan -NoNewline
                Write-Host $Date -NoNewline -ForegroundColor Yellow
                Write-Host " `@" $Time -ForegroundColor Yellow 
                Write-Host "_______________________________________________________________________________" -ForegroundColor Red

                $InstalledPatches = Get-WmiObject  Win32_QuickFixEngineering -ComputerName $Server  

                $EndTime = Get-date -Format F
                $NumberofPatches = $InstalledPatches.count

                Foreach ($InstalledPatch in $InstalledPatches){         

                    $n++
                    $SecondsElapsed = ((Get-Date) - $StartTime2).TotalSeconds
                    $SecondsRemaining = ($SecondsElapsed / ($n / $NumberofPatches)) - $SecondsElapsed

                    Write-Progress -Activity "Getting List of Patches on $($Server): Patch #$n of $($NumberofPatches)" -ParentId 1   -PercentComplete (($n/$($NumberofPatches)) * 100) -CurrentOperation "$("{0:N2}" -f ((($n/$($NumberofPatches)) * 100),2))% Complete" -SecondsRemaining $SecondsRemaining

                    Write-Host $InstalledPatch.Description -ForegroundColor Yellow -NoNewline
                    Write-Host " - " -ForegroundColor Yellow -NoNewline
                    Write-Host $InstalledPatch.HotfixID -NoNewline
                    Write-Host " - " -ForegroundColor Yellow -NoNewline
                    Write-Host $InstalledPatch.InstalledOn -ForegroundColor Cyan

                    $PatchesRetrieved = New-Object PSOBject
                    $PatchesRetrieved | Add-Member NoteProperty Server $Server
                    $PatchesRetrieved | Add-Member NoteProperty SCCMGroups $SCCMGroup
                    $PatchesRetrieved | Add-Member NoteProperty OS $ComputerDetails.OperatingSystem
                    $PatchesRetrieved | Add-Member NoteProperty PatchType $InstalledPatch.Description
                    $PatchesRetrieved | Add-Member NoteProperty KBArticle $InstalledPatch.HotfixID
                    $PatchesRetrieved | Add-Member NoteProperty DateInstalled $InstalledPatch.InstalledOn
                    $PatchesRetrieved | Add-Member NoteProperty ReportStartTime $StartTime
                    $PatchesRetrieved | Add-Member NoteProperty ReportEndTime $EndTime
                    $PatchesRetrieved | Export-Csv "C:\temp\SCCM\$CurrentDate IMCServer_Patching_History.csv" -NoTypeInformation -Append

                }
                
                Write-Host "Patching report for " -ForegroundColor Cyan -NoNewline
                Write-host $Server.ToUpper() -ForegroundColor Yellow -NoNewline
                Write-Host " ended on $EndTime"-ForegroundColor Cyan
                Write-Host 
                Write-Host

            }  

        }
    
    If(Test-Path "C:\temp\SCCM\$CurrentDate IMCServer_Patching_History.csv"){

        Write-Host " "
        Write-Host "Open Report? (Y/N)" -ForegroundColor Yellow
        $Continue = Read-Host 

        while("y","n" -notcontains $Continue){
        Write-Host "Y or N Please!" -ForegroundColor Red
        $Continue = Read-Host 
        }

        If($Continue -eq "Y"){Invoke-Item "C:\temp\SCCM\$CurrentDate IMCServer_Patching_History.csv"} 

    }
    

}   
 
#CLS

Write-Host " "
Write-Host "SCCM INSTALLED PATCHES AUDIT" -ForegroundColor Cyan
Write-Host "______________________________" -ForegroundColor Cyan
Write-Host " "

Do {

    $CurrentDate = Get-Date 
    $StartDate = Get-date -Format F
    Write-Host "Audit Date - " -ForegroundColor Green -NoNewline
    Write-Host $StartDate -ForegroundColor Yellow
    $CurrentDate = $CurrentDate.ToString('yyyy-MM-dd@HH-mm-ss')

    Write-Host "Do you want to check against one System, a List, or all of $domain Domain?" -ForegroundColor Yellow
    Write-Host "(Enter" -ForegroundColor Yellow -NoNewline
    Write-Host " 1 " -NoNewline
    Write-Host "for just one ," -ForegroundColor Yellow -NoNewline 
    Write-Host " 2 " -NoNewline
    Write-Host "for list," -ForegroundColor Yellow -NoNewline  
    Write-Host " 3 " -NoNewline
    Write-Host "for Specific Org Unit, or" -ForegroundColor Yellow -NoNewline
    Write-Host " 4 " -NoNewline
    Write-Host "for $domain Domain)" -ForegroundColor Yellow


    $Answer = Read-Host

        while("1","2","3","4" -notcontains $Answer){
        Write-Host "1, 2, 3, or 4 Please!" -ForegroundColor Red
        $Answer = Read-Host 
        }

        Switch ($Answer){

            1{

                DO{
                
                Write-Host "Enter the name of the System you want to check?" -ForegroundColor Yellow
                $Server = Read-Host
                $Servers = Get-ADComputer -LDAPFilter "(name=*$server*)" -SearchBase "$searchbase" | select Name -ExpandProperty Name

                    If (!$Servers){

                        Write-Host $Server.ToUpper() -ForegroundColor Green -NoNewline
                        Write-Host " is NOT Invalid! Try Again." -ForegroundColor Red
                        Write-Host
                    }

                }
                Until ($Servers)    
            }

            2{
                [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
                $dialog = New-Object System.Windows.Forms.OpenFileDialog
                $dialog = New-Object System.Windows.Forms.OpenFileDialog
                $dialog.FilterIndex = 0
                $dialog.InitialDirectory = "C:\Temp\Servers\Lists"
                $dialog.Multiselect = $false
                $dialog.RestoreDirectory = $true
                $dialog.Title = "Select a script file"
                $dialog.ValidateNames = $true
                $dialog.ShowDialog()
                $dialog.FileName

                $Servers = Get-content $dialog.FileName
            
            }
    
            3{

                Write-Host "Enter the Distinguished name of the OU you want to check?" -ForegroundColor Yellow
                $OrgUnits = Read-Host  

                 Foreach ($OrgUnit in $OrgUnits){

                    $Servers = (Get-ADComputer -Filter * -SearchBase "$OrgUnit" ).Name  | Sort 

            
                }    
            }

            4{

                $Servers = Get-ADComputer -LDAPFilter "(OperatingSystem=*Windows Server*)" -SearchBase "$searchbase" | Where-Object{$_.Enabled -eq "TRUE"} | Sort Name |Select-Object Name -ExpandProperty Name

            }


        }

AuditPatches


    Write-Host " "
    Write-Host "Do you want to check another server or list again? (Y/N)" -ForegroundColor Yellow
    $answer = Read-Host 

        while("y","n" -notcontains $answer){
        Write-Host "Y or N Please!" -ForegroundColor Red
        $answer = Read-Host 
        }

}
until ($Answer -eq "N") 

$Enddate = Get-date -Format F
Write-Host "Patching Report Audit Completed - " -ForegroundColor Green -NoNewline
Write-Host $Enddate -ForegroundColor Yellow