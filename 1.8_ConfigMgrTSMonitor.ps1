#requires -Version 3
$currentLocation = if($PSScriptRoot){ $PSScriptRoot } else { (Get-Location).Path }
Write-Host $currentLocation
Set-Location $currentLocation

#region Add Assemblies
Add-Type -AssemblyName PresentationFramework, System.Drawing, System.Windows.Forms, WindowsFormsIntegration
$code = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System
{
	public class IconExtractor
	{

	 public static Icon Extract(string file, int number, bool largeIcon)
	 {
	  IntPtr large;
	  IntPtr small;
	  ExtractIconEx(file, number, out large, out small, 1);
	  try
	  {
	   return Icon.FromHandle(largeIcon ? large : small);
	  }
	  catch
	  {
	   return null;
	  }

	 }
	 [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
	 private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

	}
}
"@
Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing

# Mahapps Library
if (Test-Path -Path "$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\MahApps.Metro.dll")
{
    [System.Reflection.Assembly]::LoadFrom("$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\MahApps.Metro.dll")    | out-null
    [System.Reflection.Assembly]::LoadFrom("$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\System.Windows.Interactivity.dll") | out-null
}

if (Test-Path -Path "${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\MahApps.Metro.dll")
{
    [System.Reflection.Assembly]::LoadFrom("${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\MahApps.Metro.dll")    | out-null
    [System.Reflection.Assembly]::LoadFrom("${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\System.Windows.Interactivity.dll") | out-null
}

if (Test-Path -Path "$currentLocation\MahApps.Metro.dll")
{
    [System.Reflection.Assembly]::LoadFrom("$currentLocation\MahApps.Metro.dll")    | out-null
    [System.Reflection.Assembly]::LoadFrom("$currentLocation\System.Windows.Interactivity.dll") | out-null
}

#endregion

#region GUI and Variables
### Main Window ###
# GUI
[xml]$xaml = Get-Content ($currentLocation + "\XAML\MainWindow.xaml")

$BuildExtVersionSql = @()
Import-Csv -Path ($currentLocation + "\BuildExt.csv") -Delimiter ";" -Header Build, Version | ForEach-Object {
    $BuildExtVersionSql += "WHEN sys.BuildExt like '$($_.Build)' THEN '$($_.Version)'"

}


$hash = [hashtable]::Synchronized(@{})
$reader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml)
$hash.Window = [Windows.Markup.XamlReader]::Load( $reader )
$global:PSInstances = @()
$Global:Timezones = @()

$hash.TaskSequence = $hash.Window.FindName('TaskSequence')
$hash.TimePeriod = $hash.Window.FindName('TimePeriod')
$hash.ErrorsOnly = $hash.Window.FindName('ErrorsOnly')
$hash.SuccessCode = $hash.Window.FindName('SuccessCode')
$hash.DisabledSteps = $hash.Window.FindName('DisabledSteps')
$hash.ComputerName = $hash.Window.FindName('ComputerName')
$hash.BuildExt = $hash.Window.FindName('BuildExt')
$hash.ActionName = $hash.Window.FindName('ActionName')
$hash.RefreshPeriod = $hash.Window.FindName('RefreshPeriod')
$hash.RefreshNow = $hash.Window.FindName('RefreshNow')
$hash.DataGrid = $hash.Window.FindName('DataGrid')
$hash.ActionOutput = $hash.Window.FindName('ActionOutput')
$hash.MDTGroupBox = $hash.Window.FindName('MDTGroupBox')
$hash.MDTIntegrated = $hash.Window.FindName('MDTIntegrated')
$hash.DeploymentStatus = $hash.Window.FindName('DeploymentStatus')
$hash.CurrentStep = $hash.Window.FindName('CurrentStep')
$hash.StepName = $hash.Window.FindName('StepName')
$hash.PercentComplete = $hash.Window.FindName('PercentComplete')
$hash.MDTStartTime = $hash.Window.FindName('MDTStartTime')
$hash.MDTEndTime = $hash.Window.FindName('MDTEndTime')
$hash.MDTElapsedTime = $hash.Window.FindName('MDTElapsedTime')
$hash.SettingsButton = $hash.Window.FindName('SettingsButton')
$hash.ReportButton = $hash.Window.FindName('ReportButton')
$hash.ErrorCount = $hash.Window.FindName('ErrorCount')

$hash.DeploymentStatusLabel = $hash.Window.FindName('DeploymentStatusLabel')
$hash.CurrentStepLabel = $hash.Window.FindName('CurrentStepLabel')
$hash.StepNameLabel = $hash.Window.FindName('StepNameLabel')
$hash.PercentCompleteLabel = $hash.Window.FindName('PercentCompleteLabel')
$hash.StartLabel = $hash.Window.FindName('StartLabel')
$hash.EndLabel = $hash.Window.FindName('EndLabel')
$hash.ElapsedLabel = $hash.Window.FindName('ElapsedLabel')
$hash.ProgressLabel = $hash.Window.FindName('ProgressLabel')

$hash.ProgressBar = $hash.Window.FindName('ProgressBar')
if (Test-Path -Path "$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico")
{
    $hash.Window.Icon = "$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico"
}
if (Test-Path -Path "${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico")
{
    $hash.Window.Icon = "${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico"
}
if (Test-Path -Path "$currentLocation\Grid.ico")
{
	$hash.Window.add_Loaded({
		$hash.Window.Icon = "$currentLocation\Grid.ico"
	})
	#$hash.Window.TaskbarItemInfo.Overlay = "$currentLocation\Grid.ico"
}

### Settings Window ###
[xml]$xaml2 = Get-Content ($currentLocation + "\XAML\Config.xaml")

$reader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml2)
$hash.Window2 = [Windows.Markup.XamlReader]::Load( $reader )
$hash.SQLServer = $hash.Window2.FindName('SQLServer')
$hash.Database = $hash.Window2.FindName('Database')
$hash.MDT = $hash.Window2.FindName('MDT')
$hash.MDTURL = $hash.Window2.FindName('MDTURL')
$hash.ConnectSQL = $hash.Window2.FindName('ConnectSQL')
$hash.TSList = $hash.Window2.FindName('TSList')
$hash.StartDate = $hash.Window2.FindName('StartDate')
$hash.EndDate = $hash.Window2.FindName('EndDate')
$hash.GenerateReport = $hash.Window2.FindName('GenerateReport')
$hash.SettingsTab = $hash.Window2.FindName('SettingsTab')
$hash.ReportTab = $hash.Window2.FindName('ReportTab')
$hash.Tabs = $hash.Window2.FindName('Tabs')
$hash.Working = $hash.Window2.FindName('Working')
$hash.Runasadmin = $hash.Window2.FindName('Runasadmin')
$hash.ReportProgress = $hash.Window2.FindName('ReportProgress')
$hash.Link1 = $hash.Window2.FindName('Link1')
$hash.DTFormat = $hash.Window2.FindName('DTFormat')

if (Test-Path -Path "$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico")
{
    #$hash.Window2.Icon = "$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico"
    $Hash.Window2.ShowInTaskbar = $true
}
if (Test-Path -Path "${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico")
{
    #$hash.Window2.Icon = "${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico"
    $Hash.Window2.ShowInTaskbar = $true
}
if (Test-Path -Path "$currentLocation\Grid.ico")
{
    #$hash.Window2.Icon = "$currentLocation\Grid.ico"
    $Hash.Window2.ShowInTaskbar = $true
}


$script:SQLServer = $hash.SQLServer.Text
$Script:Database = $hash.Database.Text
#endregion

#region Icons and Runspacepool
# Output SystemIcons to bmps
$icons = @()
$global:greentickiconpath = "$env:temp\GreenTick.bmp"
$icons += $greentickiconpath 
$global:redcrossiconpath = "$env:temp\RedCross.bmp"
$icons += $redcrossiconpath

if (!(Test-Path $greentickiconpath))
{
    $global:greentickicon = [System.IconExtractor]::Extract('comres.dll',8,$true).ToBitmap()
    $greentickicon.save("$greentickiconpath")
}
if (!(Test-Path $redcrossiconpath))
{
    $global:redcrossicon = [System.IconExtractor]::Extract('comres.dll',10,$true).ToBitmap()
    $redcrossicon.save("$redcrossiconpath")
}

$script:RunspacePool = [runspacefactory]::CreateRunspacePool()
$RunspacePool.ApartmentState = 'STA'
$RunspacePool.ThreadOptions = 'ReUseThread'
$RunspacePool.Open()
#endregion

#region Functions

Function Get-DateTimeFormat 
{
    if ([System.TimeZone]::CurrentTimeZone.IsDaylightSavingTime($(Get-Date)))
    {
        $TimeZone = [System.TimeZone]::CurrentTimeZone.DaylightName
    }
    Else 
    {
        $TimeZone = [System.TimeZone]::CurrentTimeZone.StandardName
    }

    #$Global:Timezones = @()
    $obj = New-Object -TypeName psobject -Property @{
        TimeZone = 'UTC'
    }
    $Global:Timezones = [Array]$Timezones + $obj
    $obj = New-Object -TypeName psobject -Property @{
        TimeZone = $TimeZone
    }
    $Global:Timezones = [Array]$Timezones + $obj
}

Function Get-TaskSequenceList 
{
    # Set variables
    $script:SQLServer = $hash.SQLServer.Text
    $Script:Database = $hash.Database.Text

    # If SQLinstance not populated, ask for connection
    if ($SQLServer -eq '<SQLServer\Instance>')
    {
        $hash.ActionOutput.Text = 'No SQL Server defined.  Click Settings, and set the SQL Server, database and MDT URL if applicable.'
        return
    }

    # Connect to SQL server
    try
    {
        $connectionString = "Server=$SQLServer;Database=$Database;Integrated Security=SSPI;"
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString
        $connection.Open()
        $hash.ActionOutput.Text = 'Connected to SQL Server database.  Select a Task Sequence.'
    }
    catch 
    {
        $hash.ActionOutput.Text = '[ERROR} Could not connect to SQL Server database!'
        return
    }
    # Run SQL query
	#$Query = "
    #    SELECT DISTINCT summ.SoftwareName AS 'Task Sequence'
    #    FROM vDeploymentSummary summ
    #    WHERE (summ.FeatureType=7)
    #    ORDER BY summ.SoftwareName
    #"
    $Query = "
		SELECT DISTINCT v_TaskSequencePackage.Name AS 'Task Sequence' FROM v_TaskSequencePackage
		INNER JOIN v_Program ON v_Program.PackageID = v_TaskSequencePackage.PackageID
		WHERE (0x00001000 & dbo.v_Program.ProgramFlags)/0x00001000 != 1
		ORDER BY v_TaskSequencePackage.Name
    "
    $command = $connection.CreateCommand()
    $command.CommandText = $Query
    $result = $command.ExecuteReader()
    $table = New-Object -TypeName 'System.Data.DataTable'
    $table.Load($result)
    $connection.Close()

    # Load data into psobject            
    $global:Views = @()
    Foreach ($Row in $table.Rows)
    {
        $obj = New-Object -TypeName psobject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name 'TS' -Value $Row.'Task Sequence'
        $global:Views = [Array]$Views + $obj
    }

    # Output to Task Sequence combobox
    $hash.Window.Dispatcher.Invoke(
        [action]{
            $hash.TaskSequence.ItemsSource = [Array]$Views.TS
    })
}

Function Get-TaskSequenceData 
{
    param ($hash,$RunspacePool)

    $code = 
    {
        param($hash,$SQLServer,$Database,$BuildExtVersionSql,$TimePeriod,$SuccessCode,$ErrorsOnly,$DisabledSteps,$ComputerName,$ActionName,$TS,$MDTIntegrated,$URL,$DTFormat)

        # Notify of data retrieval         
        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.ActionOutput.Text = 'Retrieving data...'
                $hash.DataGrid.ItemsSource = ''
        })

        # Set variable values
        if ($MyGUID)
        {
            Remove-Variable -Name MyGuid
        }
        if ($Unknowns) # put after display ####
        {
            Remove-Variable -Name Unknowns
        }

        if ($ErrorsOnly -eq 'True')
        {
            $ExitCode = $SuccessCode
        }
        else 
        {
            $ExitCode = 999999999999999999999999
        }
		
		if ($DisabledSteps -eq 'True')
        {
            $DisabledStep = "''"
        }
        Else 
        {
            $DisabledStep = '11128'
        }
        
        if ($ComputerName.DisplayName -eq '-All-' -or $ComputerName.DisplayName -eq '' -or $ComputerName -eq $Null)
        {
            $SQLComputerName = '%'
        }
        Else 
        {
            $SQLComputerName = $ComputerName.Value
			$hash.ActionOutput.Text = "Search for device: $($ComputerName.DisplayName) - $($ComputerName.Value)"
        }
		
		if ($ActionName -eq '-All-' -or $ActionName -eq '' -or $ActionName -eq $Null)
        {
            $SQLActionName = '%'
        }
        Else 
        {
            $SQLActionName = $ActionName
        }
        
        $greentickiconpath = "$env:temp\GreenTick.bmp"
        $redcrossiconpath = "$env:temp\RedCross.bmp"
        
        # Connect to SQL server
        try
        {
            $connectionString = "Server=$SQLServer;Database=$Database;Integrated Security=SSPI;"
            $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
            $connection.ConnectionString = $connectionString
            $connection.Open()
        }
        catch 
        {
            $MyError = $_.Exception.Message
            $hash.Window.Dispatcher.Invoke(
                [action]{
                    $hash.ActionOutput.Text = "[ERROR} Could not connect to SQL Server database! $MyError"
            })
            return
        }
        
        if ($MDTIntegrated -eq 'True')
        {
            # Get Unknown Computers from ConfigMgr database if there are any
            $Query = "
                Select Distinct Name0,
                SMBIOS_GUID0 as 'GUID'
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                --and DATEDIFF(hour,ExecutionTime,GETDATE()) < $TimePeriod
				and DATEDIFF(hour,(CONVERT(datetime, SWITCHOFFSET(CONVERT(datetimeoffset, ExecutionTime), DATENAME(TzOffset, SYSDATETIMEOFFSET()))) ),GETDATE()) <= $TimePeriod
                --and Name0 like '$SQLComputerName'
				--and ActionName like '$SQLActionName'
                and ExitCode not in ($ExitCode)
				and tes.LastStatusMsgID not in ($DisabledStep)
                ORDER BY Name0 Desc
            "
            $command = $connection.CreateCommand()
            $command.CommandText = $Query
            $result = $command.ExecuteReader()
            $table = New-Object -TypeName 'System.Data.DataTable'
            $table.Load($result)

        
            # Gather unknowns into PS object    
            $UnknownComputers = @()
            Foreach ($Row in $table.Rows | Where-Object -FilterScript {
                    $_.Name0 -eq 'Unknown'
            })
            {
                $obj = New-Object -TypeName psobject
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ComputerName' -Value $Row.Name0
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'GUID' -Value $Row.GUID
                $UnknownComputers += $obj
            }


            # If there are unknowns, get computername from MDT
            if ($UnknownComputers.Count -ge 1)
            {
                $URL1 = $URL.Replace('Computers','ComputerIdentities')

                # Get ID numbers and Identifiers (GUIDs)
                function GetMDTIDs 
                { 
                    param ($URL)
                    $Data = Invoke-RestMethod -Uri $URL
                    foreach($property in ($Data.content.properties) ) 
                    { 
                        New-Object -TypeName PSObject -Property @{
                            ID         = $($property.ID.'#text')
                            Identifier = $($property.Identifier)
                        }
                    } 
                }
                
                # Filter out only the GUIDs
                $MDTIDs = GetMDTIDs -URL $URL1 |
                Select-Object -Property * |
                Where-Object -FilterScript {
                    $_.Identifier -like '*-*'
                } |
                Sort-Object -Property ID
                $MDTComputerIDs = @()
                Foreach ($Computer in $UnknownComputers)
                {
                    $MDTComputerID = $MDTIDs |
                    Where-Object -FilterScript {
                        $_.Identifier -eq $Computer.GUID
                    } |
                    Select-Object -Property ID, Identifier 
                    $MDTComputerIDs += $MDTComputerID
                }
                
                # Get ComputerNames from MDT
                function GetMDTComputerNames 
                { 
                    param ($URL)
                    $Data = Invoke-RestMethod -Uri $URL
                    foreach($property in ($Data.content.properties) ) 
                    { 
                        New-Object -TypeName PSObject -Property @{
                            Name = $($property.Name)
                            ID   = $($property.ID.'#text')
                        } 
                    } 
                } 
                
                # Filter out the computer names from the IDs
                $MDTComputers = GetMDTComputerNames -URL $URL |
                Select-Object -Property * |
                Sort-Object -Property ID

                $ResolvedComputerNames = @()
                Foreach ($MDTComputerID in $MDTComputerIDs)
                {
                    $MDTComputerName = $MDTComputers |
                    Where-Object -FilterScript {
                        $_.ID -eq $MDTComputerID.ID
                    } |
                    Select-Object -ExpandProperty Name
                    $GUID = $MDTIDs |
                    Where-Object -FilterScript {
                        $_.ID -eq $MDTComputerID.ID
                    } |
                    Select-Object -ExpandProperty Identifier
                    $obj = New-Object -TypeName PSObject
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name ComputerName -Value $MDTComputerName
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name GUID -Value $GUID
                    $ResolvedComputerNames += $obj
                }
            }
            foreach ($Computer in $ResolvedComputerNames)
            {
                if ($ComputerName -eq $Computer.ComputerName)
                {
                    $MyGUID = $Computer.GUID
                }
            }
        }


        # Get TS execution data from ConfigMgr
        if ($MyGUID)
        {
            $Query = "
                Select Distinct sys.Name0 as 'Computer Name',
                sys.SMBIOS_GUID0 as 'GUID',
                tsp.Name as 'Task Sequence',
				comp.UserName0,
				CASE
					WHEN cmcbs.CNIsOnInternet = 0 THEN 'Intranet'
					WHEN cmcbs.CNIsOnInternet = 1 THEN 'Internet'
					ELSE CAST(cmcbs.CNIsOnInternet AS varchar)
					END as [Connection Type], 
				CAST(
					 CASE
						  $($BuildExtVersionSql -join "`n")
						  ELSE sys.BuildExt
					 END AS char) as BuildExt, 
				comp.Model0,
				BIOS.SMBIOSBIOSVersion0 as BIOSVersion,
                ExecutionTime,
                Step,
                tes.ActionName,
                GroupName,
                tes.LastStatusMsgName,
                ExitCode,
                ActionOutput
                from vSMS_TaskSequenceExecutionStatus tes
                INNER JOIN v_R_System sys on tes.ResourceID = sys.ResourceID
				INNER JOIN (select MachineID, Name, CNIsOnInternet, LastPolicyRequest, LastDDR as [Last Heartbeat],
				   LastHardwareScan, max(CNLastOnlinetime) as [Last Online Time]
				   FROM v_CollectionMemberClientBaselineStatus
				   GROUP BY Name, MachineID, CNIsOnInternet, ClientVersion, LastPolicyRequest, LastDDR,
				   LastHardwareScan, CNLastOnlinetime) cmcbs ON cmcbs.MachineID = sys.ResourceID
				--INNER JOIN v_CollectionMemberClientBaselineStatus cmcbs ON cmcbs.MachineID = sys.ResourceID
				LEFT JOIN v_GS_COMPUTER_SYSTEM comp ON comp.ResourceID = sys.ResourceID
				LEFT JOIN v_GS_PC_BIOS BIOS ON BIOS.ResourceID = sys.ResourceID
                LEFT JOIN v_RA_System_MACAddresses mac on tes.ResourceID = mac.ResourceID
                INNER JOIN v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                --and DATEDIFF(hour,ExecutionTime,GETDATE()) <= $TimePeriod
				and DATEDIFF(hour,(CONVERT(datetime, SWITCHOFFSET(CONVERT(datetimeoffset, ExecutionTime), DATENAME(TzOffset, SYSDATETIMEOFFSET()))) ),GETDATE()) <= $TimePeriod
                --and sys.Name0 like '$SQLComputerName'
				--and ActionName like '$SQLActionName'
                and sys.SMBIOS_GUID0 = '$MyGUID'
                and ExitCode not in ($ExitCode)
				and tes.LastStatusMsgID not in ($DisabledStep)
                ORDER BY ExecutionTime Desc
            "
            $ErrQuery = "
                Select Count(Name0) as 'Count' from (Select DISTINCT (Name0), ActionName, ExecutionTime 
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                left join v_RA_System_MACAddresses mac on tes.ResourceID = mac.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                --and DATEDIFF(hour,ExecutionTime,GETDATE()) <= $TimePeriod
				and DATEDIFF(hour,(CONVERT(datetime, SWITCHOFFSET(CONVERT(datetimeoffset, ExecutionTime), DATENAME(TzOffset, SYSDATETIMEOFFSET()))) ),GETDATE()) <= $TimePeriod
                --and Name0 like '$SQLComputerName'
				--and ActionName like '$SQLActionName'
                and sys.SMBIOS_GUID0 = '$MyGUID'
                and ExitCode not in ($SuccessCode)) as t
            "
        }

        if (!$MyGUID)
        {
			if ($SQLComputerName -eq '%') {
				$specificCondition = "and sys.Name0 like '$SQLComputerName'"
			} else {
				$specificCondition = "and sys.SMBIOS_GUID0 = '$SQLComputerName'"
			}
			
            $baseQuery = "
				Select Distinct sys.Name0 as 'Computer Name',
				sys.SMBIOS_GUID0 as 'GUID',
				tsp.Name as 'Task Sequence',
				comp.UserName0,
				CASE
					WHEN cmcbs.CNIsOnInternet = 0 THEN 'Intranet'
					WHEN cmcbs.CNIsOnInternet = 1 THEN 'Internet'
					ELSE CAST(cmcbs.CNIsOnInternet AS varchar)
				END as [Connection Type], 
				CAST(
					 CASE
						  $($BuildExtVersionSql -join "`n")
						  ELSE sys.BuildExt
					 END AS char) as BuildExt, 
				comp.Model0,
				BIOS.SMBIOSBIOSVersion0 as BIOSVersion,
				ExecutionTime,
				Step,
				tes.ActionName,
				GroupName,
				tes.LastStatusMsgName,
				ExitCode,
				ActionOutput
				from vSMS_TaskSequenceExecutionStatus tes
				INNER JOIN v_R_System sys on tes.ResourceID = sys.ResourceID
				INNER JOIN (select MachineID, Name, CNIsOnInternet, LastPolicyRequest, LastDDR as [Last Heartbeat],
				   LastHardwareScan, max(CNLastOnlinetime) as [Last Online Time]
				   FROM v_CollectionMemberClientBaselineStatus
				   GROUP BY Name, MachineID, CNIsOnInternet, ClientVersion, LastPolicyRequest, LastDDR,
				   LastHardwareScan, CNLastOnlinetime) cmcbs ON cmcbs.MachineID = sys.ResourceID
				LEFT JOIN v_GS_COMPUTER_SYSTEM comp ON comp.ResourceID = sys.ResourceID
				LEFT JOIN v_GS_PC_BIOS BIOS ON BIOS.ResourceID = sys.ResourceID
				LEFT JOIN v_RA_System_MACAddresses mac on tes.ResourceID = mac.ResourceID
				INNER JOIN v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
				where tsp.Name = '$TS'
				and DATEDIFF(hour, (CONVERT(datetime, SWITCHOFFSET(CONVERT(datetimeoffset, ExecutionTime), DATENAME(TzOffset, SYSDATETIMEOFFSET())))), GETDATE()) <= $TimePeriod
				and ActionName like '$SQLActionName'
				and ExitCode not in ($ExitCode)
				and tes.LastStatusMsgID not in ($DisabledStep)
			"

			$Query = "$baseQuery $specificCondition ORDER BY ExecutionTime Desc"
			
            $ErrQuery = "
                Select Count(Name0) as 'Count' from (Select DISTINCT (Name0), ActionName, ExecutionTime 
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                left join v_RA_System_MACAddresses mac on tes.ResourceID = mac.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                --and DATEDIFF(hour,ExecutionTime,GETDATE()) <= $TimePeriod
				and DATEDIFF(hour,(CONVERT(datetime, SWITCHOFFSET(CONVERT(datetimeoffset, ExecutionTime), DATENAME(TzOffset, SYSDATETIMEOFFSET()))) ),GETDATE()) <= $TimePeriod
				and ActionName like '$SQLActionName'
                and ExitCode not in ($SuccessCode) $specificCondition) as t
            "
			
        }
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $result = $command.ExecuteReader()
        $table = New-Object -TypeName 'System.Data.DataTable'
        $table.Load($result)
        $command = $connection.CreateCommand()
        $command.CommandText = $ErrQuery
        $erresult = $command.ExecuteReader()
        $errtable = New-Object -TypeName 'System.Data.DataTable'
        $errtable.Load($erresult)
        $connection.Close()

        if ($table.rows.Count -lt 1)
        {
            $hash.Window.Dispatcher.Invoke(
                [action]{
                    $hash.ActionOutput.Text = 'No results.'
					#$hash.ActionOutput.Text = $query
            })
            return
        }

        # Gather results into psobject            
        $global:Results = @()
        $i = 0
        Foreach ($Row in $table.Rows)
        {
            $obj = New-Object -TypeName psobject
            $i ++
            if ($Row.ExitCode -in $SuccessCode.Replace(" ", "").Split(","))
            {
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Icon' -Value $greentickiconpath
            }
            if ($Row.ExitCode -notin $SuccessCode.Replace(" ", "").Split(","))
            {
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Icon' -Value $redcrossiconpath
            }
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ComputerName' -Value $Row.'Computer Name'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'GUID' -Value $Row.'GUID'
			Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Connection Type' -Value $Row.'Connection Type'
			Add-Member -InputObject $obj -MemberType NoteProperty -Name 'BuildExt' -Value $Row.'BuildExt'
			Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Model0' -Value $Row.'Model0'
			Add-Member -InputObject $obj -MemberType NoteProperty -Name 'BIOSVersion' -Value $Row.'BIOSVersion'
            if ($DTFormat -eq 'UTC')
            {
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ExecutionTime' -Value $Row.'ExecutionTime'
            }
            Else 
            {
                $extime = [System.TimeZone]::CurrentTimeZone.ToLocalTime($($Row.'ExecutionTime' | Get-Date))
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ExecutionTime' -Value $extime
            }
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Step' -Value $Row.'Step'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ActionName' -Value $Row.'ActionName'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'GroupName' -Value $Row.'GroupName'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'LastStatusMsgName' -Value $Row.'LastStatusMsgName'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ExitCode' -Value $Row.'ExitCode'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ActionOutput' -Value $Row.'ActionOutput'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Record' -Value $i
            $Results += $obj
        }
        if ($Results.Count -eq 1)
        {
            $obj = New-Object -TypeName psobject
            $i ++
            if ($Row.ExitCode -in $SuccessCode.Replace(" ", "").Split(","))
            {
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Icon' -Value $greentickiconpath
            }
            if ($Row.ExitCode -notin $SuccessCode.Replace(" ", "").Split(","))
            {
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Icon' -Value $redcrossiconpath
            }
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ComputerName' -Value ' '
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'GUID' -Value ' '
			Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Connection Type' -Value ' '
			Add-Member -InputObject $obj -MemberType NoteProperty -Name 'BuildExt' -Value ' '
			Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Model0' -Value ' '
			Add-Member -InputObject $obj -MemberType NoteProperty -Name 'BIOSVersion' -Value ' '
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ExecutionTime' -Value ' '
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Step' -Value ' '
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ActionName' -Value ' '
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'GroupName' -Value ' '
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'LastStatusMsgName' -Value ' '
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ExitCode' -Value ' '
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ActionOutput' -Value ' '
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Record' -Value ' '
            $Results += $obj
        }

        $FilteredResults = $Results | Select-Object -Property Icon, ComputerName, GUID, 'Connection Type', BuildExt, Model0, BIOSVersion, ExecutionTime, Step, ActionName, GroupName, LastStatusMsgName, ExitCode, Record

        if (!$MyGUID -and $ComputerName -in ('-ALL-', '', $Null))
        {
            if ($FilteredResults.ComputerName -match 'unknown')
            {
                foreach ($Computer in $ResolvedComputerNames)
                {
                    $Unknowns = $FilteredResults | Where-Object -FilterScript {
                        $_.ComputerName -eq 'Unknown' -and $_.GUID -eq $Computer.GUID
                    }
                    $i = -1
                    do
                    {
                        $i ++
                        $Unknowns[$i].ComputerName = $Computer.ComputerName
                    }
                    until ($i -eq ($Unknowns.Count -1))
                }
            }
        }

        if ($MyGUID)
        {
            $Unknowns = $FilteredResults | Where-Object -FilterScript {
                $_.ComputerName -eq 'Unknown' -and $_.GUID -eq $MyGUID
            }
            if ($Unknowns)
            {
                $i = -1
                do
                {
                    $i ++
                    $Unknowns[$i].ComputerName = $ComputerName
                }
                until ($i -eq ($Unknowns.Count -1))
            }
        }
        
        # Display results in datagrid         
        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.DataGrid.ItemsSource = $FilteredResults
                $hash.ErrorCount.Text = $errtable.Count
                $hash.ActionOutput.Text = 'Click any step to see the action output.'
        })
    }

    # Set variables from Hash table
    $SQLServer = $hash.SQLServer.Text
    $Database = $hash.Database.Text
    $TimePeriod = $hash.TimePeriod.Text
	$SuccessCode = $hash.SuccessCode.Text
    $ErrorsOnly = $hash.ErrorsOnly.IsChecked
	$DisabledSteps = $hash.DisabledSteps.IsChecked
    $ComputerName = $hash.ComputerName.SelectedItem
	$ActionName = $hash.ActionName.SelectedItem
    $TS = $hash.TaskSequence.SelectedItem
    $MDTIntegrated = $hash.MDTIntegrated.IsChecked
    $URL = $hash.MDTURL.Text
    $DTFormat = $hash.DTFormat.SelectedItem

    # Create PS instance in runspace pool and execute
    $PSinstance = [powershell]::Create().AddScript($code).AddArgument($hash).AddArgument($SQLServer).AddArgument($Database).AddArgument($BuildExtVersionSql).AddArgument($TimePeriod).AddArgument($SuccessCode).AddArgument($ErrorsOnly).AddArgument($DisabledSteps).AddArgument($ComputerName).AddArgument($ActionName).AddArgument($TS).AddArgument($MDTIntegrated).AddArgument($URL).AddArgument($DTFormat)

    $PSInstances += $PSinstance
    $PSinstance.RunspacePool = $RunspacePool
    $PSinstance.BeginInvoke()
}

Function Populate-ActionOutput 
{
    param ($hash,$RunspacePool)

    $code = 
    {
        param($hash,$Record)
        $msg = $Results |
        Select-Object -Property * |
        Where-Object -FilterScript {
            $_.Record -eq $Record
        }
        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.ActionOutput.Text = $msg.ActionOutput
        })
    }

    # Set variables from Hash table
    $Record = $hash.DataGrid.SelectedItem.Record

    # Create PS instance in runspace pool and execute
    $PSinstance = [powershell]::Create().AddScript($code).AddArgument($hash).AddArgument($Record)
    $PSInstances += $PSinstance
    $PSinstance.RunspacePool = $RunspacePool
    $PSinstance.BeginInvoke()
}

Function Populate-ComputerNames 
{
    param ($hash,$RunspacePool)

    $code = 
    {
        param($hash,$SQLServer,$Database,$BuildExtVersionSql,$TimePeriod,$SuccessCode,$ErrorsOnly,$DisabledSteps,$TS,$MDTIntegrated,$URL)
        
        # Set variable values
        if ($ErrorsOnly -eq 'True')
        {
            $ExitCode = $SuccessCode
        }
        else 
        {
            $ExitCode = 999999999999999999999999
        }
		
		if ($DisabledSteps -eq 'True')
        {
            $DisabledStep = "''"
        }
        Else 
        {
            $DisabledStep = '11128'
        }

        # Connect to SQL Server
        $connectionString = "Server=$SQLServer;Database=$Database;Integrated Security=SSPI;"
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString
        $connection.Open()

        # Run SQL query
        $Query = "
            Select Distinct Name0,
            SMBIOS_GUID0 as 'GUID',
			CAST(
				CASE
				  $($BuildExtVersionSql -join "`n")
				  ELSE sys.BuildExt
				END AS char) as BuildExt
            from vSMS_TaskSequenceExecutionStatus tes
            inner join v_R_System sys on tes.ResourceID = sys.ResourceID
            inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
            where tsp.Name = '$TS'
            --and DATEDIFF(hour,ExecutionTime,GETDATE()) <= $TimePeriod
			and DATEDIFF(hour,(CONVERT(datetime, SWITCHOFFSET(CONVERT(datetimeoffset, ExecutionTime), DATENAME(TzOffset, SYSDATETIMEOFFSET()))) ),GETDATE()) <= $TimePeriod
            --and Name0 like '$SQLComputerName'
            and ExitCode not in ($ExitCode)
			and tes.LastStatusMsgID not in ($DisabledStep)
            ORDER BY Name0 Desc
        "
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $result = $command.ExecuteReader()
        $table = New-Object -TypeName 'System.Data.DataTable'
        $table.Load($result)
        $connection.Close()
         
        # Gather results into PS object    
        $PCResults = @()
        Foreach ($Row in $table.Rows)
        {
            $obj = New-Object -TypeName psobject
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ComputerName' -Value $Row.Name0
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'GUID' -Value $Row.GUID
			Add-Member -InputObject $obj -MemberType NoteProperty -Name 'BuildExt' -Value $Row.BuildExt
            $PCResults += $obj
        }
        
        # For each 'Unknown' computer in the list, get the PC name from MDT
        if ($MDTIntegrated -eq 'true')
        {
            $URL1 = $URL.Replace('Computers','ComputerIdentities')

            # Get ID numbers and Identifiers (GUIDs)
            function GetMDTIDs 
            { 
                param ($URL)
                $Data = Invoke-RestMethod -Uri $URL
                foreach($property in ($Data.content.properties) ) 
                { 
                    New-Object -TypeName PSObject -Property @{
                        ID         = $($property.ID.'#text')
                        Identifier = $($property.Identifier)
                    }
                } 
            }
                
            # Filter out only the GUIDs
            $MDTIDs = GetMDTIDs -URL $URL1 |
            Select-Object -Property * |
            Where-Object -FilterScript {
                $_.Identifier -like '*-*'
            } |
            Sort-Object -Property ID
            $UnknownComputers = $PCResults | Where-Object -FilterScript {
                $_.ComputerName -eq 'Unknown'
            }
            $MDTComputerIDs = @()
            Foreach ($Computer in $UnknownComputers)
            {
                $MDTComputerID = $MDTIDs |
                Where-Object -FilterScript {
                    $_.Identifier -eq $Computer.GUID
                } |
                Select-Object -Property ID 
                $MDTComputerIDs += $MDTComputerID
            }
                
            # Get ComputerNames from MDT
            function GetMDTComputerNames 
            { 
                param ($URL)
                $Data = Invoke-RestMethod -Uri $URL
                foreach($property in ($Data.content.properties) ) 
                { 
                    New-Object -TypeName PSObject -Property @{
                        Name = $($property.Name)
                        ID   = $($property.ID.'#text')
                    } 
                } 
            } 
                
            # Filter out the computer names from the IDs
            $MDTComputers = GetMDTComputerNames -URL $URL |
            Select-Object -Property * |
            Sort-Object -Property ID

            $AdditionalComputerNames = @()
            Foreach ($MDTComputerID in $MDTComputerIDs)
            {
                $MDTComputerName = $MDTComputers |
                Where-Object -FilterScript {
                    $_.ID -eq $MDTComputerID.ID
                } |
                Select-Object -ExpandProperty Name
                $AdditionalComputerNames += $MDTComputerName.ToUpper()
            }
                
            $ConfigMgrList = $PCResults |
            Select-Object -Property ComputerName |
            Where-Object -FilterScript {
                $_.ComputerName -ne 'Unknown'
            }
            $FinalComputerNameList = @()
            $FinalComputerNameList += $ConfigMgrList.ComputerName
            $FinalComputerNameList += $AdditionalComputerNames
			
			$BuildVersions = [String]::Join('; ', (($PCResults | Select-Object -Property BuildExt).BuildExt | Group-Object | Sort-Object Count -Descending | ForEach-Object {$_ | select * } ))
			
        }  
        
        # Add a wildcard option and add only ConfigMgr results if MDT not enabled
        if ($MDTIntegrated -eq $false)
        {
            $FinalComputerNameList = @()
            #$PCResults = $PCResults | Select-Object -ExpandProperty ComputerName
            #$FinalComputerNameList += $PCResults | Select-Object -ExpandProperty ComputerName
			$FinalComputerNameList += $PCResults | Select-Object @{Name='DisplayName'; Expression={$_.ComputerName}}, @{Name='Value'; Expression={$_.GUID}}
			$BuildVersions = [String]::Join('; ', @($PCResults | Select-Object -Property BuildExt | Group-Object BuildExt | Sort-Object Count -Descending | ForEach-Object {  "$($_.Name.Trim()) = $($_.Count)" }) )
		
        }
        #$FinalComputerNameList += '-All-'
		$FinalComputerNameList += [pscustomobject]@{ DisplayName = "-All-"; Value = 0 }
         
  
        # Display results in ComputerName comboxbox     
        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.ComputerName.ItemsSource = [Array]$FinalComputerNameList
				$hash.ComputerName.DisplayMemberPath = "DisplayName"
				$hash.BuildExt.Text = $BuildVersions
        })
    }

    # Set variables from Hash table
    $SQLServer = $hash.SQLServer.Text
    $Database = $hash.Database.Text
    $TimePeriod = $hash.TimePeriod.Text
	$SuccessCode = $hash.SuccessCode.Text
    $ErrorsOnly = $hash.ErrorsOnly.IsChecked
	$DisabledSteps = $hash.DisabledSteps.IsChecked
    $TS = $hash.TaskSequence.SelectedItem
    $MDTIntegrated = $hash.MDTIntegrated.IsChecked
    $URL = $hash.MDTURL.Text

    # Create PS instance in runspace pool and execute
    $PSinstance = [powershell]::Create().AddScript($code).AddArgument($hash).AddArgument($SQLServer).AddArgument($Database).AddArgument($BuildExtVersionSql).AddArgument($TimePeriod).AddArgument($SuccessCode).AddArgument($ErrorsOnly).AddArgument($DisabledSteps).AddArgument($TS).AddArgument($MDTIntegrated).AddArgument($URL)
    $PSInstances += $PSinstance
    $PSinstance.RunspacePool = $RunspacePool
    $PSinstance.BeginInvoke()
}

Function Populate-ActionNames 
{
    param ($hash,$RunspacePool)

    $code = 
    {
        param($hash,$SQLServer,$Database,$TimePeriod,$SuccessCode,$ErrorsOnly,$DisabledSteps,$TS)
        
        # Set variable values
        if ($ErrorsOnly -eq 'True')
        {
            $ExitCode = $SuccessCode
        }
        else 
        {
            $ExitCode = 999999999999999999999999
        }
		
		if ($DisabledSteps -eq 'True')
        {
            $DisabledStep = "''"
        }
        Else 
        {
            $DisabledStep = '11128'
        }

        # Connect to SQL Server
        $connectionString = "Server=$SQLServer;Database=$Database;Integrated Security=SSPI;"
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString
        $connection.Open()

        # Run SQL query
        $Query = "
            Select Distinct ActionName
            from vSMS_TaskSequenceExecutionStatus tes
            --inner join v_R_System sys on tes.ResourceID = sys.ResourceID
            inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
            where tsp.Name = '$TS'
            --and DATEDIFF(hour,ExecutionTime,GETDATE()) <= $TimePeriod
			and DATEDIFF(hour,(CONVERT(datetime, SWITCHOFFSET(CONVERT(datetimeoffset, ExecutionTime), DATENAME(TzOffset, SYSDATETIMEOFFSET()))) ),GETDATE()) <= $TimePeriod
            --and Name0 like '$SQLComputerName'
            and ExitCode not in ($ExitCode)
			and tes.LastStatusMsgID not in ($DisabledStep)
            ORDER BY ActionName ASC
        "
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $result = $command.ExecuteReader()
        $table = New-Object -TypeName 'System.Data.DataTable'
        $table.Load($result)
        $connection.Close()
         
        # Gather results into PS object    
        $PCResults = @()
        Foreach ($Row in $table.Rows)
        {
            $obj = New-Object -TypeName psobject
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ActionName' -Value $Row.ActionName
            $PCResults += $obj
        }
        
        $FinalActionNameList = @()
        $PCResults = $PCResults | Select-Object -ExpandProperty ActionName
        $FinalActionNameList += $PCResults

        $FinalActionNameList += '-All-'
         
  
        # Display results in ComputerName comboxbox     
        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.ActionName.ItemsSource = [Array]$FinalActionNameList
        })
    }

    # Set variables from Hash table
    $SQLServer = $hash.SQLServer.Text
    $Database = $hash.Database.Text
    $TimePeriod = $hash.TimePeriod.Text
	$SuccessCode = $hash.SuccessCode.Text
    $ErrorsOnly = $hash.ErrorsOnly.IsChecked
	$DisabledSteps = $hash.DisabledSteps.IsChecked
    $TS = $hash.TaskSequence.SelectedItem

    # Create PS instance in runspace pool and execute
    $PSinstance = [powershell]::Create().AddScript($code).AddArgument($hash).AddArgument($SQLServer).AddArgument($Database).AddArgument($TimePeriod).AddArgument($SuccessCode).AddArgument($ErrorsOnly).AddArgument($DisabledSteps).AddArgument($TS)
    $PSInstances += $PSinstance
    $PSinstance.RunspacePool = $RunspacePool
    $PSinstance.BeginInvoke()
}

Function Enable-MDT 
{
    #Dispose-PSInstances
    $hash.DeploymentStatus.IsEnabled = 'True'
    $hash.CurrentStep.IsEnabled = 'True'
    $hash.StepName.IsEnabled = 'True'
    $hash.PercentComplete.IsEnabled = 'True'
    $hash.MDTStartTime.IsEnabled = 'True'
    $hash.MDTEndTime.IsEnabled = 'True'
    $hash.MDTElapsedTime.IsEnabled = 'True'

    $hash.DeploymentStatusLabel.IsEnabled = 'True'
    $hash.CurrentStepLabel.IsEnabled = 'True'
    $hash.StepNameLabel.IsEnabled = 'True'
    $hash.PercentCompleteLabel.IsEnabled = 'True'
    $hash.StartLabel.IsEnabled = 'True'
    $hash.EndLabel.IsEnabled = 'True'
    $hash.ElapsedLabel.IsEnabled = 'True'
    $hash.ProgressLabel.IsEnabled = 'True'
}

Function Disable-MDT 
{
    #Dispose-PSInstances
    $hash.DeploymentStatus.IsEnabled = $false
    $hash.CurrentStep.IsEnabled = $false
    $hash.StepName.IsEnabled = $false
    $hash.PercentComplete.IsEnabled = $false
    $hash.MDTStartTime.IsEnabled = $false
    $hash.MDTEndTime.IsEnabled = $false
    $hash.MDTElapsedTime.IsEnabled = $false

    $hash.DeploymentStatusLabel.IsEnabled = $false
    $hash.CurrentStepLabel.IsEnabled = $false
    $hash.StepNameLabel.IsEnabled = $false
    $hash.PercentCompleteLabel.IsEnabled = $false
    $hash.StartLabel.IsEnabled = $false
    $hash.EndLabel.IsEnabled = $false
    $hash.ElapsedLabel.IsEnabled = $false
    $hash.ProgressLabel.IsEnabled = $false
}

Function Get-MDTData 
{
    param ($hash,$RunspacePool)

    $code = 
    {
        param($hash,$URL,$ComputerName,$IsMDTIntegrated,$DTFormat)

        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.DeploymentStatus.Text = ''
                $hash.CurrentStep.Text = ''
                $hash.StepName.Text = ''
                $hash.PercentComplete.Text = ''
                $hash.MDTStartTime.Text = ''
                $hash.MDTEndTime.Text = ''
                $hash.MDTElapsedTime.Text = ''
                $hash.ProgressBar.Value = 0
        })

        if ($IsMDTIntegrated -eq $true -and $ComputerName -ne '-All-' -and $ComputerName -ne '' -and $ComputerName -ne $Null)
        {
            function GetMDTData2 
            { 
                param ($URL)
                $Data = Invoke-RestMethod -Uri $URL
        
                foreach($property in ($Data.content.properties) ) 
                { 
                    New-Object -TypeName PSObject -Property @{
                        Name             = $($property.Name)
                        PercentComplete  = $($property.PercentComplete.'#text')
                        CurrentStep      = $($property.CurrentStep.'#text')
                        StepName         = $($property.StepName)
                        Warnings         = $($property.Warnings.'#text')
                        Errors           = $($property.Errors.'#text')
                        DeploymentStatus = $( 
                            Switch ($property.DeploymentStatus.'#text') { 
                                1 
                                {
                                    'Active/Running'
                                } 
                                2 
                                {
                                    'Failed'
                                } 
                                3 
                                {
                                    'Successfully completed'
                                } 
                                Default 
                                {
                                    'Unknown'
                                } 
                            } 
                        )
                        StartTime        = $($property.StartTime.'#text') -replace 'T', ' '
                        EndTime          = $($property.EndTime.'#text') -replace 'T', ' '
                    } 
                } 
            } 
            try 
            {
                $MDT = GetMDTData2 -URL $URL | 
                Select-Object -Property Name, DeploymentStatus, PercentComplete, CurrentStep, StepName, Warnings, Errors, StartTime, EndTime | 
                Sort-Object -Property Name | 
                Where-Object -FilterScript {
                    $_.Name -eq $ComputerName
                }
                        
                if ($MDT)
                {
                    # Calculate times
                    $MDTServer = $URL.Split('//')[2].Split(':')[0]
                    $Start = $MDT.StartTime | Get-Date
                    if ($DTFormat -ne 'UTC')
                    {
                        $Start = [System.TimeZone]::CurrentTimeZone.ToLocalTime($Start)
                    }
                    if (!$MDT.EndTime)
                    {
                        $MDTDate = Invoke-Command -ComputerName $MDTServer -ScriptBlock {
                            (Get-Date).ToUniversalTime()
                        }
                        if ($DTFormat -ne 'UTC')
                        {
                            $MDTDate = [System.TimeZone]::CurrentTimeZone.ToLocalTime($MDTDate)
                        }
                        $Elapsed = $MDTDate - $Start
                        $Elapsed = "$($Elapsed.Hours)h $($Elapsed.Minutes)m $($Elapsed.Seconds)s"
                        $Elapsed
                    }
                    if ($MDT.EndTime)
                    {
                        $End = $MDT.EndTime | Get-Date
                        if ($DTFormat -ne 'UTC')
                        {
                            $End = [System.TimeZone]::CurrentTimeZone.ToLocalTime($End)
                        }
                        $Elapsed = $End - $Start
                        $Elapsed = "$($Elapsed.Hours)h $($Elapsed.Minutes)m $($Elapsed.Seconds)s"
                        $Elapsed
                    }

                    $hash.Window.Dispatcher.Invoke(
                        [action]{
                            $hash.DeploymentStatus.Text = $MDT.DeploymentStatus
                            $hash.CurrentStep.Text = $MDT.CurrentStep
                            $hash.StepName.Text = $MDT.StepName
                            $hash.PercentComplete.Text = $MDT.PercentComplete
                            $hash.ProgressBar.Value = $MDT.PercentComplete
                            $hash.MDTStartTime.Text = $Start
                            $hash.MDTEndTime.Text = $End
                            $hash.MDTElapsedTime.Text = $Elapsed
                    })
                }
                Else 
                {
                    $hash.Window.Dispatcher.Invoke(
                        [action]{
                            $hash.DeploymentStatus.Text = 'No data found'
                    })
                }
            }
            catch
            {
                $hash.Window.Dispatcher.Invoke(
                    [action]{
                        $hash.ActionOutput.Text = '[ERROR] Could not connect to MDT Web Service'
                })
            }
        }
    }

    # Set variables from Hash table
    $ComputerName = $hash.ComputerName.SelectedItem
    $IsMDTIntegrated = $hash.MDTIntegrated.IsChecked
    $URL = $hash.MDTURL.Text
    $DTFormat = $hash.DTFormat.SelectedItem

    # Create PS instance in runspace pool and execute
    $PSinstance = [powershell]::Create().AddScript($code).AddArgument($hash).AddArgument($URL).AddArgument($ComputerName).AddArgument($IsMDTIntegrated).AddArgument($DTFormat)
    $PSInstances += $PSinstance
    $PSinstance.RunspacePool = $RunspacePool
    $PSinstance.BeginInvoke()
}

function Dispose-PSInstances 
{
    foreach ($PSinstance in $PSInstances)
    {
        if ($PSinstance.InvocationStateInfo.State -eq 'Completed')
        {
            $PSinstance.Dispose()
        }
    }
}

Function Create-Timer 
{
    $global:Timer = New-Object -TypeName System.Windows.Forms.Timer
    $timer.Interval = [int]$hash.RefreshPeriod.Text * 60000
}

Function Start-Timer 
{
    if ($timer)
    {
        $timer.Start()
    }
}

Function Stop-Timer 
{
    if ($timer)
    {
        $timer.Stop()
    }
}

Function Update-Registry 
{
    param($hash)
    # Test whether running as admin first
    If (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator'))
    {
        if (Test-Path -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor')
        {
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor' -Name SQLServer -Value $hash.SQLServer.Text
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor' -Name Database -Value $hash.Database.Text
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor' -Name MDTURL -Value $hash.MDTURL.Text
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor' -Name DTFormat -Value $hash.DTFormat.SelectedItem
        }

        if (Test-Path -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor')
        {
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name SQLServer -Value $hash.SQLServer.Text
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name Database -Value $hash.Database.Text
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name MDTURL -Value $hash.MDTURL.Text
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name DTFormat -Value $hash.DTFormat.SelectedItem
        }
    }
}

Function Read-Registry 
{
    if (Test-Path -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor')
    {
        $regsql = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor' -Name SQLServer | Select-Object -ExpandProperty SQLServer
        $regdb = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor' -Name Database | Select-Object -ExpandProperty Database
        $regmdt = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor' -Name MDTURL | Select-Object -ExpandProperty MDTURL
        $regdtformat = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor' -Name DTFormat | Select-Object -ExpandProperty DTFormat
    }

    if (Test-Path -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor')
    {
        $regsql = Get-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name SQLServer | Select-Object -ExpandProperty SQLServer
        $regdb = Get-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name Database | Select-Object -ExpandProperty Database
        $regmdt = Get-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name MDTURL | Select-Object -ExpandProperty MDTURL
        $regdtformat = Get-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name DTFormat | Select-Object -ExpandProperty DTFormat
    }
    if ($regsql -ne $Null -and $regsql -ne '')
    {
        $hash.SQLServer.Text = $regsql
    }
    if ($regdb -ne $Null -and $regdb -ne '')
    {
        $hash.Database.Text = $regdb
    }
    if ($regmdt -ne $Null -and $regmdt -ne '')
    {
        $hash.MDTURL.Text = $regmdt
    }

    if ($regdtformat -ne $Null -and $regdtformat -ne '')
    {
        if ($regdtformat -eq 'UTC')
        {
            if (!$CurrentDateTimeF)
            {
                $Global:CurrentDateTimeF = 'UTC'
            }
        }
    }
}

Function Update-ConfigFile 
{
    param($hash)
    $XML_Config = "Config.xml"
    If(Test-Path $XML_Config){
        [xml]$Get_Config = get-content $XML_Config

        $Get_Config.Config.sql = $hash.SQLServer.Text
        $Get_Config.Config.db = $hash.Database.Text
		$Get_Config.Config.mdt = $hash.MDT.IsChecked.toString()
        $Get_Config.Config.mdturl = $hash.MDTURL.Text
        $Get_Config.Config.dtformat = $hash.DTFormat.SelectedItem
        
        $Get_Config.Save((Resolve-Path $XML_Config))

    }
}

Function Read-ConfigFile
{
	$XML_Config = "Config.xml"
	If(Test-Path $XML_Config){
		[xml]$Get_Config = get-content $XML_Config
		
		$regsql = $Get_Config.Config.sql
        $regdb = $Get_Config.Config.db
		$regmdt = if($Get_Config.Config.mdt -eq $true -or $Get_Config.Config.mdt -eq "true"){$true} else { $false }
        $regmdturl = $Get_Config.Config.mdturl
        $regdtformat = $Get_Config.Config.dtformat
	}
	    if ($regsql -ne $Null -and $regsql -ne '')
    {
        $hash.SQLServer.Text = $regsql
    }
    if ($regdb -ne $Null -and $regdb -ne '')
    {
        $hash.Database.Text = $regdb
    }
	if (![string]::IsNullOrEmpty($regmdt))
    {
        $hash.MDT.isChecked = $regmdt
    }
    if ($regmdturl -ne $Null -and $regmdturl -ne '')
    {
        $hash.MDTURL.Text = $regmdturl
    }

    if ($regdtformat -ne $Null -and $regdtformat -ne '')
    {
        if ($regdtformat -eq 'UTC')
        {
            if (!$CurrentDateTimeF)
            {
                $Global:CurrentDateTimeF = 'UTC'
            }
        }
    }
	
}

Function Update-MdtView {
	if($hash.mdt.isChecked){
		$hash.MDTGroupBox.Visibility = "Visible" 
	}
	else {
		$hash.MDTGroupBox.Visibility = "Collapsed"
	}
}

Function Generate-Report 
{
    param ($hash,$RunspacePool)

    $code = 
    {
        param($hash,$SQLServer,$Database,$StartDate,$EndDate,$TS,$DTFormat)
        $Results = @()

        if ($DTFormat -ne 'UTC')
        {
            [datetime]$StartDate = $StartDate.ToUniversalTime()
            [datetime]$EndDate = $EndDate.ToUniversalTime()
        }

        # Set dates to ISO standard format for SQL Server
        $SQLStart = $StartDate | Get-Date -Format s
        $SQLEnd = $EndDate | Get-Date -Format s

        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.Working.Content = 'Working...'
                $hash.ReportProgress.Visibility = 'Visible'
                $hash.ReportProgress.Value = 10
        })

        $connectionString = "Server=$SQLServer;Database=$Database;Integrated Security=SSPI;"
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString
        $connection.Open()

        # Find all resourceID for TS steps between the selected dates
        $Query = "
            select distinct tes.ResourceID
            from vSMS_TaskSequenceExecutionStatus tes
            --inner join v_R_System sys on tes.ResourceID = sys.ResourceID
            inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
            where tsp.Name = '$TS'
            and tes.ExecutionTime >= '$SQLStart'
            and tes.ExecutionTime <= '$SQLEnd'
        "

        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $reader = $command.ExecuteReader()
        $table = New-Object -TypeName 'System.Data.DataTable'
        $table.Load($reader)

        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.ReportProgress.Value = 20
        })

        foreach ($ResourceID in $table.Rows.ResourceID)
        {
            $Query = "
                Select (select top(1) convert(datetime,ExecutionTime,121)
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                and tes.ExecutionTime >= '$SQLStart' 
                and tes.ExecutionTime <= '$SQLEnd'
                and LastStatusMsgName = 'The task sequence execution engine started execution of a task sequence'
                and Step = 0
                and tes.ResourceID = $ResourceID
                order by ExecutionTime desc) as 'Start',
                (select top(1) convert(datetime,ExecutionTime,121)
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                and tes.ExecutionTime >= '$SQLStart'
                and tes.ExecutionTime <= '$SQLEnd'
                and LastStatusMsgName = 'The task sequence execution engine successfully completed a task sequence'
                and tes.ResourceID = $ResourceID
                order by ExecutionTime desc) as 'Finish',
                (Select name0 from v_R_System sys where sys.ResourceID = $ResourceID) as 'ComputerName',
                (select Model0 from v_GS_Computer_System comp where comp.ResourceID = $ResourceID) as 'Model'
            "
            $command = $connection.CreateCommand()
            $command.CommandText = $Query
            $reader = $command.ExecuteReader()
            $table = New-Object -TypeName 'System.Data.DataTable'
            $table.Load($reader)


            if ($table.rows[0].Start.GetType().Name -eq 'DBNull')
            {
                $Start = ''
            }
            Else 
            {
                if ($DTFormat -eq 'UTC')
                {
                    $Start = $table.rows[0].Start
                }
                Else 
                {
                    $Start = [System.TimeZone]::CurrentTimeZone.ToLocalTime($($table.rows[0].Start | Get-Date))
                }
            }

            if ($table.rows[0].Finish.GetType().Name -eq 'DBNull')
            {
                $Finish = ''
            }
            Else 
            {
                if ($DTFormat -eq 'UTC')
                {
                    $Finish = $table.rows[0].Finish
                }
                Else 
                {
                    $Finish = [System.TimeZone]::CurrentTimeZone.ToLocalTime($($table.rows[0].Finish | Get-Date))
                }
            }


            #$table
            if ($Start -eq '' -or $Finish -eq '')
            {
                $diff = $Null
            }
            else 
            {
                $diff = $Finish-$Start
            }


            $PC = New-Object -TypeName psobject
            Add-Member -InputObject $PC -MemberType NoteProperty -Name ComputerName -Value $table.rows[0].ComputerName
            Add-Member -InputObject $PC -MemberType NoteProperty -Name StartTime -Value $Start
            Add-Member -InputObject $PC -MemberType NoteProperty -Name FinishTime -Value $Finish
            if ($Start -eq '' -or $Finish -eq '')
            {
                Add-Member -InputObject $PC -MemberType NoteProperty -Name DeploymentTime -Value ''
            }
            else
            {
                Add-Member -InputObject $PC -MemberType NoteProperty -Name DeploymentTime -Value $("$($diff.hours)" + ' hours ' + "$($diff.minutes)" + ' minutes')
            }
            Add-Member -InputObject $PC -MemberType NoteProperty -Name Model -Value $table.rows[0].Model
            $Results += $PC
        }

        $Results = $Results | Sort-Object -Property ComputerName

        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.ReportProgress.Value = 50
        })

        $Query = "
            select sys.Name0 as 'ComputerName',
            tsp.Name 'Task Sequence',
            comp.Model0 as Model,
            tes.ExecutionTime,
            tes.Step,
            tes.GroupName,
            tes.ActionName,
            tes.LastStatusMsgName,
            tes.ExitCode,
            tes.ActionOutput
            from vSMS_TaskSequenceExecutionStatus tes
            left join v_R_System sys on tes.ResourceID = sys.ResourceID
            left join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
            left join v_GS_COMPUTER_SYSTEM comp on tes.ResourceID = comp.ResourceID
            where tsp.Name = '$TS'
            and tes.ExecutionTime >= '$SQLStart'
            and tes.ExecutionTime <= '$SQLEnd'
            and tes.ExitCode not in (0,-2147467259)
            Order by tes.ExecutionTime desc
        "

        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $reader = $command.ExecuteReader()
        $table = New-Object -TypeName 'System.Data.DataTable'
        $table.Load($reader)

        if ($DTFormat -ne 'UTC')
        {
            $newdates = foreach ($item in $table.rows.ExecutionTime)
            {
                [System.TimeZone]::CurrentTimeZone.ToLocalTime($item)
            }
            $i = -1
            $table.rows.ExecutionTime | ForEach-Object -Process {
                $i ++
                $table.Rows[$i].ExecutionTime = $newdates[$i]
            }
        }

        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.ReportProgress.Value = 80
        })

        #Convert dates if necessary
        if ($DTFormat -ne 'UTC')
        {
            $StartDate = [System.TimeZone]::CurrentTimeZone.ToLocalTime($StartDate)
            $EndDate = [System.TimeZone]::CurrentTimeZone.ToLocalTime($EndDate)
        }         

        # Create html email
        $style = @"
<style>
body {
    color:#012E34;
    font-family:Calibri,Tahoma;
    font-size: 10pt;
}
h1 {
    text-align:center;
}
h2 {
    border-top:1px solid #666666;
}
 
 
th {
    font-weight:bold;
    color:#012E34;
    background-color:#69969C;
}
.odd  { background-color:#012E34; }
.even { background-color:#012E34; }
</style>
"@


        $HEaders = @"
<H1>Task Sequence Execution Summary Report</H1>
<H3>Starting Date: $StartDate</H3>
<H3>End Date: $EndDate</H3>
<H3>Task Sequence: $TS</H3>
<H3>TimeZone for Date/Time: $DTFormat</H3>
"@

        $body1 = $Results | 
        Select-Object -Property ComputerName, StartTime, FinishTime , DeploymentTime, Model |
        ConvertTo-Html -Head $style -Body "<H2>Task Sequence Executions ($($Results.Count))</H2>" | 
        Out-String

        $body2 = $table | 
        Select-Object -Property ComputerName, 'Task Sequence', Model, ExecutionTime, Step, GroupName, ActionName, LastStatusMsgName, ExitCode |
        ConvertTo-Html -Head $style -Body "<H2>Task Sequence Execution Errors ($($table.Rows.Count))</H2>" | 
        Out-String

        $Body = $HEaders + $body1 + $body2

        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.Working.Content = ''
                $hash.ReportProgress.Value = 100
        })

        $Body | Out-File -FilePath $env:temp\TSReport.htm -Force
        Invoke-Item -Path $env:temp\TSReport.htm


        # Close the connection
        $connection.Close()
    }

    # Set variables from Hash table
    $SQLServer = $hash.SQLServer.Text
    $Database = $hash.Database.Text
    $DTFormat = $hash.DTFormat.SelectedItem
    #if ($DTFormat -eq "UTC")
    #   {
    [datetime]$StartDate = $hash.StartDate.Text | Get-Date -Format "MM'/'dd'/'yyyy HH':'mm':'ss"
    [datetime]$EndDate = $hash.EndDate.Text | Get-Date -Format "MM'/'dd'/'yyyy HH':'mm':'ss"
    #  }
    #Else {
    #       [datetime]$StartDate = [System.TimeZone]::CurrentTimeZone.ToLocalTime($($hash.StartDate.Text | Get-Date -Format "MM'/'dd'/'yyyy HH':'mm':'ss"))
    #      [datetime]$EndDate = [System.TimeZone]::CurrentTimeZone.ToLocalTime($($hash.EndDate.Text | Get-Date -Format "MM'/'dd'/'yyyy HH':'mm':'ss"))
    # }

    $EndDate = $EndDate.AddDays(1).AddSeconds(-1)
    $TS = $hash.TSList.SelectedItem

    # Create PS instance in runspace pool and execute
    $PSinstance = [powershell]::Create().AddScript($code).AddArgument($hash).AddArgument($SQLServer).AddArgument($Database).AddArgument($StartDate).AddArgument($EndDate).AddArgument($TS).AddArgument($DTFormat)
    $PSInstances += $PSinstance
    $PSinstance.RunspacePool = $RunspacePool
    $PSinstance.BeginInvoke()
}

Function Clear-MDT 
{
    param ($hash)

    $hash.Window.Dispatcher.Invoke(
        [action]{
            $hash.DeploymentStatus.Text = ''
            $hash.CurrentStep.Text = ''
            $hash.StepName.Text = ''
            $hash.PercentComplete.Text = ''
            $hash.ProgressBar.Value = 0
            $hash.MDTStartTime.Text = ''
            $hash.MDTEndTime.Text = ''
            $hash.MDTElapsedTime.Text = ''
    })
}

#endregion

#region Event Handlers

$hash.Window.Add_ContentRendered({
        #Disable-MDT
        #Read-Registry
		Read-ConfigFile
		Update-MdtView
        Get-DateTimeFormat
        Get-TaskSequenceList
})

$hash.TaskSequence.Add_SelectionChanged({
        $Count = $hash.ComputerName.Items.Count
        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.ComputerName.SelectedIndex = ($Count -1)
        })
        Dispose-PSInstances
        Clear-MDT -hash $hash
        Populate-ComputerNames -hash $hash -RunspacePool $RunspacePool
		Populate-ActionNames -hash $hash -RunspacePool $RunspacePool
        Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
        Get-MDTData -hash $hash -RunspacePool $RunspacePool
        Stop-Timer
        Create-Timer
        $timer.add_Tick({
                Populate-ComputerNames -hash $hash -RunspacePool $RunspacePool
				Populate-ActionNames -hash $hash -RunspacePool $RunspacePool
                Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
                Get-MDTData -hash $hash -RunspacePool $RunspacePool
        })
        Start-Timer
        $Global:CurrentTS = $hash.TaskSequence.SelectedItem
})

$hash.ErrorsOnly.Add_Checked({
        Dispose-PSInstances
        Stop-Timer
        Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
        Start-Timer
})

$hash.DisabledSteps.Add_Checked({
        Dispose-PSInstances
        Stop-Timer
        Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
        Start-Timer
})

$hash.MDTIntegrated.Add_Checked({
        Populate-ComputerNames -hash $hash -RunspacePool $RunspacePool
        Enable-MDT
})

$hash.MDTIntegrated.Add_Unchecked({
        Populate-ComputerNames -hash $hash -RunspacePool $RunspacePool
        Disable-MDT
})

$hash.ErrorsOnly.Add_Unchecked({
        Dispose-PSInstances
        Stop-Timer
        Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
        Start-Timer
})

$hash.DisabledSteps.Add_Unchecked({
        Dispose-PSInstances
        Stop-Timer
        Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
        Start-Timer
})

$hash.DataGrid.Add_SelectionChanged({
        Dispose-PSInstances
        Populate-ActionOutput -hash $hash -RunspacePool $RunspacePool
})

$hash.RefreshNow.Add_Click({
        Dispose-PSInstances
        Stop-Timer
        Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
        Populate-ComputerNames -hash $hash -RunspacePool $RunspacePool
		Populate-ActionNames -hash $hash -RunspacePool $RunspacePool
        Get-MDTData -hash $hash -RunspacePool $RunspacePool
        $timer.Interval = [int]$hash.RefreshPeriod.Text * 60000
        Start-Timer
})

$hash.ComputerName.Add_SelectionChanged({
        if ($hash.TaskSequence.SelectedItem -eq $CurrentTS)
        {
            Dispose-PSInstances
            Stop-Timer
            Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
            Get-MDTData -hash $hash -RunspacePool $RunspacePool
            Start-Timer
        }
})

$hash.ActionName.Add_SelectionChanged({
        if ($hash.TaskSequence.SelectedItem -eq $CurrentTS)
        {
            Dispose-PSInstances
            Stop-Timer
            Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
            Get-MDTData -hash $hash -RunspacePool $RunspacePool
            Start-Timer
        }
})

$hash.TimePeriod.Add_KeyDown({
        if ($_.Key -eq 'Return')
        {
            Dispose-PSInstances
            Stop-Timer
            Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
            Populate-ComputerNames -hash $hash -RunspacePool $RunspacePool
            Get-MDTData -hash $hash -RunspacePool $RunspacePool
            $timer.Interval = [int]$hash.RefreshPeriod.Text * 60000
            Start-Timer
        }
})

$hash.RefreshPeriod.Add_TextChanged({
        Stop-Timer
        $timer.Interval = [int]$hash.RefreshPeriod.Text * 60000
        Start-Timer
})

$hash.SettingsButton.Add_Click({
        $reader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml2)
        $hash.Window2 = [Windows.Markup.XamlReader]::Load( $reader )
        $hash.SQLServer = $hash.Window2.FindName('SQLServer')
        $hash.Database = $hash.Window2.FindName('Database')
        $hash.MDT = $hash.Window2.FindName('MDT')
		$hash.MDTURL = $hash.Window2.FindName('MDTURL')
        $hash.ConnectSQL = $hash.Window2.FindName('ConnectSQL')
        $hash.TSList = $hash.Window2.FindName('TSList')
        $hash.StartDate = $hash.Window2.FindName('StartDate')
        $hash.EndDate = $hash.Window2.FindName('EndDate')
        $hash.GenerateReport = $hash.Window2.FindName('GenerateReport')
        $hash.SettingsTab = $hash.Window2.FindName('SettingsTab')
        $hash.ReportTab = $hash.Window2.FindName('ReportTab')
        $hash.Tabs = $hash.Window2.FindName('Tabs')
        $hash.Runasadmin = $hash.Window2.FindName('Runasadmin')
        $hash.Working = $hash.Window2.FindName('Working')
        $hash.ReportProgress = $hash.Window2.FindName('ReportProgress')
        $hash.Link1 = $hash.Window2.FindName('Link1')
        $hash.DTFormat = $hash.Window2.FindName('DTFormat')
        if (Test-Path -Path "$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico")
        {
            #$hash.Window2.Icon = "$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico"
        }
        if (Test-Path -Path "${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico")
        {
            #$hash.Window2.Icon = "${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico"
        }

        #Read-Registry
		Read-ConfigFile
		Update-MdtView
        $hash.SettingsTab.Focus()
        If (!(([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')))
        {
            $hash.Runasadmin.Visibility = 'Visible'
        }
        $hash.TSList.ItemsSource = $Views.TS
        $hash.DTFormat.ItemsSource = $Timezones.TimeZone
        If ($CurrentDateTimeF -eq 'UTC')
        {
            $hash.DTFormat.SelectedIndex = 0
        }
        Else 
        {
            $hash.DTFormat.SelectedIndex = 1
        }

        $hash.SQLServer.Add_GotMouseCapture({
                if ($hash.SQLServer.Text -eq '<SQLServer\Instance>')
                {
                    $hash.SQLServer.Text = ''
                }
        })

        $hash.SQLServer.Add_GotKeyboardFocus({
                if ($hash.SQLServer.Text -eq '<SQLServer\Instance>')
                {
                    $hash.SQLServer.Text = ''
                }
        })

        $hash.Database.Add_GotMouseCapture({
                if ($hash.Database.Text -eq '<Database>')
                {
                    $hash.Database.Text = ''
                }
        })

        $hash.Database.Add_GotKeyboardFocus({
                if ($hash.Database.Text -eq '<Database>')
                {
                    $hash.Database.Text = ''
                }
        })
        $hash.ConnectSQL.Add_Click({
                #Update-Registry -hash $hash
				Update-ConfigFile -hash $hash
				Update-MdtView
                Get-TaskSequenceList
        })

        $hash.GenerateReport.Add_Click({
                Generate-Report -hash $hash -RunspacePool $RunspacePool
        })

        $hash.Link1.Add_Click({
                Start-Process -FilePath 'http://smsagent.wordpress.com/tools/configmgr-task-sequence-monitor/'
        })

        $hash.DTFormat.Add_SelectionChanged({
                $Global:CurrentDateTimeF = $hash.DTFormat.SelectedItem
        })

        $hash.Window2.Add_Closed({
                #Update-Registry -hash $hash
				Update-ConfigFile -hash $hash
				Update-MdtView
        })

        $Null = $hash.Window2.ShowDialog()
})

$hash.ReportButton.Add_Click({
        $reader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml2)
        $hash.Window2 = [Windows.Markup.XamlReader]::Load( $reader )
        $hash.SQLServer = $hash.Window2.FindName('SQLServer')
        $hash.Database = $hash.Window2.FindName('Database')
		$hash.MDT = $hash.Window2.FindName('MDT')
        $hash.MDTURL = $hash.Window2.FindName('MDTURL')
        $hash.ConnectSQL = $hash.Window2.FindName('ConnectSQL')
        $hash.TSList = $hash.Window2.FindName('TSList')
        $hash.StartDate = $hash.Window2.FindName('StartDate')
        $hash.EndDate = $hash.Window2.FindName('EndDate')
        $hash.GenerateReport = $hash.Window2.FindName('GenerateReport')
        $hash.SettingsTab = $hash.Window2.FindName('SettingsTab')
        $hash.ReportTab = $hash.Window2.FindName('ReportTab')
        $hash.Tabs = $hash.Window2.FindName('Tabs')
        $hash.Runasadmin = $hash.Window2.FindName('Runasadmin')
        $hash.Working = $hash.Window2.FindName('Working')
        $hash.ReportProgress = $hash.Window2.FindName('ReportProgress')
        $hash.Link1 = $hash.Window2.FindName('Link1')
        $hash.DTFormat = $hash.Window2.FindName('DTFormat')
        if (Test-Path -Path "$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico")
        {
            #$hash.Window2.Icon = "$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico"
        }
        if (Test-Path -Path "${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico")
        {
            #$hash.Window2.Icon = "${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico"
        }

        #Read-Registry
		Read-ConfigFile
		Update-MdtView
        $hash.ReportTab.Focus()
        If (!(([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')))
        {
            $hash.Runasadmin.Visibility = 'Visible'
        }
        $hash.TSList.ItemsSource = $Views.TS
        $hash.DTFormat.ItemsSource = $Timezones.TimeZone
        If ($CurrentDateTimeF -eq 'UTC')
        {
            $hash.DTFormat.SelectedIndex = 0
        }
        Else 
        {
            $hash.DTFormat.SelectedIndex = 1
        }

        $hash.SQLServer.Add_GotMouseCapture({
                if ($hash.SQLServer.Text -eq '<SQLServer\Instance>')
                {
                    $hash.SQLServer.Text = ''
                }
        })

        $hash.SQLServer.Add_GotKeyboardFocus({
                if ($hash.SQLServer.Text -eq '<SQLServer\Instance>')
                {
                    $hash.SQLServer.Text = ''
                }
        })

        $hash.Database.Add_GotMouseCapture({
                if ($hash.Database.Text -eq '<Database>')
                {
                    $hash.Database.Text = ''
                }
        })

        $hash.Database.Add_GotKeyboardFocus({
                if ($hash.Database.Text -eq '<Database>')
                {
                    $hash.Database.Text = ''
                }
        })

        $hash.ConnectSQL.Add_Click({
                #Update-Registry -hash $hash
				Update-ConfigFile -hash $hash
				Update-MdtView
                Get-TaskSequenceList
        })

        $hash.GenerateReport.Add_Click({
                Generate-Report -hash $hash -RunspacePool $RunspacePool
        })

        $hash.Link1.Add_Click({
                Start-Process -FilePath 'http://smsagent.wordpress.com/tools/configmgr-task-sequence-monitor/'
        })

        $hash.DTFormat.Add_SelectionChanged({
                $Global:CurrentDateTimeF = $hash.DTFormat.SelectedItem
        })

        $hash.Window2.Add_Closed({
                #Update-Registry -hash $hash
				Update-ConfigFile -hash $hash
				Update-MdtView
        })

        $Null = $hash.Window2.ShowDialog()
})

$hash.Window.Add_Closed({
        Stop-Timer
        Dispose-PSInstances
        $RunspacePool.close()
        $RunspacePool.Dispose()
})

# Stop process on closing, #comment our for development
$hash.window.Add_Closing({[System.Windows.Forms.Application]::Exit(); Stop-Process $pid})
#endregion


# Make PowerShell Disappear #comment our for development
if($debug){
	$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);' 
	$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru 
	$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)
}

#$app = [Windows.Application]::new()
$app = New-Object Windows.Application
$app.Run($Hash.Window)

