<#                                  Functions                             #>
############################################################################

####       ##
####Info Gathering##

function Get-OS                     
{
<#
.SYNOPSIS

Gets operating System based on WMI caption

.DESCRIPTION

Gets operating System based on WMI caption and returns the value

.INPUTS

None

.OUTPUTS

System.String

.Example

PS> Get-OS
Microsoft Windows 10 Enterprise

#>

    $OperatingSystem = (gwmi win32_operatingsystem).caption
    return $OperatingSystem
}

function Get-Arch                   
{
<#
.SYNOPSIS

Gets current architecture based on IntPtr size

.DESCRIPTION

Gets current architecture based on IntPtr size
Returns the value '32 bit' or '64 bit' based on the [System.Intptr]::size value

.INPUTS

None

.OUTPUTS

System.String

.Example

PS> Get-Arch
32 bit
#>
    $ptrsize = [System.Intptr]::size
    if ($ptrsize -eq "8")
    {
        
        return "64 bit"
    }
    else
    {
        
        return "32 bit"
    }
}

function Set-ProgramFiles32bit     
                                    
{
<#
.SYNOPSIS

Returns file path for 'Program Files' or 'Program Files (x86)' (Requires Get-Arch function)

.DESCRIPTION

Returns file path for 'Program Files' or 'Program Files (x86)' (Requires Get-Arch function)
Sets file path for 32 bit programs to install to (defaults $env:Drive\Program Files for 32 bit 
and 64 $env:SystemRoot\Program Files (x86) for other architectures

.INPUTS

None

.OUTPUTS

System.String

.Example

PS> Set-ProgramFiles32bit
C:\Program Files

.Example

PS> Set-ProgramFiles32bit
C:\Program Files (x86)

#>

    $Arch = Get-Arch
    if ($Arch-eq "32 bit")
    {
        $ProgramFiles = "$env:SystemDrive\Program Files"
    }
    else
    {
        $ProgramFiles = "$env:SystemDrive\Program Files (x86)"
    }
    
    return $ProgramFiles
}

function Set-ProgramFiles64bit      
{
<#
.SYNOPSIS

Returns file path for 'Program Files' (Requires Get-Arch function)

.DESCRIPTION

Returns file path for 'Program Files' (Requires Get-Arch function)
Sets file path for 64 bit programs to install to 

.INPUTS

None

.OUTPUTS

System.String

.Example

PS> Set-ProgramFiles64bit
C:\Program Files
#>

    $Arch = Get-Arch
    $ProgramFiles = "$env:SystemDrive\Program Files"
    return $ProgramFiles
}


function Filter-Program             
{
<#
.SYNOPSIS

Used to filter GUIDs based on program name. (Requires Get-ProgramGUIDs function)


.DESCRIPTION

Used to filter GUIDs based on program name. (Requires Get-ProgramGUIDs function)
This function can output data on the console or through gridview.
No parameters will accept a named program in quotes


.INPUTS

None.

.OUTPUTS
System.Array. Filter-Program returns an array with Names, GUIDs, and MSI Location for
the system, or based on the < FilteredProg > parameter.
Setting < GUI > to "Yes" will output information to Gridview

.Example

PS> Filter-Program "skype"
Processing Program Data:


        Name: Skype Meetings App
        GUID: {3C66E71B-4243-46FF-8C4A-29A926968FBC}
        MSI Location: C:\Windows\Installer\3ffe264.msi

.Example

PS > Filter-Program -FilteredProg "skype"
Processing Program Data:


        Name: Skype Meetings App
        GUID: {3C66E71B-4243-46FF-8C4A-29A926968FBC}
        MSI Location: C:\Windows\Installer\3ffe264.msi

.Example

PS > Filter-Program -FilteredProg "skype" -GUID "Yes"
This will show a GridView window of information on the skype package

.Link

Get-ProgramGUIDs
#>
    param (
        [string]$FilteredProg,
        [string]$GUI="No"
        )
    
    if ($GUI -eq "No")
    {
        $progra = Get-WmiObject win32_product -Filter "Name LIKE '%$FilteredProg%'"|Get-ProgramGUIDs -Verbose
        Write-Output $progra
    }

    if ($GUI -eq "Yes")
    {
        $Grid =[Ordered]@{
        Filter = "Name LIKE '%$FilteredProg%'"
        Class = "win32_product"
        }
        Get-WmiObject @Grid | Out-GridView
    }


        
}



function Get-ProgramGUIDs           
{
<#
.SYNOPSIS

Can be used alone to filter items from get-wmiobject pipeline
Used primarily in conjunction with Filter-Program.

.DESCRIPTION

Can be used alone to filter items from get-wmiobject pipeline
Used primarily in conjunction with Filter-Program. Returns values
of programs based on Name, IdentifyingNumber, and LocalPackage from
WMI tables.

.INPUTS

Accepts output from getting WMI\win32_product table.

.OUTPUTS

System.Array. Object. Returns values from piped WMI information from
win32_product table

.Example

PS> get-wmiobject win32_product |get-programguids
Processing Program Data:

        Name: Skype Meetings App
        GUID: {3C66E71B-4243-46FF-8C4A-29A926968FBC}
        MSI Location: C:\Windows\Installer\3ffe264.msi

        ... (all programs, GUIDs, and MSI locations)

.Link

Filter-Program 
#>
   [cmdletbinding()]
    Param (
        [parameter(ValueFromPipelineByPropertyName=$True)]
        [string[]]$Name,
        [parameter(ValueFromPipelineByPropertyName=$True)]
        [string[]]$IdentifyingNumber,
        [parameter(ValueFromPipelineByPropertyName=$True)]
        [string[]]$LocalPackage
        
        
    )
    
        Begin 
        {
            Write-Output "Processing Program Data:`n"
        }

        Process 
        {
        
        Write-Output "
        Name: $Name
        GUID: $IdentifyingNumber
        MSI Location: $LocalPackage
        "
        }

        End 
        {
            $Report
        }
  
}


Function Get-MSIs
{
<#
.SYNOPSIS

Gets MSI paths from current folder (with no parameters)
Can be used to get MSI files for a specified folder (with parameters)

.DESCRIPTION

Gets MSI paths from current folder (with no parameters)
Can be used to get MSI files for a specified folder (with parameters)
Does a recursive search and returns full path names for .msi files 
within the folder specified. 

.INPUTS

None. 

.OUTPUTS

System.Object. Hashtable. Returns a numbered key to msi location 
pair.

.Example

PS> Get-MSIs
C:\Users\user1> Get-MSIs

Name                           Value                                                                                                                                                                       
----                           -----                                                                                                                                                                       
2                              C:\Users\User1\Documents\Folder1\Folder2\MSI3.msi
1                              C:\Users\User1\Documents\Folder1\MSI2.msi
0                              C:\Users\User1\Documents\MSI1.msi

.Example

PS> Get-MSIs -folder "C:\Users\User1\Documents\Folder1\Folder2\"
C:\Users\user1> Get-MSIs

Name                           Value                                                                                                                                                                       
----                           -----                                                                                                                                                                       
0                              C:\Users\User1\Documents\Folder1\Folder2\MSI3.msi

.Example

PS> Get-MSIs "C:\Users\User1\Documents\Folder1\Folder2\"
C:\Users\user1> Get-MSIs

Name                           Value                                                                                                                                                                       
----                           -----                                                                                                                                                                       
0                              C:\Users\User1\Documents\Folder1\Folder2\MSI3.msi

#>
    param(
    [string]$Folder="."
    )

    $MSIs = @((ls $Folder -Filter "*.msi" -Recurse).FullName)
    $count = $msis.count
    $MSIArray=@{}
    Foreach ($msi in $MSIs)
    {
        $MSIArray.add(($count -1), $msi)
        $count+=-1
  
    }
        
    #$MSIArray.keys | Select @{l='FullPath';e={$_}}
    
    return $MSIArray#, "Total files: $count"
    
    
}

####               ##
###Installation/Run##

function Run-Wait                  ##Starts a program and waits for exit by using system process diagnostics
                                   #Example of usage: Run-wait setup.exe , or Run-wait "setup.exe" "/quiet"
                                   #the first is run without parameters, the second with the quiet switch
{
    param (
        [string]$file, 
        [string]$params = ""
    )
    Append-Log "Run-Wait  File:    [$file]`n  Params:  [$params]"
    [diagnostics.process]::start($file, $params).waitForExit()
}

function Uninstall-MSI             ##Uninstalls an MSI based on the GUID
                                   #Example of usage: Uninstall-MSI {123-456-789-abc-def}

{
     param (
            [string]$GUID
        )

    Append-Log "UninstallMSI ($GUID)"

    $r=get-wmiobject win32_product | where {$_.IdentifyingNumber -match $GUID}
    if ($r -ne $null) {
        $name = $r.Name
        Append-Log "  Uninstalling MSI package:  $name"
        $RunCommand = "msiexec"
        #$Args =  "/x $GUID REBOOT=ReallySuppress /qb"  
        $Args =  "/x $GUID REBOOT=ReallySuppress MSIRESTARTMANAGERCONTROL=Disable /qn"  
        
        [diagnostics.process]::start($RunCommand, $Args).waitForExit() 
    } else {
        Append-Log "  MSI GUID not found for removal:  $GUID"
    }
}

####               ##
###User Information##

function Get-CurrentUser           ##Finds the current user logged in
                                   #Example of usage: Get-CurrentUser
{
    $username = ((Get-WmiObject -Class win32_computersystem).username -split "\\")
    $username = $username[1]
    if ($username -eq $null)
    {
        Append-Log "Error with Username"
        
    }
    else 
    {
        return $username
    }
}

function Get-UserPath              ##Gets user's folder i.e. C:\users\user1. This calls Get-CurrentUser for 
                                   #retrieving the username
                                   #Example of usage: Get-UserPath
{
    $username = Get-CurrentUser   
    $userpath = "$env:HOMEDRIVE\users\$username"
    return $userpath
}
    
####      ##
###Logging##

function Get-Log                   ##Used to Return log path
                                   #Example Usage: Get-log
                                   ##Creates a log in the "$env:SystemDrive\SCCMFiles\Logs" path
                                   
{
    $LogPath       = "$env:SystemDrive\SCCMFILES\Logs";if (!(Test-Path $LogPath)) { new-item -ItemType directory -Path $LogPath}
    $LogPath       = "$LogPath\$Application" + ".log.txt"
    return $Logpath
}
    

function Append-Log                ##Used to append to log
                                   #Example usages: 
                                   #Append-Log "Words" 
                                   #"Words"|Append-Log 
                                   #(get-aduser username).sid | Append-Log
                                   #($currentuser = get-currentuser)| Append-Log
{
    [cmdletbinding()]
    param 
    ( 
    [parameter(ValueFromPipeline=$True)][string[]]$text

    )
    $logpath=Get-Log
    $timestamp= Get-Date -Format "MM/dd/yyyy HH:mm"
    Write-Output "[$timestamp] $text" # Write to console
    Write-Output "[$timestamp] $text"| Out-File $logpath -Append # Write to log
}



function Get-ErrorLog                   ##Used to set/return Errorlog
{
    $LogPath       = "$env:SystemDrive\SCCMFILES\Logs";if (!(Test-Path $LogPath)) { new-item -ItemType directory -Path $LogPath}
    $LogPath       = "$LogPath\$Application" + ".errorlog.txt"
    return $Logpath
}
    

function Append-ErrorLog           ##Used to append to error log
{
    [cmdletbinding()]
    param 
    ( 
    [parameter(ValueFromPipeline=$True)][string[]]$text

    )
    $logpath=Get-ErrorLog
    $timestamp= Get-Date -Format "MM/dd/yyyy HH:mm"
    Write-Host "!!!Error`t`t[$timestamp] $text" -BackgroundColor "red" # Write to console
    Write-Output "[$timestamp] $text"| Out-File $logpath -Append # Write to log
    
}

function Clear-Logs
{
    $null > $(get-errorlog)
    $null > $(get-log)

}

function Open-Log
{
    param(
        [string]$log="all"
        )
 
    if ($log -eq "all")
    {
        notepad.exe $(get-errorlog)
        notepad.exe $(get-log)
    }
 
    if ($log -like "e*")
    {
        notepad.exe $(get-errorlog)
    }
 
    if ($log -like "m*")
    {
        notepad.exe $(get-log)
    }
 
    
    
}




####    ##
###Debug##

function Get-Errors               #Use to get errors when running
                                   <#
                                   Example: 
                                   
                                   $startingerrors = Get-Errors
                                   
                                   $startCount = $startingerrors.count
                                   $currCount = $startCount
                                   
                                   try{Code}

                                   catch
                                   {
                                   $currCount=(get-errors).count
                                   append-errorlog $error[0]
                                   }

                                   finally{
                                   if ($startCount -ne $currCount)
                                   append-log "Script completed with errors"
                                   }
                                   #>

{
    
    $previousErrors = @()
    foreach ($err in $Error)
    {
        $previousErrors+=$err
    }
    return $previousErrors
}          


                                   #From -- https://stackoverflow.com/questions/2434133/progress-during-large-file-copy-copy-item-write-progress
function Copy-Statusbar            ##For large File copies creates status bar in powershell                       
                                   #Example usage: Copy-Statusbar -from "C:\My Documents\bigfile.zip" -to "C:\temp\"
{ 
    param( [string]$from, [string]$to)
    $ffile = [io.file]::OpenRead($from)
    $tofile = [io.file]::OpenWrite($to)
    Write-Progress -Activity "Copying file" -status "$from -> $to" -PercentComplete 0
    try {
        [byte[]]$buff = new-object byte[] 4096
        [int]$total = [int]$count = 0
        do {
            $count = $ffile.Read($buff, 0, $buff.Length)
            $tofile.Write($buff, 0, $count)
            $total += $count
            if ($total % 1mb -eq 0) {
                Write-Progress -Activity "Copying file" -status "$from -> $to" `
                   -PercentComplete ([int]($total/$ffile.Length* 100))
            }
        } while ($count -gt 0)
    }
    finally {
        $ffile.Dispose()
        $tofile.Dispose()
        Write-Progress -Activity "Copying file" -Status "Ready" -Completed
    }
}


function Get-ProcessList           ##Used to get all processes running, Called by Compare-ProcessList
                                   #Example usage: Get-ProcessList
{
    $processnames= @()
    foreach ($process in ((Get-Process).Name |sort -Unique)) {$processnames += $process}
    return $processnames
}

function Compare-Process{
    New-Item compareprocess.ps1
    Set-Content .\compareprocess.ps1 @'
function Get-ProcessList           ##Used to get all processes running, Called by Compare-ProcessList
                                   #Example usage: Get-ProcessList
{
    $processnames= @()
    foreach ($process in ((Get-Process).Name |sort -Unique)) {$processnames += $process}
    return $processnames
}

function Append-Log                ##Used to append to log
                                   #Example usages: 
                                   #Append-Log "Words" 
                                   #"Words"|Append-Log 
                                   #(get-aduser username).sid | Append-Log
                                   #($currentuser = get-currentuser)| Append-Log
{
    [cmdletbinding()]
    param 
    ( 
    [parameter(ValueFromPipeline=$True)][string[]]$text

    )
    $logpath=Get-Log
    $timestamp= Get-Date -Format "MM/dd/yyyy HH:mm"
    Write-Output "[$timestamp] $text" # Write to console
    Write-Output "[$timestamp] $text"| Out-File $logpath -Append # Write to log
}

function Get-Log                   ##Used to Return log path
                                   #Example Usage: Get-log
                                   ##Creates a log in the "$env:SystemDrive\SCCMFiles\Logs" path
                                   
{
    $LogPath       = "$env:SystemDrive\SCCMFILES\Logs";if (!(Test-Path $LogPath)) { new-item -ItemType directory -Path $LogPath}
    $LogPath       = "$LogPath\$Application" + ".log.txt"
    return $Logpath
}
    $Filterprog=
    $newprocesses=@()
    $process_before = Get-Process
    $count=0
    while ($true)
    {
        $Process_after = Get-Process
        $comparison = Compare-Object $process_before $Process_after
        if (($comparison) -eq $Null)
        {
        }
        else
        {
        
            if ($comparison.sideindicator -eq '=>')
            {
                $status = "Opened"
            }
            else
            {
                $status = "Closed"
            }
        
            $newprocesses+=($comparison).inputobject
            $process_before=$Process_after
            if ($Filterprog -ne $Null)
            {
                $strName = "'%"+${Filterprog}+"%'"
            }
            else
            {
                $strName="'%"+${newprocesses}[$count].Name+"%'"
            }
            $processinfo=(Get-WmiObject Win32_Process -Filter "name like $strname").CommandLine
                      
            if ($processinfo -ne $Null)
            {
            Append-Log "Program $status with Command line :  ", ($newprocesses[$count]).name, $processinfo, "`t`t", "PID: " ,($newprocesses[$count]).id
            }

            else
            {
            Append-Log "Program $status with Command line :  ", ($newprocesses[$count]).name,  "`t`t", "PID: " ,($newprocesses[$count]).id      
            }
            $count+=1
        }
    }

'@

powershell.exe .\compareprocess.ps1
}     #This is used to create a separate powershell instance works with Powershell 2.0



                                   #(This will need to be run in a separate powershell window)
function Compare-ProcessList       ##Creates running process tracker and writes to log and screen          
                                   #This calls on Get-ProcessList and Append-Log
                                   #Example usage: Compare-ProcessList
{
    param(
    [string]$Filterprog
    )


    $newprocesses=@()
    $process_before = Get-Process
    $count=0
    while ($true)
    {
        $Process_after = Get-Process
        $comparison = Compare-Object $process_before $Process_after
        
        if (($comparison) -eq $Null)
        {
        }
        else
        {
        
            if ($comparison.sideindicator -eq '=>')
            {
                $status = "Opened"
            }
            else
            {
                $status = "Closed"
            }
        
            $newprocesses+=($comparison).inputobject
            $process_before=$Process_after
            if ($Filterprog -ne $Null)
            {
                $strName = "'%"+${Filterprog}+"%'"
            }
            else
            {
                $strName="'%"+${newprocesses}[$count].Name+"%'"
            }
            $processinfo=(Get-WmiObject Win32_Process -Filter "name like $strname").CommandLine
            if ($processinfo -ne $Null)
            {
            Append-Log "Program $status with Command line :  ", ($newprocesses[$count]).name, $processinfo, "`t`t", "PID: " ,($newprocesses[$count]).id
            }

            else
            {
            Append-Log "Program $status with Command line :  ", ($newprocesses[$count]).name,  "`t`t", "PID: " ,($newprocesses[$count]).id      
            }
            $count+=1
        }
    }
}

    
    

####                 ##
###System Interaction##
function Block-User                ##Blocks and unblocks user interaction
                                   #Use: Block-User y or Block-User -Block y to block user input
                                   #Block-user or Block-user -Block n to unblock user input
{


    param(
        [string]$Block
        )
    
$code = 
@'
    [DllImport("user32.dll")]
    public static extern bool BlockInput(bool fBlockIt);
'@

    $userInput = Add-Type -MemberDefinition $code -Name Blocker -Namespace UserInput -PassThru

    # block user input
    if ($block -eq "y")
    {
        $null = $userInput::BlockInput($true)
    }
    if ($Block -eq "n")
    {
        $null = $userInput::BlockInput($false)
    }
    else
    {
        $null = $userInput::BlockInput($false)
    }
}
            

function Get-ProgramWindow         ##Gets all active Program Windows,      optional getting PID
                                   #Example Usage: Get-ProgramWindow       this will get all active windows
                                   #Get-ProgramWindow -ProcessIDquery y    this will give you PIDs with the Windows
                                   #This gets called by Filter-ProgramWindow, but filter by PIDs is not implemented
{
    param(
    [string]$ProcessIDquery
    )
    
    $alltitles = (get-process).MainWindowTitle              #Get all window titles
    $processidq = (get-process).Id                          #Get all PIDs
    $count=0
    $titlearray=[System.Collections.ArrayList]::new()
    
    if ($ProcessIDquery -ne "y")
    {
        foreach ($wintitle in $alltitles) 
        {
            if ($wintitle[$count] -match '\p{Ll}') #\p{Ll} to match at least 1 letter, this helps exclude null
            {
                $titlearray.Add($wintitle)   
                $count+=1
            }
                
        }
            return $titlearray
    }
    
    
    if ($ProcessIDquery -eq "y")
    {
        
        foreach ($wintitle in $alltitles) 
        {
        
            if ($wintitle[$count] -match '\p{Ll}')
            {
                
                $x=$processidq[$count]

                $titlearray.Add("$wintitle `tPID: $x")   
                $count+=1
            }
                
        }
            return $titlearray
    }
}



function Filter-ProgramWindow      ##Gets Program Windows based on similar Program name
                                   #Example Usage:  Filter-ProgramWindow -ProgramTitle "Chrome"   this will return all programs with
                                   #the word Chrome in the window title
                                   #This gets called on by SendKeys-ToProgramWindow
{
    param (
    [string]$ProgramTitle,
    [string]$ProcessIDquery
    )

    if ($ProcessIDquery -eq "y")
    {
        $titlearray = Get-ProgramWindow -ProcessIDquery y
    }
    else
    {
        $titlearray = Get-ProgramWindow
    }

    foreach ($item in $titlearray)
    {       
        if ($item -like "*$programtitle*")
        {
                
            return $item
        }
    }
}

function SendKeys-ToProgramWindow  ##Sends keystrokes to window title designated by input
                                   #Example Usage: SendKeys-ToProgramWindow -ProgramTitle "Untitled.docx" -KeysToSend "It was the best of times"
                                   #This is best used in conjunction with Block-User and very specific Window Titles
{
    param (
        [string]$ProgramTitle,
        [string]$KeysToSend
        )
        $programTitle = Filter-ProgramWindow $ProgramTitle

                write-host "Sending Keys `"$keystosend`" to program with window titled: $item"
                $wshell = New-Object -ComObject wscript.shell;
                $wshell.AppActivate("$programtitle")
                $wshell.SendKeys($KeysToSend)
}

function Set-RegistryACL
{
    param(
    [string]$regpath,
    [string]$group="Authenticated Users",
    [string]$accesstype="FullControl",
    [string]$allow="Allow",
    [string]$setORremove="SetAccessRule" #RemoveAccessRule
    )

    if ((test-path $regpath) -ne $true)
    {
        New-Item $regpath -Force
    }
    $acl =Get-Acl $regpath
    $rule = New-Object System.Security.AccessControl.RegistryAccessRule ("$group","$accesstype","$allow")

    if ($setORremove -eq "SetAccessRule")
    {
        $acl.SetAccessRule($rule)
    }
    if ($setORremove -eq "RemoveAccessRule")
    {
        $acl.RemoveAccessRule($rule)
    }
$acl |Set-Acl -Path $regpath

}


function Grant-Permissions 
{
    param (
        [string]$folder,
        [string]$force="yes"
    )

    Append-Log "GrantPermissions to ($folder)"
    
    if (-not (test-path $folder)) {
         Append-Log "Error, path `$path not found:  $folder"
         Append-ErrorLog "Error, path `$path not found:  $folder"
         return
    }

    $exe = "icacls"
    if ($folder.endswith('\')) { 
        $folder = $folder.Substring(0, ($folder.length-1))  # Remove last character if it is a backslash 
    } 
    $args = " `"$folder`" /grant `"Authenticated Users`":(OI)(CI)M" #7/20/17 changed from 'F' to 'M'.  F allows editing other user's permissions.
    Run-Wait $exe $args
}


################Testing

function Get-ExecutableType   #https://gallery.technet.microsoft.com/scriptcenter/Identify-16-bit-32-bit-and-522eae75
{
    <#
    .Synopsis
       Determines whether an executable file is 16-bit, 32-bit or 64-bit.
    .DESCRIPTION
       Attempts to read the MS-DOS and PE headers from an executable file to determine its type.
       The command returns one of four strings (assuming no errors are encountered while reading the
       file):
       "Unknown", "16-bit", "32-bit", or "64-bit"
    .PARAMETER Path
       Path to the file which is to be checked.
    .EXAMPLE
       Get-ExecutableType -Path C:\Windows\System32\more.com
    .INPUTS
       None.  This command does not accept pipeline input.
    .OUTPUTS
       String
    .LINK
        http://msdn.microsoft.com/en-us/magazine/cc301805.aspx
    #>

    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
        [string]
        $Path
    )

    try
    {
        try
        {
            $stream = New-Object System.IO.FileStream(
                $PSCmdlet.GetUnresolvedProviderPathFromPSPath($Path),
                [System.IO.FileMode]::Open,
                [System.IO.FileAccess]::Read,
                [System.IO.FileShare]::Read
            )
        }
        catch
        {
            throw "Error opening file $Path for Read: $($_.Exception.Message)"
        }

        $exeType = 'Unknown'
        
        if ([System.IO.Path]::GetExtension($Path) -eq '.COM')
        {
            # 16-bit .COM files may not have an MS-DOS header.  We'll assume that any .COM file with no header
            # is a 16-bit executable, even though it may technically be a non-executable file that has been
            # given a .COM extension for some reason.

            $exeType = '16-bit'
        }

        $bytes = New-Object byte[](4)

        if ($stream.Length -ge 64 -and
            $stream.Read($bytes, 0, 2) -eq 2 -and
            $bytes[0] -eq 0x4D -and $bytes[1] -eq 0x5A)
        {
            $exeType = '16-bit'

            if ($stream.Seek(0x3C, [System.IO.SeekOrigin]::Begin) -eq 0x3C -and
                $stream.Read($bytes, 0, 4) -eq 4)
            {
                if (-not [System.BitConverter]::IsLittleEndian) { [Array]::Reverse($bytes, 0, 4) }
                $peHeaderOffset = [System.BitConverter]::ToUInt32($bytes, 0)

                if ($stream.Length -ge $peHeaderOffset + 6 -and
                    $stream.Seek($peHeaderOffset, [System.IO.SeekOrigin]::Begin) -eq $peHeaderOffset -and
                    $stream.Read($bytes, 0, 4) -eq 4 -and
                    $bytes[0] -eq 0x50 -and $bytes[1] -eq 0x45 -and $bytes[2] -eq 0 -and $bytes[3] -eq 0)
                {
                    $exeType = 'Unknown'

                    if ($stream.Read($bytes, 0, 2) -eq 2)
                    {
                        if (-not [System.BitConverter]::IsLittleEndian) { [Array]::Reverse($bytes, 0, 2) }
                        $machineType = [System.BitConverter]::ToUInt16($bytes, 0)

                        switch ($machineType)
                        {
                            0x014C { $exeType = '32-bit' }
                            0x0200 { $exeType = '64-bit' }
                            0x8664 { $exeType = '64-bit' }
                        }
                    }
                }
            }
        }
        
        return $exeType
    }
    catch
    {
        throw
    }
    finally
    {
        if ($null -ne $stream) { $stream.Dispose() }
    }
    
}


Function Get-MSIProperty {
param([string]$pathToMSI = ".")
    $msiOpenDatabaseModeReadOnly = 0
    $msiOpenDatabaseModeTransact = 1

    $windowsInstaller = New-Object -ComObject windowsInstaller.Installer

    

    $database = $windowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $windowsInstaller, @($pathToMSI, $msiOpenDatabaseModeReadOnly))

    $query = "SELECT Property, Value FROM Property"
    $propView = $database.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $database, ($query))
    $propView.GetType().InvokeMember("Execute", "InvokeMethod", $null, $propView, $null) | Out-Null
    $propRecord = $propView.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $propView, $null)
		
    while  ($propRecord -ne $null)
    {
	    $col1 = $propRecord.GetType().InvokeMember("StringData", "GetProperty", $null, $propRecord, 1)
	    $col2 = $propRecord.GetType().InvokeMember("StringData", "GetProperty", $null, $propRecord, 2)
 
	    write-host $col1 - $col2
	
	    #fetch the next record
	    $propRecord = $propView.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $propView, $null)	
    }

    $propView.GetType().InvokeMember("Close", "InvokeMethod", $null, $propView, $null) | Out-Null          
    $propView = $null 
    $propRecord = $null
    $database = $null
}   #https://www.alkanesolutions.co.uk/blog/2016/12/13/query-windows-installer-msi-using-powershell/


function Get-PSScriptExitCode
{
    param(
        [Switch]$s,
        [String]$program,
        [String]$scriptname
        )
   
    $code = (Start-Process "$program" "$scriptname" -Wait -Passthru).ExitCode
    return $code
}

function Switch-Null
{
    param(
        [switch]$Nullswitch,
        [string]$username
        )
        
        if ($Nullswitch.IsPresent -eq $true)
        {
            
            $usernames=((Get-ADUser $username -Properties *).name, (Get-ADUser $username -Properties *).sid.value, (Get-ADUser $username -Properties *).mail )
            [System.Collections.Arraylist]$uproperties=$usernames
        }
            
            

            return $uproperties, "$username was here."
        


}

    
Function Get-Something 
{ #https://learn-powershell.net/2013/05/07/tips-on-implementing-pipeline-support/
    [cmdletbinding()]
    Param (
        [parameter(ValueFromPipelineByPropertyName=$True)]
        [string[]]$Name,
        [parameter(ValueFromPipelineByPropertyName=$True)]
        [string[]]$Directory
    )
    Begin {
        Write-Verbose "Initialize stuff in Begin block"
    }

    Process {
        Write-Verbose "Process block"
        Write-Host "Name: $Name"
        Write-Host "Directory: $Directory"
    }

    End {
        Write-Verbose "Final work in End block"
        $Report
    }
}

function Pipe-Text
{
    [cmdletbinding()]
    param(
        [parameter(ValueFromPipeline=$True)][string[]]$test
        )
        return $test
}

function Create-Shortcut() 
<#Usage $folder = " Applications"
                           
$folder = "FolderName"
$ICO = "Icon.exe"
$params = @{
    'File'              = "Shortcut.lnk"
    'Subfolder'         = $folder
    'ICO'               = "$Installpath\$ICO"
    'TargetPath'        = "$ProgramFiles\PATH2.exe" 
    'WorkingDirectory'  = ""
    'Description'       = ""
    'Arguments'         =  "https://webpage.com"
}
Create-Shortcut @params

#>
{
    param (
        [string] $file,             # LNK or URL.  Ex. "Unity Real Time.lnk"     
        [string] $subfolder,
        [string] $ICO,
        [string] $TargetPath,
        [string] $WorkingDirectory,
        [string] $Description,
        [string] $Arguments  
    )
    

    Append-Log "# CreateShortcut()"

    # Determine if URL.  If URL, file will be written as text file.    
    $URL = 0
    if ($file.ToUpper().EndsWith(".URL")) { 
        $URL = 1 # File is a URL
        Append-Log "URL $URL"
    }
              
    # Verify $InstallPath exists.  $ICO is stored here.
    if (-not (test-path $InstallPath)) {
        Append-Log "Error, file not found:  $InstallPath"
        exit
    }
    
    
    # Verify $ICO exists 
    if (-not (test-path $Ico)) {
        Append-Log "Error, ICO file not found:  $ICO"
        exit
    }


    # Define both start menu path(s)    
    
    $pathStartMenu1 = "C:\Environment\StartMenu\Programs\$subfolder"
    $pathStartMenu2 = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\$subfolder"

    <# NOTE:
        For Windows 7, the shortcut needs to be created in the C:\Environmments folder first due to the replication service copying to C:\ProgramData.
        This process is one-way - files found in C:\ProgramData that are not in C:\Environments are deleted by the service.
        This function will also create the shortcut at C:\ProgramData so it appears immediately.  This avoids the perception that the installation failed
        and issues with detection methods.
        
        For Windows 10, the shortcut is just created in the C:\ProgramData path.
    #>



    #--- If OS is Windows 10, set $pathStartMenu1 = $pathStartMenu2
    
    if ($OS -like "*Windows 10*") {
        $pathStartMenu1 = $pathStartMenu2
        Append-Log "OS is Windows 10.  Start Menu path set to:  $pathStartMenuPath1"
    }


    # Verify $pathStartMenu1 exists
    if (-not (Test-Path $pathStartMenu1)) {
        new-item $pathStartMenu1 -type directory
        Append-Log "  Created path:  $pathStartMenu1"
    }

    # If not Windows 10, verify $pathStartMenu2 exists
    if (!($OS -like "*Windows 10*")) {
        if (-not (Test-Path $pathStartMenu2)) {
            new-item $pathStartMenu2 -type directory
            Append-Log "  Created path:  [$pathStartMenu2]"
        }
    }

    
    # Create shortcut initially at $pathStartMenu1.  URL files are written as text files; all others are created as wscript shell objects.
    if ($URL) {
        <# Initially write file at $pathStartMenu1
        set-content "$pathStartMenu1\$file" "[InternetShortcut]"
        add-content "$pathStartMenu1\$file" "URL=$targetpath"
        add-content "$pathStartMenu1\$file" "IconFile=$ICO"
        add-content "$pathStartMenu1\$file" "IconIndex=0"
        add-content "$pathStartMenu1\$file" "HotKey=0"
        #>
        
         # Initially create object at $pathStartMenu1
        $oShell = New-Object -ComObject wscript.shell
        $oLnk = $oShell.createshortcut("$pathStartMenu1\$file") 
        $oLnk.TargetPath = $TargetPath
        $oLnk.save()
        
        add-content "$pathStartMenu1\$file" "IconFile=$ICO"
        add-content "$pathStartMenu1\$file" "IconIndex=0"
        add-content "$pathStartMenu1\$file" "HotKey=0"
        
        
        
    } else {
      # Initially create object at $pathStartMenu1
        $oShell = New-Object -ComObject wscript.shell
        $oLnk = $oShell.createshortcut("$pathStartMenu1\$file") 
        $oLnk.TargetPath = $TargetPath
        $oLnk.IconLocation = "$ICO, 0"
        $oLnk.Description = $Description
        $oLnk.WorkingDirectory = $WorkingDirectory # Note working directory is not specified for URL targets.
        $oLnk.Arguments = $Arguments # Note this cannot be specific for URLs.    
        $oLnk.save()
    }

    $chkFile = "$pathStartMenu1\$file"
    $count = 0 
    while ($count -le 30) {
        if (test-path $chkFile) {
            Append-Log "  Verified file:  $chkFile"
            break
        }
        Append-Log "Waiting for file to exist:  $chkFile"
        start-sleep 1
    }
    if (!(test-path $chkFile)) { Append-Log "WARNING, failed to create file:  $chkFile"}
    
    
    # If not Windows 10, copy file from first start menu to second
    if (!($OS -like "*Windows 10*")) {
        Copy-Item "$pathStartMenu1\$file" $pathStartMenu2
    }
}

function Change-Admin 
{

    param(
        [string]$newadmin
        )

    $computer = [ADSI]("WinNT://$env:COMPUTERNAME,computer")
    $userlist = $computer.psbase.children | where-object {$_.psbase.schemaclassname -eq 'user'}



    foreach ($user in $userList)
    {
    # Create a user object in order to get its SID
        $userObject = New-Object System.Security.Principal.NTAccount($user.Name)
        $userSID = $userObject.Translate([System.Security.Principal.SecurityIdentifier])
   # Look for local “Administrator” SID
        if(($userSID.Value.substring(0,6) -eq "S-1-5-") -and ($userSID.Value.substring($userSID.Value.Length – 4, 4) -eq "-500"))
        {
      # Rename local Administrator account 
            $Error.Clear()
            try
            {    
                $localAdmin=[adsi]"WinNT://./$userObject,user"
                $localAdmin.psbase.rename($newadmin)
            }
            catch [System.UnauthorizedAccessException]         
            {
            Write-Host "Access Denied" -ForegroundColor Red
            }
        }
    }
}


function Kill-process                            #Usage kill-process -kill "name"                kill exact process name match
                                                 #Usage kill-process -kill "name" -all            kill all processes like name
{
    param(
    [string]$kill,
    [switch]$all
    )

    $processList = Get-ProcessList
    foreach ($p2kill in $processList) 
    {
        if ($kill -eq $p2kill) #Only exact matching processes
        {
            Stop-Process $p2kill -Force
        }
        
        if ($all.IsPresent -eq $true -and $kill -like $p2kill) #Searches all instances
        {
            Stop-Process $p2kill -force
        }

    }
}

function ReMove-StartMenuShortcuts   #Use    no param           to remove shortcuts, 
                                     #Use   -moveto pathname    to move
{

    param (
    [string]$shortcutname,
    [string]$moveto
    )

    $StartMenuPath = "C:\Environment\StartMenu\Programs"
    If ($OS -like "*Windows 10*") 
    {
        $StartMenuPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs"
    }
    $filearray = @()

    $shortcuts= (ls $StartMenuPath | Select-String $shortcutname).Filename|select-string .lnk
    Append-log "List of shortcuts $shortcuts"

    foreach ($file in $x)
    {
        $filearray+=$file
    }

    if ($moveto.IsPresent -eq $true)
    {
        foreach ($shortcut in $filearray)
        {
            Append-Log "Moving $shortcut"
            Move-Item "$startmenupath\$shortcut" "$startmenupath\$moveto"
        }
    }
    else
    {
        foreach ($shortcut in $filearray)
        {
            Append-Log "Removing $shortcut"
            Remove-Item "$startmenupath\$shortcut"
        }
    }
}


<#
Loop to retry ACTIONS
do{
    $Failed= $false
        Try
        {
            VARIABLES
            CODE -erroraction Stop
        }
        Catch {
        $Failed = $true
        }
{
while ($Failed)    

-------------------------------------------------------------------
#>
<#

$processlist = Get-processlist
#$exclusionlist=@("smartscreen","backgroundTaskHost","runtimebroker","conhost","svchost","teams","UpdateTrustedSites","SearchProtocolHost")

while ($true)
{
    
    $comparelist = get-processlist
    $comparison = Compare-Object -referenceobject $processlist -differenceobject $comparelist
    
    


    if ($exclusionlist.Contains($comparison.inputobject) -ne $True)
    {
        if ($comparison.SideIndicator -eq '=>')
        {
            $processinfo=(Get-WmiObject Win32_Process -Filter "name like '%$($comparison.inputobject)%'").CommandLine
            $programcount = $processinfo.Count
            write-host $comparison.InputObject "was opened: $processinfo"
            $processlist = Get-ProcessList
        }
        if ($comparison.SideIndicator -eq '<=')
        {
            $processinfo=(Get-WmiObject Win32_Process -Filter "name like '%$($comparison.inputobject)%'").CommandLine
            write-host $comparison.InputObject "was closed. `n Stillopen:$processinfo"   
            $processlist = Get-ProcessList
        }
    }
    else
    {
        $processlist = Get-ProcessList
    }    
    
}


#>


################Testing


<#Example script
#$ErrorActionPreference = "continue"
$Error.Clear()
$errcount = 0
$startingerrors = Get-Errors
$startCount = $startingerrors.count
$currenterrors = get-errors
$currCount = (get-errors).count


try
{

#Variables can be piped into Append-Log if you wish to record the variables as well

    "Current path is " + ($scriptpath =((pwd).path)) |Append-Log
    $Application = "Test"
    Append-Log $Application    
    ($OS = Get-OS)|Append-Log  
    Write-Output "Architecture is $($Arch = Get-arch)"|append-log
    #Set-ProgramFiles32bit for 32bit installers, Set-ProgramFiles64bit for 64 bit only installer
    "Program Files path is " + ($ProgramFiles = Set-ProgramFiles32bit) |Append-Log                 #This way always works     "string" + ($variable=function) | Append-Log
    Write-Output "Current user is $(currentuser = Get-CurrentUser)"|append-log
    "Log path is " + (Get-Log)|append-log
    "Program installation path is " + ($installpath = "$ProgramFiles\ApplicationName") |Append-Log
    "Current path is " + ($scriptpath =((pwd).path)) |Append-Log
#Error
    try {DESTROYALLHUMANS}
    catch
    {
    Append-Log "DESTROYALLHUMANS failure"
    Append-ErrorLog $error[0]
    $errcount+=1
    }
    blaka


}

catch
{

        
    if ($currCount -ne $startCount)
    {
        $errcount+=1
        if ($currenterrors -ne $null)
        {
            $comparision = Compare-Object $currenterrors $startingerrors
        }
        
        $errlog = Get-ErrorLog
        if ((test-path $errlog) -ne $true)
        {
            write-host -BackgroundColor Yellow "`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t"
            Write-output "Error file created in $errlog" | Append-Log 
            write-host -BackgroundColor Yellow "`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t"
            
        }
        Append-Log "Script Failure"
        Append-ErrorLog $Error[0]
        $startingerrors =Get-errors
        $errCount+=1
    }
    else{$errCount+=1}
        
}

finally
{
    if ($currCount -eq $startCount)
    {
        Append-Log "Script Completed with no errors"
    }
    else
    {
        Append-Log "Script Completed with $errcount errors"
        
    }

}

#>
