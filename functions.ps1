<#                                  Functions                             #>

############################################################################

 

####       ##

####Info Gathering##

 

function Get-OS                     ##Gets operating System based on WMI caption

{

    $OperatingSystem = (gwmi win32_operatingsystem).caption

    return $OperatingSystem

}

 

function Get-Arch                   ##Gets architecture based on stackpointer size

{

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

 

function Set-ProgramFiles32bit      ##Sets file path for 32 bit programs to install to (defaults $env:Drive\Program Files for 32 bit

                                    #and 64 $env:SystemRoot\Program Files (x86) for other architectures

{

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

 

function Set-ProgramFiles64bit      ##Sets file path for 64 bit programs to install to

{

    $Arch = Get-Arch

    $ProgramFiles = "$env:Drive\Program Files" 

    return $ProgramFiles

}

 

 

function Filter-Program             ##Used to filter GUIDs based on program name

                                    #Example of usage: Filter-Program -FilteredProg "Outlook" would find all installed programs, GUIDs, and MSI locations

                                    #for programs with the word "Outlook" in the add/remove program list

{

    param (

        [string]$FilteredProg

        )

    $progra = Get-WmiObject win32_product -Filter "Name LIKE '%$FilteredProg%'"|Get-ProgramGUIDs -Verbose

    Write-Output $progra   

}

function Get-ProgramGUIDs           ##Used by calling Filter-Program, can be used alone to filter from get-wmiobject pipeline

                                    #Example of lone usage: get-wmiobject win32_product |get-programguids would return

                                    #all programs, GUIDs, and MSI locations.

{

   [cmdletbinding()]

    Param (

        [parameter(ValueFromPipelineByPropertyName=$True)]

        [string[]]$Name,

        [parameter(ValueFromPipelineByPropertyName=$True)]

        [string[]]$IdentifyingNumber,

        [parameter(ValueFromPipelineByPropertyName=$True)]

        [string[]]$LocalPackage

       

    )

    Begin {

        Write-Output "Processing Program Data:%60n%22

    }

 

    Process {

       

        Write-Output "

        Name: $Name

        GUID: $IdentifyingNumber

        MSI Location: $LocalPackage

        "

    }

 

    End {

        $Report

    }

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

    Append-Log "Run-Wait  File: [$file]`n  Params:  [$params]"

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

        #$Args =  "/x $GUID REBOOT=ReallySuppress /qb"  # Prompts for reboot by default.Uninstall CellaVision 5.01")

        $Args =  "/x $GUID REBOOT=ReallySuppress MSIRESTARTMANAGERCONTROL=Disable /qn"  # Prompts for reboot by default.Uninstall CellaVision 5.01")

       

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

        break

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

 

function Create-Log                ##Creates a log in the "$env:SystemDrive\SCCMFiles\Logs" path

                                   #If the path does not exist it creates this path

                                   #Example Usage: Create-Log "Outlook"

{

    param (

    [string]$Application

    )

    $LogPath       = "$env:SystemDrive\SCCMFILES\Logs";if (!(Test-Path $LogPath)) { new-item -ItemType directory -Path $LogPath}

    $LogFile       = "$LogPath\$Application" + ".log.txt"

}

 

function Get-Log                   ##Used to Return log path

                                   #Example Usage: Get-log

{

    $LogPath       = "$env:SystemDrive\SCCMFILES\Logs"

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

 

####    ##

###Debug##

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

 

                                   #(This will need to be run in a separate powershell window)

function Compare-ProcessList       ##Creates running process tracker and writes to log and screen         

                                   #This calls on Get-ProcessList and Append-Log

                                   #Example usage: Compare-ProcessList

{

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

            $strName="'%"+${newprocesses}[$count].Name+"%'"

            $processinfo=(Get-WmiObject Win32_Process -Filter "name like $strname").CommandLine

            if ($processinfor -ne $Null)

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

   

    $alltitles = (get-process).mainwindowtitle

    $processidq = (get-process).Id

    $count=0

    $titlearray=[System.Collections.ArrayList]::new()

   

    if ($ProcessIDquery -ne "y")

    {

        foreach ($wintitle in $alltitles)

        {

            if ($wintitle[$count] -match '\p{Ll}')

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

 

                $titlearray.Add("$wintitle $x")  

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

    [string]$ProgramTitle

    #[string]$PIDnum

    )

 

 

    $titlearray = Get-ProgramWindow

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

 

 

################Testing

function Get-PSScriptExitCode

{

    param(

        [Switch]$s,

        [String]$program,

        [String]$scriptname

        )

  

   <# $code = (Start-Process "$program" "$scriptname" -Wait -Passthru).ExitCode

    return $code#>

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

 

   

Function Get-Something { #https://learn-powershell.net/2013/05/07/tips-on-implementing-pipeline-support/

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

 

################Testing

 

 

#Example Usages

 

$ScriptPath    = split-path $MyInvocation.MyCommand.Definition

$Application = "Test"

Create-Log $Application

($OS = Get-OS)|append-log

($Arch= Get-arch)|append-log

#Set-ProgramFiles32bit for 32bit installer, Set-ProgramFiles64bit for 64 bit only

($ProgramFiles  = Set-ProgramFiles32bit)|append-log

($currentuser = Get-CurrentUser)|append-log

($Logpath = Get-Log)|append-log
