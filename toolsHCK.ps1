<#
.SYNOPSIS
  toolsHCK: Powershell wrapper for HCK\HLK Studio API.
.DESCRIPTION
  A script tool set for HCK\HLK automation powered by HCK\HLK API provided with the Microsoft's Windows HCK\HLK Studio.
.NOTES
  Author:         Bishara AbuHattoum <Bishara@daynix.com>
  License:        BSD
#>

# server switch to inform the script is to run as a server
[CmdletBinding()]
param([Switch]$server)

if ($env:WTTSTDIO -like "*\Hardware Certification Kit\*") {
    $Studio = "hck"
    if ($env:PROCESSOR_ARCHITECTURE -ne "x86") {

        if (-Not $json) {
            Write-Warning "HCK script should be run under a 32bit PowerShell"
            Write-Host "Redirecting ..."
        }

        $PowerShell = [System.IO.Path]::Combine($PSHOME.tolower().replace("system32","sysWOW64"), "powershell.exe")

        if ($MyInvocation.Line) {

            &"$PowerShell" -NoProfile $MyInvocation.Line
        }else{

            &"$PowerShell" -NoProfile -File "$($MyInvocation.InvocationName)"
        }

    exit $LASTEXITCODE
    }
} else {
    $Studio = "hlk"
}

##
$MaxJsonDepth = 6
##

#
# Loadinf HCK\HLK libraries
[System.Reflection.Assembly]::LoadFrom($env:WTTSTDIO + "\microsoft.windows.kits.hardware.filterengine.dll") | Out-Null
[System.Reflection.Assembly]::LoadFrom($env:WTTSTDIO + "\microsoft.windows.kits.hardware.objectmodel.dll") | Out-Null
[System.Reflection.Assembly]::LoadFrom($env:WTTSTDIO + "\microsoft.windows.kits.hardware.objectmodel.dbconnection.dll") | Out-Null
[System.Reflection.Assembly]::LoadFrom($env:WTTSTDIO + "\microsoft.windows.kits.hardware.objectmodel.submission.dll") | Out-Null
[System.Reflection.Assembly]::LoadFrom($env:WTTSTDIO + "\microsoft.windows.kits.hardware.objectmodel.submission.package.dll") | Out-Null

#
# Task
function New-Task($name, $stage, $status, $taskerrormessage, $tasktype, $childtasks) {
    $task = New-Object PSObject
    $task | Add-Member -type NoteProperty -Name name -Value $name
    $task | Add-Member -type NoteProperty -Name stage -Value $stage
    $task | Add-Member -type NoteProperty -Name status -Value $status
    $task | Add-Member -type NoteProperty -Name taskerrormessage -Value $taskerrormessage
    $task | Add-Member -type NoteProperty -Name tasktype -Value $tasktype
    $task | Add-Member -type NoteProperty -Name childtasks -Value $childtasks
    return $task
}

#
# PackageProgressInfo
function New-PackageProgressInfo($current, $maximum, $message) {
    $packageprogressinfo = New-Object PSObject
    $packageprogressinfo | Add-Member -type NoteProperty -Name current -Value $current
    $packageprogressinfo | Add-Member -type NoteProperty -Name maximum -Value $maximum
    $packageprogressinfo | Add-Member -type NoteProperty -Name message -Value $message
    return $packageprogressinfo
}

#
# ProjectPackage
function New-ProjectPackage($name, $projectpackagepath) {
    $projectpackage = New-Object PSObject
    $projectpackage | Add-Member -type NoteProperty -Name name -Value $name
    $projectpackage | Add-Member -type NoteProperty -Name projectpackagepath -Value $projectpackagepath
    return $projectpackage
}

#
# TestResultLogsZip
function New-TestResultLogsZip($testname, $testid,$status, $logszippath) {
    $testresultlogszip = New-Object PSObject
    $testresultlogszip | Add-Member -type NoteProperty -Name testname -Value $testname
    $testresultlogszip | Add-Member -type NoteProperty -Name testid -Value $testid
    $testresultlogszip | Add-Member -type NoteProperty -Name status -Value $status
    $testresultlogszip | Add-Member -type NoteProperty -Name logszippath -Value $logszippath
    return $testresultlogszip
}

#
# TestResult
function New-TestResult($name, $completiontime, $scheduletime, $starttime, $status, $arefiltersapplied, $target, $tasks) {
    $testresult = New-Object PSObject
    $testresult | Add-Member -type NoteProperty -Name name -Value $name
    $testresult | Add-Member -type NoteProperty -Name completiontime -Value $completiontime
    $testresult | Add-Member -type NoteProperty -Name scheduletime -Value $scheduletime
    $testresult | Add-Member -type NoteProperty -Name starttime -Value $starttime
    $testresult | Add-Member -type NoteProperty -Name status -Value $status
    $testresult | Add-Member -type NoteProperty -Name arefiltersapplied -Value $arefiltersapplied
    $testresult | Add-Member -type NoteProperty -Name target -Value $target
    $testresult | Add-Member -type NoteProperty -Name tasks -Value $tasks
    return $testresult
}

#
# FilterResult
function New-FilterResult($appliedfilterson) {
    $filterresult = New-Object PSObject
    $filterresult | Add-Member -type NoteProperty -Name appliedfilterson -Value $appliedfilterson
    return $filterresult
}

#
# Test
function New-Test($name, $id, $testtype, $estimatedruntime, $requiresspecialconfiguration, $requiressupplementalcontent, $scheduleoptions, $status, $executionstate) {
    $test = New-Object PSObject
    $test | Add-Member -type NoteProperty -Name name -Value $name
    $test | Add-Member -type NoteProperty -Name id -Value $id
    $test | Add-Member -type NoteProperty -Name testtype -Value $testtype
    $test | Add-Member -type NoteProperty -Name estimatedruntime -Value $estimatedruntime
    $test | Add-Member -type NoteProperty -Name requiresspecialconfiguration -Value $requiresspecialconfiguration
    $test | Add-Member -type NoteProperty -Name requiressupplementalcontent -Value $requiressupplementalcontent
    $test | Add-Member -type NoteProperty -Name scheduleoptions -Value $scheduleoptions
    $test | Add-Member -type NoteProperty -Name status -Value $status
    $test | Add-Member -type NoteProperty -Name executionstate -Value $executionstate
    return $test
}

#
# ProductInstanceTarget
function New-ProductInstanceTarget($name, $key, $machine) {
    $productinstancetarget = New-Object PSObject
    $productinstancetarget | Add-Member -type NoteProperty -Name name -Value $name
    $productinstancetarget | Add-Member -type NoteProperty -Name key -Value $key
    $productinstancetarget | Add-Member -type NoteProperty -Name machine -Value $machine
    return $productinstancetarget
}

#
# ProductInstance
function New-ProductInstance($name, $osplatform, $targetedpool, $targets) {
    $productinstance = New-Object PSObject
    $productinstance | Add-Member -type NoteProperty -Name name -Value $name
    $productinstance | Add-Member -type NoteProperty -Name osplatform -Value $osplatform
    $productinstance | Add-Member -type NoteProperty -Name targetedpool -Value $targetedpool
    $productinstance | Add-Member -type NoteProperty -Name targets -Value $targets
    return $productinstance
}

#
# Project
function New-Project($name, $creationtime, $modifiedtime, $status, $productinstances) {
    $project = New-Object PSObject
    $project | Add-Member -type NoteProperty -Name name -Value $name
    $project | Add-Member -type NoteProperty -Name creationtime -Value $creationtime
    $project | Add-Member -type NoteProperty -Name modifiedtime -Value $modifiedtime
    $project | Add-Member -type NoteProperty -Name status -Value $status
    $project | Add-Member -type NoteProperty -Name productinstances -Value $productinstances
    return $project
}

#
# Target
function New-Target($name, $key) {
    $target = New-Object PSObject
    $target | Add-Member -type NoteProperty -Name name -Value $name
    $target | Add-Member -type NoteProperty -Name key -Value $key
    return $target
}

#
# Machine
function New-Machine($name, $state, $lastheartbeat) {
    $machine = New-Object PSObject
    $machine | Add-Member -type NoteProperty -Name name -Value $name
    $machine | Add-Member -type NoteProperty -Name state -Value $state
    $machine | Add-Member -type NoteProperty -Name lastheartbeat -Value $lastheartbeat
    return $machine
}

#
# Pool
function New-Pool($name, $machines) {
    $pool = New-Object PSObject
    $pool | Add-Member -type NoteProperty -Name name -Value $name
    $pool | Add-Member -type NoteProperty -Name machines -Value $machines
    return $pool
}

#
# ActionResult
function New-ActionResult($content, $exception = $nil) {
    $actionresult = New-Object PSObject
    if ([String]::IsNullOrEmpty($exception)) {
        $actionresult | Add-Member -type NoteProperty -Name result -Value "Success"
        if (-Not [String]::IsNullOrEmpty($content)) {
            $jsoncontent = (ConvertFrom-Json $content)
            if ($jsoncontent -is [System.Object[]]) {
                $actionresult | Add-Member -type NoteProperty -Name content -Value $jsoncontent.SyncRoot
            } else {
                $actionresult | Add-Member -type NoteProperty -Name content -Value $jsoncontent
            }
        }
    } else {
        $actionresult | Add-Member -type NoteProperty -Name result -Value "Failure"
        if ([String]::IsNullOrEmpty($exception.InnerException)) {
            $actionresult | Add-Member -type NoteProperty -Name message -Value $exception.Message
        } else {
            $actionresult | Add-Member -type NoteProperty -Name message -Value $exception.InnerException.Message
        }
    }
    return $actionresult
}

# ------------------------------------------------------------ #
# Functions, one for each action the script is able to perform #
# ------------------------------------------------------------ #
# ListPools
function listpools {
    [CmdletBinding()]
    param([Switch]$help)

    function Usage {
        Write-Output "listpools:"
        Write-Output ""
        Write-Output "A script that lists the pools info."
        Write-Output "and last heart beat."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "listpools [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output " help = Shows this message."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if (-Not $json) {
        foreach ($Pool in $RootPool.GetChildPools()) {
            Write-Output "============================================="
            Write-Output "Pool name : $($Pool.Name)"

            Write-Output ""

            $Machines = $Pool.GetMachines()

            if ($Machines.Count -lt 1) {
                Write-Output "    The pool is empty!"
            } else {
                Write-Output "    Machines :"

                foreach ($Machine in $Machines) {
                    Write-Output "        Name            : $($Machine.Name)"
                    Write-Output "        State           : $($Machine.Status)"
                    Write-Output "        Last heart beat : $($Machine.LastHeartBeat)"
                    Write-Output ""
                }
            }

            Write-Output "============================================="
        }
    } else {
        $poolslist = New-Object System.Collections.ArrayList
        foreach ($Pool in $RootPool.GetChildPools()) {
            $machineslist = New-Object System.Collections.ArrayList
            $Machines = $Pool.GetMachines()
            foreach ($Machine in $Machines) {
                $machineslist.Add((New-Machine $Machine.Name $Machine.Status.ToString() $Machine.LastHeartBeat.ToString())) | Out-Null
            }
            $poolslist.Add((New-Pool $Pool.Name $machineslist)) | Out-Null
        }
        ConvertTo-Json @($poolslist) -Depth 3 -Compress
    }
}
#
# CreatePool
function createpool {
    [CmdletBinding()]
    param([Switch]$help, [Parameter(Position=1)][String]$pool)

    function Usage {
        Write-Output "createpool:"
        Write-Output ""
        Write-Output "A script that creates a pool."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "createpool <poolname> [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "     help = Shows this message."
        Write-Output ""
        Write-Output " poolname = The name of the pool."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if ([String]::IsNullOrEmpty($pool)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a pool's name."
            Usage; return
        } else {
            throw "Please provide a pool's name."
        }
    }

    if (-Not $json) { Write-Output "Creating pool $pool in Root pool." }
    $RootPool.CreateChildPool($pool) | Out-Null
}
#
# DeletePool
function deletepool {
    [CmdletBinding()]
    param([Switch]$help, [Parameter(Position=1)][String]$pool)

    function Usage {
        Write-Output "deletepool:"
        Write-Output ""
        Write-Output "A script that deletes a pool."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "deletepool <poolname> [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "     help = Shows this message."
        Write-Output ""
        Write-Output " poolname = The name of the pool."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if ([String]::IsNullOrEmpty($pool)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a pool's name."
            Usage; return
        } else {
            throw "Please provide a pool's name."
        }
    }

    if (-Not ($WntdPool = $RootPool.GetChildPools() | Where-Object { $_.Name -eq $pool })) { throw "Provided pool's name is not valid, aborting..."}

    if (-Not $json) { Write-Output "Deleting pool $pool in Root pool." }
    $RootPool.DeleteChildPool($WntdPool)
}
#
# MoveMachine
function movemachine {
    [CmdletBinding()]
    param([Switch]$help, [Parameter(Position=1)][String]$machine, [Parameter(Position=2)][String]$from, [Parameter(Position=3)][String]$to)

    function Usage {
        Write-Output "movemachine:"
        Write-Output ""
        Write-Output "A script that moves a machine from one pool to another."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "movemachine <machine> <frompool> <topool> [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "     help = Shows this message."
        Write-Output ""
        Write-Output "  machine = The name of the machine as registered with the HCK\HLK controller."
        Write-Output ""
        Write-Output " frompool = The name of the source pool."
        Write-Output ""
        Write-Output "   topool = The name of the destination pool."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if ([String]::IsNullOrEmpty($machine)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a machine's name."
            Usage; return
        } else {
            throw "Please provide a machine's name."
        }
    }
    if ([String]::IsNullOrEmpty($from)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a source pool's name."
            Usage; return
        } else {
            throw "Please provide a source pool's name."
        }
    }
    if ([String]::IsNullOrEmpty($to)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a destination pool's name."
            Usage; return
        } else {
            throw "Please provide a destination pool's name."
        }
    }

    if (-Not ($WntdFromPool = $RootPool.GetChildPools() | Where-Object { $_.Name -eq $from })) { throw "Provided source pool's name is not valid, aborting..." }
    if (-Not ($WntdToPool = $RootPool.GetChildPools() | Where-Object { $_.Name -eq $to })) { throw "Provided destination pool's name is not valid, aborting..." }
    if (-Not ($WntdMachine = $WntdFromPool.GetMachines() | Where-Object { $_.Name -eq $machine })) { throw "Provided machines's name is not valid, aborting..." }

    if (-Not $json) { Write-Output "Moving machine $($WntdMachine.Name) from $($WntdFromPool.Name) to $($WntdToPool.Name) pool." }
    $WntdFromPool.MoveMachineTo($WntdMachine, $WntdToPool)
}
#
# SetMachineState
function setmachinestate {
    [CmdletBinding()]
    param([Switch]$help, [Int]$timeout = -1, [Parameter(Position=1)][String]$machine, [Parameter(Position=2)][String]$pool, [Parameter(Position=3)][String]$state)

    function Usage {
        Write-Output "setmachinestate:"
        Write-Output ""
        Write-Output "A script that sets the state of a machine to Ready or NotReady."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "setmachinestate <machine> <poolname> <state> [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "     help = Shows this message."
        Write-Output ""
        Write-Output "  machine = The name of the machine as registered with the HCK\HLK controller."
        Write-Output ""
        Write-Output " poolname = The name of the pool."
        Write-Output ""
        Write-Output "    state = The state, Ready or NotReady."
        Write-Output ""
        Write-Output "  timeout = The operation's timeout in seconds, disabled by default."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if ([String]::IsNullOrEmpty($machine)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a machine's name."
            Usage; return
        } else {
            throw "Please provide a machine's name."
        }
    }
    if ([String]::IsNullOrEmpty($pool)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a pool's name."
            Usage; return
        } else {
            throw "Please provide a pool's name."
        }
    }
    if ([String]::IsNullOrEmpty($state)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a state."
            Usage; return
        } else {
            throw "Please provide a state."
        }
    }

    if (-Not ($WntdPool = $RootPool.GetChildPools() | Where-Object { $_.Name -eq $pool })) { throw "Provided pool's name is not valid, aborting..." }
    if (-Not ($WntdMachine = $WntdPool.GetMachines() | Where-Object { $_.Name -eq $machine })) { throw "Provided machines's name is not valid, aborting..." }
    if (-Not ($timeout -eq -1)) { $timeout = $timeout * 1000 }

    if (-Not $json) { Write-Output "Setting machine $($WntdMachine.Name) to $state state..." }
    switch ($state) {
        "Ready" {
            if (-Not $WntdMachine.SetMachineStatus([Microsoft.Windows.Kits.Hardware.ObjectModel.MachineStatus]::Ready, $timeout)) { throw "Unable to change machine state, timed out." }
        }
        "NotReady" {
            if (-Not $WntdMachine.SetMachineStatus([Microsoft.Windows.Kits.Hardware.ObjectModel.MachineStatus]::NotReady, $timeout))  { throw "Unable to change machine state, timed out." }
        }
        default {
            throw "Provided desired machines's sate is not valid, aborting..."
        }
    }
}
#
# DeleteMachine
function deletemachine {
    [CmdletBinding()]
    param([Switch]$help, [Parameter(Position=1)][String]$machine, [Parameter(Position=2)][String]$pool)

    function Usage {
        Write-Output "deletemachine:"
        Write-Output ""
        Write-Output "A script that deletes a machine."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "deletemachine <machine> <poolname> [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "     help = Shows this message."
        Write-Output ""
        Write-Output "  machine = The name of the machine as registered with the HCK\HLK controller."
        Write-Output ""
        Write-Output " poolname = The name of the pool."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if ([String]::IsNullOrEmpty($machine)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a machine's name."
            Usage; return
        } else {
            throw "Please provide a machine's name."
        }
    }
    if ([String]::IsNullOrEmpty($pool)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a pool's name."
            Usage; return
        } else {
            throw "Please provide a pool's name."
        }
    }

    if (-Not ($WntdPool = $RootPool.GetChildPools() | Where-Object { $_.Name -eq $pool })) { throw "Provided pool's name is not valid, aborting..." }

    if (-Not $json) { Write-Output "Deleting machine $machine..." }
    $WntdPool.DeleteMachine($machine)
}
#
# ListMachineTargets
function listmachinetargets {
    [CmdletBinding()]
    param([Switch]$help, [Parameter(Position=1)][String]$machine, [Parameter(Position=2)][String]$pool)

    function Usage {
        Write-Output "listmachinetargets:"
        Write-Output ""
        Write-Output "A script that lists the target devices of a machine that are available to be tested."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "listmachientargets <testmachine> <poolname> [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "        help = Shows this message."
        Write-Output ""
        Write-Output " testmachine = The name of the machine as registered with the HCK\HLK controller."
        Write-Output "               NOTE: test machine should be in a READY state."
        Write-Output ""
        Write-Output "   poolname  = The name of the pool."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if ([String]::IsNullOrEmpty($machine)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a machine's name."
            Usage; return
        } else {
            throw "Please provide a machine's name."
        }
    }
    if ([String]::IsNullOrEmpty($pool)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a pool's name."
            Usage; return
        } else {
            throw "Please provide a pool's name."
        }
    }

    if (-Not ($WntdPool = $RootPool.GetChildPools() | Where-Object { $_.Name -eq $pool })) { throw "Did not find pool $pool in Root pool, aborting..." }
    if (-Not ($WntdMachine = $WntdPool.GetMachines()| Where-Object { $_.Name -eq $machine })) { throw "The test machine was not found, aborting..." }

    if (-Not $json) {
        Write-Output ""
        Write-Output "The tests targets on $($WntdMachine.Name) are:"
        Write-Output ""

        foreach ($TestTarget in $WntdMachine.GetTestTargets()) {
            Write-Output "============================================="
            Write-Output "Target name : $($TestTarget.Name)"
            Write-Output ""
            Write-Output "    Key  : $($TestTarget.Key)"
            Write-Output "    Type : $($TestTarget.TargetType)"
            Write-Output ""
            Write-Output "============================================="
        }
    } else {
        $targetslist = New-Object System.Collections.ArrayList
        foreach ($TestTarget in $WntdMachine.GetTestTargets()) {
            $targetslist.Add((New-Target $TestTarget.Name $TestTarget.Key)) | Out-Null
        }
        ConvertTo-Json @($targetslist) -Compress
    }
}
#
# ListProjects
function listprojects {
    [CmdletBinding()]
    param([Switch]$help)

    function Usage {
        Write-Output "listprojects:"
        Write-Output ""
        Write-Output "A script that lists the projects info."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "listprojects [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "      help = Shows this message."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if (-Not $json) {
        foreach ($ProjectName in $Manager.GetProjectNames()) {
            $Project = $Manager.GetProject($ProjectName)
            Write-Output "============================================="
            Write-Output "Project name : $($Project.Name)"

            Write-Output ""
            Write-Output "    Creation time : $($Project.CreationTime)"
            Write-Output "    Modified time : $($Project.ModifiedTime)"
            Write-Output "    Status        : $($Project.Info.Status)"
            Write-Output ""
            $ProductInstances = $Project.GetProductInstances()
            if ($ProductInstances.Count -lt 1) {
                Write-Output "    No product instances!"
            } else {
                Write-Output "    Product instances :"
                foreach ($Pi in $ProductInstances) {
                    Write-Output "        Name          : $($Pi.Name)"
                    Write-Output "        OSPlatform    : $($Pi.OSPlatform.Name)"
                    Write-Output "        Targeted pool : $($Pi.MachinePool.Name)"
                    Write-Output "        Targets       :"
                    foreach ($Target in $Pi.GetTargets()) {
                        Write-Output "            Name    : $($Target.Name)"
                        Write-Output "            Key     : $($Target.Key)"
                        Write-Output "            Type    : $($Target.TargetType)"
                        Write-Output "            Machine : $($Target.Machine.Name)"
                        Write-Output ""
                    }
                }
            }

            Write-Output "============================================="
        }
    } else {
        $projectslist = New-Object System.Collections.ArrayList
        foreach ($ProjectName in $Manager.GetProjectNames()) {
            $Project = $Manager.GetProject($ProjectName)
            $ProductInstances = $Project.GetProductInstances()
            $productinstanceslist = New-Object System.Collections.ArrayList
            foreach ($Pi in $ProductInstances) {
                $targetslist = New-Object System.Collections.ArrayList
                foreach ($Target in $Pi.GetTargets()) {
                    $targetslist.Add((New-ProductInstanceTarget $Target.Name $Target.Key $Target.Machine.Name)) | Out-Null
                }
                $productinstanceslist.Add((New-ProductInstance $Pi.Name $Pi.OSPlatform.Name $Pi.MachinePool.Name $targetslist)) | Out-Null
            }
            $projectslist.Add((New-Project $Project.Name $Project.CreationTime.ToString() $Project.ModifiedTime.ToString() $Project.Info.Status.ToString() $productinstanceslist)) | Out-Null
        }
        ConvertTo-Json @($projectslist) -Depth 5 -Compress
    }
}
#
# CreateProject
function createproject {
    [CmdletBinding()]
    param([Switch]$help, [Parameter(Position=1)][String]$project)

    function Usage {
        Write-Output "createproject:"
        Write-Output ""
        Write-Output "A script that creates a project."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "createproject <projectname> [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "               help = Shows this message."
        Write-Output ""
        Write-Output "        projectname = The name of the project."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if ([String]::IsNullOrEmpty($project)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a project's name."
            Usage; return
        } else {
            throw "Please provide a project's name."
        }
    }

    if ($Manager.GetProjectNames().Contains($project)) {
        throw "A project with the name $($project) already exists, aborting..."
    } else {
        if (-Not $json) { Write-Output "Creating a new project named $($project)." }
        $WntdProject = $Manager.CreateProject($project)
    }
}
#
# DeleteProject
function deleteproject {
    [CmdletBinding()]
    param([Switch]$help, [Parameter(Position=1)][String]$project)

    function Usage {
        Write-Output "deleteproject:"
        Write-Output ""
        Write-Output "A script that deletes a project."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "deleteproject <projectname> [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "               help = Shows this message."
        Write-Output ""
        Write-Output "        projectname = The name of the project."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if ([String]::IsNullOrEmpty($project)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a project's name."
            Usage; return
        } else {
            throw "Please provide a project's name."
        }
    }

    if (-Not $json) { Write-Output "Deleting project $project..." }
    $Manager.DeleteProject($project)
}
#
# CreateProjectTarget
function createprojecttarget {
    [CmdletBinding()]
    param([Switch]$help, [Parameter(Position=1)][String]$target, [Parameter(Position=2)][String]$project, [Parameter(Position=3)][String]$machine, [Parameter(Position=4)][String]$pool)

    function Usage {
        Write-Output "createprojecttarget:"
        Write-Output ""
        Write-Output "A script that creates a project's target."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "createprojecttarget <targetkey> <projectname> <testmachine> <poolname> [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "        help = Shows this message."
        Write-Output ""
        Write-Output "    tagetkey = The key of the target, use listmachinetargets to get it."
        Write-Output ""
        Write-Output " projectname = The name of the project."
        Write-Output ""
        Write-Output " testmachine = The name of the machine as registered with the HCK\HLK controller."
        Write-Output "               NOTE: test machine should be in a READY state."
        Write-Output ""
        Write-Output "    poolname = The name of the pool."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if ([String]::IsNullOrEmpty($target)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a target's key."
            Usage; return
        } else {
            throw "Please provide a target's key."
        }
    }
    if ([String]::IsNullOrEmpty($project)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a project's name."
            Usage; return
        } else {
            throw "Please provide a project's name."
        }
    }
    if ([String]::IsNullOrEmpty($machine)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a machine's name."
            Usage; return
        } else {
            throw "Please provide a machine's name."
        }
    }
    if ([String]::IsNullOrEmpty($pool)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a pool's name."
            Usage; return
        } else {
            throw "Please provide a pool's name."
        }
    }

    if (-Not ($WntdPool = $RootPool.GetChildPools() | Where-Object { $_.Name -eq $pool })) { throw "Did not find pool $pool in Root pool, aborting..." }
    if (-Not ($WntdMachine = $WntdPool.GetMachines()| Where-Object { $_.Name -eq $machine })) { throw "The test machine was not found, aborting..." }
    if (-Not ($WntdTarget = $WntdMachine.GetTestTargets() | Where-Object { $_.Key -eq $target })) { throw "A target that matches the target's key given was not found in the specified machine, aborting..." }
    if (-Not ($Manager.GetProjectNames().Contains($project))) { throw "No project with the given name was found, aborting..." } else { $WntdProject = $Manager.GetProject($project) }
    $CreatedPI = $false
    if (-Not ($WntdPI = $WntdProject.GetProductInstances() | Where-Object { $_.OSPlatform -eq $WntdMachine.OSPlatform })) {
        if (-Not $WntdProject.CanCreateProductInstance($WntdMachine.OSPlatform.Description, $WntdPool, $WntdMachine.OSPlatform)) {
            throw "Can't create the project's product instance, it may be due to the project having another product instance that matches the wanted machine's pool or platform."
        } else {
            $WntdPI = $WntdProject.CreateProductInstance($WntdMachine.OSPlatform.Description, $WntdPool, $WntdMachine.OSPlatform)
            $CreatedPI = $true
        }
    }

    try {
        $WntdPITargets = $WntdPI.GetTargets()
        if (($WntdTarget.TargetType -eq "System") -and ($WntdPITargets | Where-Object { $_.TargetType -ne "System" })) { throw "The project already has non-system targets, can't mix system and non-system targets, aborting..." }
        if (($WntdTarget.TargetType -ne "System") -and ($WntdPITargets | Where-Object { $_.TargetType -eq "System" })) { throw "The project already has system targets, can't mix system and non-system targets, aborting..." }
        else {
            if (-Not $json) { Write-Output "Creating a new project's target from $($WntdTarget.Name)." }
            $WntdtoTarget = New-Object System.Collections.ArrayList
            if ($WntdTarget.TargetType -eq "TargetCollection") {
                foreach ($toTarget in $WntdPI.FindTargetFromContainer($WntdTarget.ContainerId)) {
                    if ($toTarget.Machine.Equals($WntdMachine)){
                        $WntdtoTarget.Add($toTarget) | Out-Null
                    }
                }
            } else {
                $WntdtoTarget.Add($WntdTarget) | Out-Null
            }
            foreach ($toTarget in $WntdtoTarget) {
                if ($WntdPITargets | Where-Object { $_.Key -eq $toTarget.Key }) { throw "The target is already being targeted in the project, aborting..." }
                switch ($toTarget.TargetType) {
                    "Filter" { [String[]]$HardwareIds = $toTarget.Key }
                    "System" { [String[]]$HardwareIds = "[SYSTEM]" }
                    default { [String[]]$HardwareIds = $toTarget.HardwareId }
                }
                if (-Not ($WntdDeviceFamily = $Manager.GetDeviceFamilies() | Where-Object { $_.Name -eq $HardwareIds[0] })) {
                    $WntdDeviceFamily = $Manager.CreateDeviceFamily($HardwareIds[0], $HardwareIds)
                }
                $WntdTargetFamily = $WntdPI.CreateTargetFamily($WntdDeviceFamily)
                $WntdTargetFamily.CreateTarget($toTarget) | Out-Null
            }
        }
    } catch {
        if ($CreatedPI) { $WntdProject.DeleteProductInstance($WntdMachine.OSPlatform.Description) }
        throw
    }
}
#
# DeleteProjectTarget
function deleteprojecttarget {
    [CmdletBinding()]
    param([Switch]$help, [Parameter(Position=1)][String]$target, [Parameter(Position=2)][String]$project, [Parameter(Position=3)][String]$machine, [Parameter(Position=4)][String]$pool)

    function Usage {
        Write-Output "deleteprojecttarget:"
        Write-Output ""
        Write-Output "A script that deletes a project's target."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "deleteprojecttarget <targetkey> <projectname> <testmachine> <poolname> [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "        help = Shows this message."
        Write-Output ""
        Write-Output "    tagetkey = The key of the target, use listmachinetargets to get it."
        Write-Output ""
        Write-Output " projectname = The name of the project."
        Write-Output ""
        Write-Output " testmachine = The name of the machine as registered with the HCK\HLK controller."
        Write-Output "               NOTE: test machine should be in a READY state."
        Write-Output ""
        Write-Output "    poolname = The name of the pool."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if ([String]::IsNullOrEmpty($target)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a target's key."
            Usage; return
        } else {
            throw "Please provide a target's key."
        }
    }
    if ([String]::IsNullOrEmpty($project)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a project's name."
            Usage; return
        } else {
            throw "Please provide a project's name."
        }
    }
    if ([String]::IsNullOrEmpty($machine)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a machine's name."
            Usage; return
        } else {
            throw "Please provide a machine's name."
        }
    }
    if ([String]::IsNullOrEmpty($pool)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a pool's name."
            Usage; return
        } else {
            throw "Please provide a pool's name."
        }
    }

    if (-Not ($WntdPool = $RootPool.GetChildPools() | Where-Object { $_.Name -eq $pool })) { throw "Did not find pool $pool in Root pool, aborting..." }
    if (-Not ($WntdMachine = $WntdPool.GetMachines()| Where-Object { $_.Name -eq $machine })) { throw "The test machine was not found, aborting..." }
    if (-Not ($WntdTarget = $WntdMachine.GetTestTargets() | Where-Object { $_.Key -eq $target })) { throw "A target that matches the target's key given was not found in the specified machine, aborting..." }
    if (-Not ($Manager.GetProjectNames().Contains($project))) { throw "No project with the given name was found, aborting..." } else { $WntdProject = $Manager.GetProject($project) }
    if (-Not ($WntdPI = $WntdProject.GetProductInstances() | Where-Object { $_.OSPlatform -eq $WntdMachine.OSPlatform })) { throw "Machine pool not targeted in the project." }

    if (-Not $json) { Write-Output "Deleting a new project's target from $($WntdTarget.Name)." }
    $WntdPI.DeleteTarget($WntdTarget.Key, $WntdTarget.Machine)
    if ($WntdPI.GetTargets().Count -lt 1) { $WntdProject.DeleteProductInstance($WntdPI.Name) }
}
#
# ListTests
function listtests {
    [CmdletBinding()]
    param([Switch]$help, [Switch]$manual, [Switch]$auto, [Switch]$failed, [Switch]$inqueue, [Switch]$notrun, [Switch]$passed, [Switch]$running, [String]$playlist, [Parameter(Position=1)][String]$target, [Parameter(Position=2)][String]$project, [Parameter(Position=3)][String]$machine, [Parameter(Position=4)][String]$pool)

    function Usage {
        Write-Output "listtests:"
        Write-Output ""
        Write-Output "A script that lists a project target's tests."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "listtests <targetkey> <projectname> <testmachine> <poolname> [-manual]"
        Write-Output "                           [-auto] [-failed] [-inqueue] [-notrun] [-passed] [-running]"
        Write-Output "                               [-playlist] [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "        help = Shows this message."
        Write-Output ""
        Write-Output "    tagetkey = The key of the target, use listmachinetargets to get it."
        Write-Output ""
        Write-Output " projectname = The name of the project."
        Write-Output ""
        Write-Output " testmachine = The name of the machine as registered with the HCK\HLK controller."
        Write-Output "               NOTE: test machine should be in a READY state."
        Write-Output ""
        Write-Output "    poolname = The name of the pool."
        Write-Output ""
        Write-Output "    playlist = List only the tests that matches the given playlist, (by path)."
        Write-Output ""
        Write-Output "      manual = List only the manual run tests."
        Write-Output ""
        Write-Output "        auto = List only the auto run tests."
        Write-Output ""
        Write-Output "      failed = List only the failed tests."
        Write-Output ""
        Write-Output "     inqueue = List only the tests that are in the run queue."
        Write-Output ""
        Write-Output "      notrun = List only the tests that haven't been run."
        Write-Output ""
        Write-Output "      passed = List only the passed tests."
        Write-Output ""
        Write-Output "     running = List only the running tests."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if ([String]::IsNullOrEmpty($target)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a target's key."
            Usage; return
        } else {
            throw "Please provide a target's key."
        }
    }
    if ([String]::IsNullOrEmpty($project)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a project's name."
            Usage; return
        } else {
            throw "Please provide a project's name."
        }
    }
    if ([String]::IsNullOrEmpty($machine)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a machine's name."
            Usage; return
        } else {
            throw "Please provide a machine's name."
        }
    }
    if ([String]::IsNullOrEmpty($pool)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a pool's name."
            Usage; return
        } else {
            throw "Please provide a pool's name."
        }
    }
    if ((-Not [String]::IsNullOrEmpty($playlist)) -and $Studio -ne "hlk") {
        if (-Not $json) {
            Write-Output "WARNING: Playlist provided but HCK doesn't support playlists, aborting..."
            Usage; return
        } else {
            throw "Playlist provided but HCK doesn't support playlists, aborting..."
        }
    }

    if (-Not ($manual -or $auto)) {
        $manual = $true
        $auto = $true
    }
    if (-Not ($notrun -or $failed -or $passed -or $running -or $inqueue)) {
        $notrun = $true
        $failed = $true
        $passed = $true
        $running = $true
        $inqueue = $true
    }

    if (-Not ($WntdPool = $RootPool.GetChildPools() | Where-Object { $_.Name -eq $pool })) { throw "Did not find pool $pool in Root pool, aborting..." }
    if (-Not ($WntdMachine = $WntdPool.GetMachines()| Where-Object { $_.Name -eq $machine })) { throw "The test machine was not found, aborting..." }
    if (-Not ($WntdTarget = $WntdMachine.GetTestTargets() | Where-Object { $_.Key -eq $target })) { throw "A target that matches the target's key given was not found in the specified machine, aborting..." }
    if (-Not ($Manager.GetProjectNames().Contains($project))) { throw "No project with the given name was found, aborting..." } else { $WntdProject = $Manager.GetProject($project) }
    if (-Not ($WntdPI = $WntdProject.GetProductInstances() | Where-Object { $_.OSPlatform -eq $WntdMachine.OSPlatform })) { throw "Machine pool not targeted in the project." }

    $WntdPITargets = New-Object System.Collections.ArrayList

    if ($WntdTarget.TargetType -eq "TargetCollection") {
        if (-Not ($WntdPITargetsToAdd = $WntdPI.GetTargets() | Where-Object { ($_.ContainerId -eq $WntdTarget.ContainerId) -and ($_.Machine.Equals($WntdMachine)) })) { throw "The target is not being targeted by the project." }
        $WntdPITargets.AddRange($WntdPITargetsToAdd)
    } else {
        if (-Not ($WntdPITarget = $WntdPI.GetTargets() | Where-Object { ($_.Key -eq $WntdTarget.Key) -and ($_.Machine.Equals($WntdMachine)) })) { throw "The target is not being targeted by the project." }
        $WntdPITargets.Add($WntdPITarget) | Out-Null
    }

    $WntdTests = New-Object System.Collections.ArrayList

    if (-Not [String]::IsNullOrEmpty($playlist)) {
        $PlaylistManager = New-Object Microsoft.Windows.Kits.Hardware.ObjectModel.PlaylistManager $WntdProject
        $WntdPlaylist = [Microsoft.Windows.Kits.Hardware.ObjectModel.PlaylistManager]::DeserializePlaylist($playlist)
        foreach ($tTest in $PlaylistManager.GetTestsFromProjectThatMatchPlaylist($WntdPlaylist)) {
            if ($tTest.GetTestTargets() | Where-Object { $WntdPITargets.Contains($_) }) { $WntdTests.Add($tTest) | Out-Null }
        }
    } else {
         $WntdPITargets | foreach { $WntdTests.AddRange($_.GetTests()) }
    }

    if (-Not $json) {
        Write-Output ""
        Write-Output "The requested project project target's tests:"
        Write-Output ""

        foreach ($tTest in $WntdTests) {
            if (-Not (($manual -and ($tTest.TestType -eq "Manual")) -or ($auto -and ($tTest.TestType -eq "Automated")))) {
                continue
            } elseif (-Not (($notrun -and ($tTest.Status -eq "NotRun")) -or ($failed -and ($tTest.Status -eq "Failed")) -or ($passed -and ($tTest.Status -eq "Passed")) -or ($running -and ($tTest.ExecutionState -eq "Running")) -or ($inqueue -and ($tTest.ExecutionState -eq "InQueue")))) {
                continue
            }
            Write-Output "============================================="
            Write-Output "Test name : $($tTest.Name)"
            Write-Output ""
            Write-Output "    Test id                        : $($tTest.Id)"
            Write-Output "    Test type                      : $($tTest.TestType)"
            Write-Output "    Estimated runtime              : $($tTest.EstimatedRuntime)"
            Write-Output "    Requires special configuration : $($tTest.RequiresSpecialConfiguration)"
            Write-Output "    Requires supplemental content  : $($tTest.RequiresSupplementalContent)"
            Write-Output "    Schedule options               : $($tTest.ScheduleOptions)"
            Write-Output "    Test status                    : $($tTest.Status)"
            Write-Output "    Execution State                : $($tTest.ExecutionState)"
            Write-Output ""
            Write-Output "============================================="
        }
    } else {
        $testslist = New-Object System.Collections.ArrayList
        foreach ($tTest in $WntdTests) {
            if (-Not (($manual -and ($tTest.TestType -eq "Manual")) -or ($auto -and ($tTest.TestType -eq "Automated")))) {
                continue
            } elseif (-Not (($notrun -and ($tTest.Status -eq "NotRun")) -or ($failed -and ($tTest.Status -eq "Failed")) -or ($passed -and ($tTest.Status -eq "Passed")) -or ($running -and ($tTest.ExecutionState -eq "Running")) -or ($inqueue -and ($tTest.ExecutionState -eq "InQueue")))) {
                continue
            }
            $testslist.Add((New-Test $tTest.Name $tTest.Id $tTest.TestType.ToString() $tTest.EstimatedRuntime.ToString() $tTest.RequiresSpecialConfiguration.ToString() $tTest.RequiresSupplementalContent.ToString() ($tTest.ScheduleOptions.ToString() -split ', ') $tTest.Status.ToString() $tTest.ExecutionState.ToString())) | Out-Null
        }
        ConvertTo-Json @($testslist) -Compress
    }
}
#
# GetTestInfo
function gettestinfo {
    [CmdletBinding()]
    param([Switch]$help, [Parameter(Position=1)][String]$test, [Parameter(Position=2)][String]$target, [Parameter(Position=3)][String]$project, [Parameter(Position=4)][String]$machine, [Parameter(Position=5)][String]$pool)

    function Usage {
        Write-Output "gettestinfo:"
        Write-Output ""
        Write-Output "A script that gets a project target's test info."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "gettestinfo <testid> <targetkey> <projectname> <testmachine> <poolname> [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "        help = Shows this message."
        Write-Output ""
        Write-Output "      testid = The id of the test, use listtests action to get it."
        Write-Output ""
        Write-Output "    tagetkey = The key of the target, use listmachinetargets to get it."
        Write-Output ""
        Write-Output " projectname = The name of the project."
        Write-Output ""
        Write-Output " testmachine = The name of the machine as registered with the HCK\HLK controller."
        Write-Output "               NOTE: test machine should be in a READY state."
        Write-Output ""
        Write-Output "    poolname = The name of the pool."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if ([String]::IsNullOrEmpty($test)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a test's id."
            Usage; return
        } else {
            throw "Please provide a test's id."
        }
    }
    if ([String]::IsNullOrEmpty($target)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a target's key."
            Usage; return
        } else {
            throw "Please provide a target's key."
        }
    }
    if ([String]::IsNullOrEmpty($project)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a project's name."
            Usage; return
        } else {
            throw "Please provide a project's name."
        }
    }
    if ([String]::IsNullOrEmpty($machine)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a machine's name."
            Usage; return
        } else {
            throw "Please provide a machine's name."
        }
    }
    if ([String]::IsNullOrEmpty($pool)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a pool's name."
            Usage; return
        } else {
            throw "Please provide a pool's name."
        }
    }

    if (-Not ($WntdPool = $RootPool.GetChildPools() | Where-Object { $_.Name -eq $pool })) { throw "Did not find pool $pool in Root pool, aborting..." }
    if (-Not ($WntdMachine = $WntdPool.GetMachines()| Where-Object { $_.Name -eq $machine })) { throw "The test machine was not found, aborting..." }
    if (-Not ($WntdTarget = $WntdMachine.GetTestTargets() | Where-Object { $_.Key -eq $target })) { throw "A target that matches the target's key given was not found in the specified machine, aborting..." }
    if (-Not ($Manager.GetProjectNames().Contains($project))) { throw "No project with the given name was found, aborting..." } else { $WntdProject = $Manager.GetProject($project) }
    if (-Not ($WntdPI = $WntdProject.GetProductInstances() | Where-Object { $_.OSPlatform -eq $WntdMachine.OSPlatform })) { throw "Machine pool not targeted in the project." }
    if (-Not ($WntdPITarget = $WntdPI.GetTargets() | Where-Object { ($_.Key -eq $WntdTarget.Key) -and ($_.Machine.Equals($WntdMachine)) })) { throw "The target is not being targeted by the project." }
    if (-Not ($WntdTest = $WntdPITarget.GetTests() | Where-Object { $_.Id -eq $test })) { throw "Didn't find a test with the id given." }

    if (-Not $json) {
        Write-Output ""
        Write-Output "The requested project project target's test:"
        Write-Output ""
        Write-Output "============================================="
        Write-Output "Test name : $($WntdTest.Name)"
        Write-Output ""
        Write-Output "    Test id                        : $($WntdTest.Id)"
        Write-Output "    Test type                      : $($WntdTest.TestType)"
        Write-Output "    Estimated runtime              : $($WntdTest.EstimatedRuntime)"
        Write-Output "    Requires special configuration : $($WntdTest.RequiresSpecialConfiguration)"
        Write-Output "    Requires supplemental content  : $($WntdTest.RequiresSupplementalContent)"
        Write-Output "    Schedule options               : $($WntdTest.ScheduleOptions)"
        Write-Output "    Test status                    : $($WntdTest.Status)"
        Write-Output "    Execution State                : $($WntdTest.ExecutionState)"
        Write-Output ""
        Write-Output "============================================="
    } else {
        @((New-Test $WntdTest.Name $WntdTest.Id $WntdTest.TestType.ToString() $WntdTest.EstimatedRuntime.ToString() $WntdTest.RequiresSpecialConfiguration.ToString() $WntdTest.RequiresSupplementalContent.ToString() ($WntdTest.ScheduleOptions.ToString() -split ', ') $WntdTest.Status.ToString() $WntdTest.ExecutionState.ToString())) | ConvertTo-Json -Compress
    }
}
#
# QueueTest
function queuetest {
    [CmdletBinding()]
    param([Switch]$help, [String]$sup, [String]$IPv6, [Parameter(Position=1)][String]$test, [Parameter(Position=2)][String]$target, [Parameter(Position=3)][String]$project, [Parameter(Position=4)][String]$machine, [Parameter(Position=5)][String]$pool)

    function Usage {
        Write-Output "queuetest:"
        Write-Output ""
        Write-Output "A script that queues a test, use listtestresults action to get the results."
        Write-Output "(if the test needs two machines to run use -sup flag)"
        Write-Output "(if the test needs the IPv6 address of the support machine use -IPv6 flag)"
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "queuetest <testid> <targetkey> <projectname> <testmachine> <poolname> [-sup <name>]"
        Write-Output "              [-IPv6 <address>] [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "        help = Shows this message."
        Write-Output ""
        Write-Output "      testid = The id of the test, use listtests action to get it."
        Write-Output ""
        Write-Output "    tagetkey = The key of the target, use listmachinetargets to get it."
        Write-Output ""
        Write-Output " projectname = The name of the project."
        Write-Output ""
        Write-Output " testmachine = The name of the machine as registered with the HCK\HLK controller."
        Write-Output "               NOTE: test machine should be in a READY state."
        Write-Output ""
        Write-Output "    poolname = The name of the pool."
        Write-Output ""
        Write-Output "        IPv6 = The support machines's ""SupportDevice0"" IPv6 address."
        Write-Output ""
        Write-Output "         sup = The support machine's name as registered with the HCK\HLK controller."
        Write-Output "               NOTE: test machine should be in a READY state."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if ([String]::IsNullOrEmpty($test)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a test's id."
            Usage; return
        } else {
            throw "Please provide a test's id."
        }
    }
    if ([String]::IsNullOrEmpty($target)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a target's key."
            Usage; return
        } else {
            throw "Please provide a target's key."
        }
    }
    if ([String]::IsNullOrEmpty($project)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a project's name."
            Usage; return
        } else {
            throw "Please provide a project's name."
        }
    }
    if ([String]::IsNullOrEmpty($machine)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a machine's name."
            Usage; return
        } else {
            throw "Please provide a machine's name."
        }
    }
    if ([String]::IsNullOrEmpty($pool)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a pool's name."
            Usage; return
        } else {
            throw "Please provide a pool's name."
        }
    }

    if (-Not ($WntdPool = $RootPool.GetChildPools() | Where-Object { $_.Name -eq $pool })) { throw "Did not find pool $pool in Root pool, aborting..." }
    if (-Not ($WntdMachine = $WntdPool.GetMachines()| Where-Object { $_.Name -eq $machine })) { throw "The test machine was not found, aborting..." }
    if (-Not ($WntdTarget = $WntdMachine.GetTestTargets() | Where-Object { $_.Key -eq $target })) { throw "A target that matches the target's key given was not found in the specified machine, aborting..." }
    if (-Not ($Manager.GetProjectNames().Contains($project))) { throw "No project with the given name was found, aborting..." } else { $WntdProject = $Manager.GetProject($project) }
    if (-Not ($WntdPI = $WntdProject.GetProductInstances() | Where-Object { $_.OSPlatform -eq $WntdMachine.OSPlatform })) { throw "Machine pool not targeted in the project." }
    if (-Not ($WntdPITarget = $WntdPI.GetTargets() | Where-Object { ($_.Key -eq $WntdTarget.Key) -and ($_.Machine.Equals($WntdMachine)) })) { throw "The target is not being targeted by the project." }
    if (-Not ($WntdTest = $WntdPITarget.GetTests() | Where-Object { $_.Id -eq $test })) { throw "Didn't find a test with the id given." }

    if (-Not $json) { Write-Output "Queueing test $($WntdTest.Name)..." }

    if (-Not [String]::IsNullOrEmpty($IPv6)) {
        $WntdTest.SetParameter("WDTFREMOTESYSTEM", $IPv6, [Microsoft.Windows.Kits.Hardware.ObjectModel.ParameterSetAsDefault]::DoNotSetAsDefault) | Out-Null
    }

    if (-Not [String]::IsNullOrEmpty($sup)) {
        if (-Not ($WntdSMachine = $WntdPool.GetMachines()| Where-Object { $_.Name -eq $sup })) { throw "The support machine was not found, aborting..." }
        $MachineSet = $WntdTest.GetMachineRole()
        foreach ($Role in $MachineSet.Roles) {
            if ($Role.Name -eq "Support") {
                $Role.AddMachine($WntdSMachine)
            }
        }
        $MachineSet.ApplyMachineDimensions()
        $WntdTest.QueueTest($MachineSet) | Out-Null
    } else {
        $WntdTest.QueueTest() | Out-Null
    }
}
#
# ApplyProjectFilters
function applyprojectfilters {
    [CmdletBinding()]
    param([Switch]$help, [Parameter(Position=1)][String]$project)

    function Usage {
        Write-Output "applyprojectfilters:"
        Write-Output ""
        Write-Output "A script that applies the filters on a project's test results."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "applyprojectfilters <projectname> [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "        help = Shows this message."
        Write-Output ""
        Write-Output " projectname = The name of the project."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if ([String]::IsNullOrEmpty($project)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a project's name."
            Usage; return
        } else {
            throw "Please provide a project's name."
        }
    }

    if (-Not ($Manager.GetProjectNames().Contains($project))) { throw "No project with the given name was found, aborting..." } else { $WntdProject = $Manager.GetProject($project) }

    if (-Not $json) { Write-Output "Applying filters on project $($WntdProject.Name)..." }

    $WntdFilterEngine = New-Object Microsoft.Windows.Kits.Hardware.FilterEngine.DatabaseFilterEngine $Manager
    $WntdFilterResultDictionary = $WntdFilterEngine.Filter($WntdProject)
    $Count = 0
    foreach ($tFilterResultCollection in $WntdFilterResultDictionary.Values) {
        $Count += $tFilterResultCollection.Count
    }

    if (-Not $json) {
        Write-Output "Applied filters on $Count tasks."
    } else {
        @(New-FilterResult $Count) | ConvertTo-Json -Compress
    }
}
#
# ApplyTestResultsFilters
function applytestresultfilters {
    [CmdletBinding()]
    param([Switch]$help, [Parameter(Position=1)][String]$result, [Parameter(Position=2)][String]$test, [Parameter(Position=3)][String]$target, [Parameter(Position=4)][String]$project, [Parameter(Position=5)][String]$machine, [Parameter(Position=6)][String]$pool)

    function Usage {
        Write-Output "applytestresultfilters:"
        Write-Output ""
        Write-Output "A script that applies filters on a test result."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "applytestresultfilters <resultindex> <testid> <targetkey> <projectname> <testmachine> <poolname> [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "        help = Shows this message."
        Write-Output ""
        Write-Output " resultindex = The index of the test result, use listtestresults action to get it."
        Write-Output ""
        Write-Output "      testid = The id of the test, use listtests action to get it."
        Write-Output ""
        Write-Output "    tagetkey = The key of the target, use listmachinetargets to get it."
        Write-Output ""
        Write-Output " projectname = The name of the project."
        Write-Output ""
        Write-Output " testmachine = The name of the machine as registered with the HCK\HLK controller."
        Write-Output "               NOTE: test machine should be in a READY state."
        Write-Output ""
        Write-Output "    poolname = The name of the pool."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if ([String]::IsNullOrEmpty($result)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a test result's index."
            Usage; return
        } else {
            throw "Please provide a test result's index."
        }
    }
    if ([String]::IsNullOrEmpty($test)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a test's id."
            Usage; return
        } else {
            throw "Please provide a test's id."
        }
    }
    if ([String]::IsNullOrEmpty($target)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a target's key."
            Usage; return
        } else {
            throw "Please provide a target's key."
        }
    }
    if ([String]::IsNullOrEmpty($project)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a project's name."
            Usage; return
        } else {
            throw "Please provide a project's name."
        }
    }
    if ([String]::IsNullOrEmpty($machine)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a machine's name."
            Usage; return
        } else {
            throw "Please provide a machine's name."
        }
    }
    if ([String]::IsNullOrEmpty($pool)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a pool's name."
            Usage; return
        } else {
            throw "Please provide a pool's name."
        }
    }

    if (-Not ($WntdPool = $RootPool.GetChildPools() | Where-Object { $_.Name -eq $pool })) { throw "Did not find pool $pool in Root pool, aborting..." }
    if (-Not ($WntdMachine = $WntdPool.GetMachines()| Where-Object { $_.Name -eq $machine })) { throw "The test machine was not found, aborting..." }
    if (-Not ($WntdTarget = $WntdMachine.GetTestTargets() | Where-Object { $_.Key -eq $target })) { throw "A target that matches the target's key given was not found in the specified machine, aborting..." }
    if (-Not ($Manager.GetProjectNames().Contains($project))) { throw "No project with the given name was found, aborting..." } else { $WntdProject = $Manager.GetProject($project) }
    if (-Not ($WntdPI = $WntdProject.GetProductInstances() | Where-Object { $_.OSPlatform -eq $WntdMachine.OSPlatform })) { throw "Machine pool not targeted in the project." }
    if (-Not ($WntdPITarget = $WntdPI.GetTargets() | Where-Object { ($_.Key -eq $WntdTarget.Key) -and ($_.Machine.Equals($WntdMachine)) })) { throw "The target is not being targeted by the project." }
    if (-Not ($WntdTest = $WntdPITarget.GetTests() | Where-Object { $_.Id -eq $test })) { throw "Didn't find a test with the id given." }
    if (-Not ($WntdTest.GetTestResults().Count -ge 1)) { throw "The test hasen't been queued, can't find test results." } else { $WntdResult = $WntdTest.GetTestResults()[$result] }

    if (-Not $json) { Write-Output "Applying filters on test result..." }

    $WntdFilterEngine = New-Object Microsoft.Windows.Kits.Hardware.FilterEngine.DatabaseFilterEngine $Manager
    $WntdFilterResultCollection = $WntdFilterEngine.Filter($WntdResult)

    if (-Not $json) {
        Write-Output "Applied filters on $($WntdFilterResultCollection.Count) tasks."
    } else {
        @(New-FilterResult $WntdFilterResultCollection.Count) | ConvertTo-Json -Compress
    }
}
#
# ListTestResults
function listtestresults {
    [CmdletBinding()]
    param([Switch]$help, [Parameter(Position=1)][String]$test, [Parameter(Position=2)][String]$target, [Parameter(Position=3)][String]$project, [Parameter(Position=4)][String]$machine, [Parameter(Position=5)][String]$pool)

    function Usage {
        Write-Output "listtestresults:"
        Write-Output ""
        Write-Output "A script that lists all of the test results and lists them and their info."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "listtestresults <testid> <targetkey> <projectname> <testmachine> <poolname> [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "        help = Shows this message."
        Write-Output ""
        Write-Output "      testid = The id of the test, use listtests action to get it."
        Write-Output ""
        Write-Output "    tagetkey = The key of the target, use listmachinetargets to get it."
        Write-Output ""
        Write-Output " projectname = The name of the project."
        Write-Output ""
        Write-Output " testmachine = The name of the machine as registered with the HCK\HLK controller."
        Write-Output "               NOTE: test machine should be in a READY state."
        Write-Output ""
        Write-Output "    poolname = The name of the pool."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if ([String]::IsNullOrEmpty($test)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a test's id."
            Usage; return
        } else {
            throw "Please provide a test's id."
        }
    }
    if ([String]::IsNullOrEmpty($target)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a target's key."
            Usage; return
        } else {
            throw "Please provide a target's key."
        }
    }
    if ([String]::IsNullOrEmpty($project)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a project's name."
            Usage; return
        } else {
            throw "Please provide a project's name."
        }
    }
    if ([String]::IsNullOrEmpty($machine)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a machine's name."
            Usage; return
        } else {
            throw "Please provide a machine's name."
        }
    }
    if ([String]::IsNullOrEmpty($pool)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a pool's name."
            Usage; return
        } else {
            throw "Please provide a pool's name."
        }
    }

    if (-Not ($WntdPool = $RootPool.GetChildPools() | Where-Object { $_.Name -eq $pool })) { throw "Did not find pool $pool in Root pool, aborting..." }
    if (-Not ($WntdMachine = $WntdPool.GetMachines()| Where-Object { $_.Name -eq $machine })) { throw "The test machine was not found, aborting..." }
    if (-Not ($WntdTarget = $WntdMachine.GetTestTargets() | Where-Object { $_.Key -eq $target })) { throw "A target that matches the target's key given was not found in the specified machine, aborting..." }
    if (-Not ($Manager.GetProjectNames().Contains($project))) { throw "No project with the given name was found, aborting..." } else { $WntdProject = $Manager.GetProject($project) }
    if (-Not ($WntdPI = $WntdProject.GetProductInstances() | Where-Object { $_.OSPlatform -eq $WntdMachine.OSPlatform })) { throw "Machine pool not targeted in the project." }
    if (-Not ($WntdPITarget = $WntdPI.GetTargets() | Where-Object { ($_.Key -eq $WntdTarget.Key) -and ($_.Machine.Equals($WntdMachine)) })) { throw "The target is not being targeted by the project." }
    if (-Not ($WntdTest = $WntdPITarget.GetTests() | Where-Object { $_.Id -eq $test })) { throw "Didn't find a test with the id given." }
    if (-Not ($WntdTest.GetTestResults().Count -ge 1)) { throw "The test hasen't been queued, can't find test results." } else { $WntdResults = $WntdTest.GetTestResults() }

    if (-Not $json) {
        Write-Output ""
        Write-Output "The requested project test's results:"
        Write-Output ""

        foreach ($tTestResult in $WntdResults) {
            $tTestResult.Refresh()
            Write-Output "============================================="
            Write-Output "Test result index : $($WntdResults.IndexOf($tTestResult))"
            Write-Output ""
            Write-Output "    Test name           : $($tTestResult.Test.Name)"
            Write-Output "    Completion time     : $($tTestResult.CompletionTime)"
            Write-Output "    Schedule time       : $($tTestResult.ScheduleTime)"
            Write-Output "    Start time          : $($tTestResult.StartTime)"
            Write-Output "    Status              : $($tTestResult.Status)"
            Write-Output "    Are filters applied : $($tTestResult.AreFiltersApplied)"
            Write-Output "    Target name         : $($tTestResult.Target.Name)"
            Write-Output "    Tasks               :"
            foreach ($tTask in $tTestResult.GetTasks()) {
                Write-Output "        $($tTask.Name):"
                Write-Output "            Stage              : $($tTask.Stage)"
                Write-Output "            Status             : $($tTask.Status)"
                if (-Not [String]::IsNullOrEmpty($tTask.TaskErrorMessage)) {
                    Write-Output "            Task error message : $($tTask.TaskErrorMessage)"
                }
                Write-Output "            Task type          : $($tTask.TaskType)"
                if ($tTask.GetChildTasks()) {
                    Write-Output "            Sub tasks          :"

                    foreach ($subtTask in $tTask.GetChildTasks()) {
                        Write-Output "                $($subtTask.Name):"
                        Write-Output "                    Stage              : $($subtTask.Stage)"
                        Write-Output "                    Status             : $($subtTask.Status)"
                        if (-Not [String]::IsNullOrEmpty($subtTask.TaskErrorMessage)) {
                            Write-Output "                    Task error message : $($subtTask.TaskErrorMessage)"
                        }
                        Write-Output "                    Task type          : $($subtTask.TaskType)"
                        if (-Not ($subtTask -eq $tTask.GetChildTasks()[-1])) {
                            Write-Output ""
                        }
                    }
                }
                Write-Output ""
            }
            Write-Output "============================================="
        }
    } else {
        $testresultlist = New-Object System.Collections.ArrayList

        foreach ($tTestResult in $WntdResults) {
            $tTestResult.Refresh()
            $taskslist = New-Object System.Collections.ArrayList

            foreach ($tTask in $tTestResult.GetTasks()) {
                $subtaskslist = New-Object System.Collections.ArrayList

                if ($tTask.GetChildTasks()) {
                    foreach ($subtTask in $tTask.GetChildTasks()) {
                        $subtasktype = (New-Task $subtTask.Name $subtTask.Stage $subtTask.Status.ToString() $subtTask.TaskErrorMessage $subtTask.TaskType (New-Object System.Collections.ArrayList))
                        $subtaskslist.Add($subtasktype) | Out-Null
                    }
                }
                $tasktype = (New-Task $tTask.Name $tTask.Stage $tTask.Status.ToString() $tTask.TaskErrorMessage $tTask.TaskType $subtaskslist)
                $taskslist.Add($tasktype) | Out-Null
            }

            $testresultlist.Add((New-TestResult $tTestResult.Test.Name $tTestResult.CompletionTime.ToString() $tTestResult.ScheduleTime.ToString() $tTestResult.StartTime.ToString() $tTestResult.Status.ToString() $tTestResult.AreFiltersApplied.ToString() $tTestResult.Target.Name $taskslist)) | Out-Null
        }

        ConvertTo-Json @($testresultlist) -Depth $MaxJsonDepth -Compress
    }
}
#
# ZipTestResultLogs
function ziptestresultlogs {
    [CmdletBinding()]
    param([Switch]$help, [Parameter(Position=1)][String]$result, [Parameter(Position=2)][String]$test, [Parameter(Position=3)][String]$target, [Parameter(Position=4)][String]$project, [Parameter(Position=5)][String]$machine, [Parameter(Position=6)][String]$pool)

    function Usage {
        Write-Output "ziptestresultlogs:"
        Write-Output ""
        Write-Output "A script that zips a test result's logs to the returned zip file path."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "ziptestresultlogs <resultindex> <testid> <targetkey> <projectname> <testmachine> <poolname> [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "        help = Shows this message."
        Write-Output ""
        Write-Output " resultindex = The index of the test result, use listtestresults action to get it."
        Write-Output ""
        Write-Output "      testid = The id of the test, use listtests action to get it."
        Write-Output ""
        Write-Output "    tagetkey = The key of the target, use listmachinetargets to get it."
        Write-Output ""
        Write-Output " projectname = The name of the project."
        Write-Output ""
        Write-Output " testmachine = The name of the machine as registered with the HCK\HLK controller."
        Write-Output "               NOTE: test machine should be in a READY state."
        Write-Output ""
        Write-Output "    poolname = The name of the pool."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if ([String]::IsNullOrEmpty($result)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a test result's index."
            Usage; return
        } else {
            throw "Please provide a test result's index."
        }
    }
    if ([String]::IsNullOrEmpty($test)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a test's id."
            Usage; return
        } else {
            throw "Please provide a test's id."
        }
    }
    if ([String]::IsNullOrEmpty($target)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a target's key."
            Usage; return
        } else {
            throw "Please provide a target's key."
        }
    }
    if ([String]::IsNullOrEmpty($project)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a project's name."
            Usage; return
        } else {
            throw "Please provide a project's name."
        }
    }
    if ([String]::IsNullOrEmpty($machine)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a machine's name."
            Usage; return
        } else {
            throw "Please provide a machine's name."
        }
    }
    if ([String]::IsNullOrEmpty($pool)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a pool's name."
            Usage; return
        } else {
            throw "Please provide a pool's name."
        }
    }

    if (-Not ($WntdPool = $RootPool.GetChildPools() | Where-Object { $_.Name -eq $pool })) { throw "Did not find pool $pool in Root pool, aborting..." }
    if (-Not ($WntdMachine = $WntdPool.GetMachines()| Where-Object { $_.Name -eq $machine })) { throw "The test machine was not found, aborting..." }
    if (-Not ($WntdTarget = $WntdMachine.GetTestTargets() | Where-Object { $_.Key -eq $target })) { throw "A target that matches the target's key given was not found in the specified machine, aborting..." }
    if (-Not ($Manager.GetProjectNames().Contains($project))) { throw "No project with the given name was found, aborting..." } else { $WntdProject = $Manager.GetProject($project) }
    if (-Not ($WntdPI = $WntdProject.GetProductInstances() | Where-Object { $_.OSPlatform -eq $WntdMachine.OSPlatform })) { throw "Machine pool not targeted in the project." }
    if (-Not ($WntdPITarget = $WntdPI.GetTargets() | Where-Object { ($_.Key -eq $WntdTarget.Key) -and ($_.Machine.Equals($WntdMachine)) })) { throw "The target is not being targeted by the project." }
    if (-Not ($WntdTest = $WntdPITarget.GetTests() | Where-Object { $_.Id -eq $test })) { throw "Didn't find a test with the id given." }
    if (-Not ($WntdTest.GetTestResults().Count -ge 1)) { throw "The test hasen't been queued, can't find test results." } else { $WntdResult = $WntdTest.GetTestResults()[$result] }

    $DayStamp = $(get-date).ToString("dd-MM-yyyy")
    $TimeStamp = $(get-date).ToString("hh_mm_ss")

    $LogsDir = $env:TEMP + "\prometheus_test_logs\$DayStamp\[$TimeStamp]" + $WntdTest.Id
    $ZipPath = $env:TEMP + "\prometheus_test_logs\$DayStamp\$DayStamp" + "_" + $TimeStamp + "_" + $WntdTest.Id + ".zip"

    if (-Not $json) {
        Write-Output "The test has $($WntdResult.Status)!."
        Write-Output "Logs zipped to $ZipPath"
    }
    foreach ($Log in $WntdResult.GetLogs()) {
        $Log.WriteLogTo([System.IO.Path]::Combine($LogsDir, $Log.LogType, $Log.Name))
    }
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [IO.Compression.ZipFile]::CreateFromDirectory($LogsDir, $ZipPath)
    if ($json) {
        @(New-TestResultLogsZip $WntdTest.Name $WntdTest.Id $WntdResult.Status.ToString() $ZipPath) | ConvertTo-Json -Compress
    }
}
#
# CreateProjectPackage
function createprojectpackage {
    [CmdletBinding()]
    param([Switch]$help, [Switch]$rph, [Parameter(Position=1)][String]$project, [Parameter(Position=2)][String]$package)

    function Usage {
        Write-Output "createprojectpackage:"
        Write-Output ""
        Write-Output "A script that creates a project's package and saves it to a file at <package> if used,"
        Write-Output "if not to %TEMP%\prometheus_packages\..."
        Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
        Write-Output ""
        Write-Output "Usage:"
        Write-Output ""
        Write-Output "createprojectpackage <projectname> [<package>] [-help]"
        Write-Output ""
        Write-Output "Any parameter in [] is optional."
        Write-Output ""
        Write-Output "        help = Shows this message."
        Write-Output ""
        Write-Output " projectname = The name of the project."
        Write-Output ""
        Write-Output "     package = The path to the output package file."
        Write-Output ""
        Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
    }

    if ($help) {
        if (-Not $json) { Usage; return } else { throw "Help requested, ignoring..." }
    }

    if ([String]::IsNullOrEmpty($project)) {
        if (-Not $json) {
            Write-Output "WARNING: Please provide a project's name."
            Usage; return
        } else {
            throw "Please provide a project's name."
        }
    }

    if (-Not ($Manager.GetProjectNames().Contains($project))) { throw "No project with the given name was found, aborting..." } else { $WntdProject = $Manager.GetProject($project) }

    [Int]$global:Steps = 1

    [Action[Microsoft.Windows.Kits.Hardware.ObjectModel.Submission.PackageProgressInfo]]$action = {
        param([Microsoft.Windows.Kits.Hardware.ObjectModel.Submission.PackageProgressInfo]$progressinfo)

        if (($progressinfo.Current -eq 0) -and ($progressinfo.Maximum -eq 0)) {
            $jsonprogressinfo = @(New-PackageProgressInfo $progressinfo.Current $progressinfo.Maximum $progressinfo.Message) | ConvertTo-Json -Compress
            Write-Host $jsonprogressinfo
        } else {
            if ($global:Steps -lt $progressinfo.Current) {
                Write-Host -NoNewline "toolsHCK@$($ControllerName):createprojectpackage($project)> "
                [Int]$global:Steps = Read-Host
            }
            $jsonprogressinfo = @(New-PackageProgressInfo $progressinfo.Current $progressinfo.Maximum $progressinfo.Message) | ConvertTo-Json -Compress
            Write-Host $jsonprogressinfo
        }
    }

    if (-Not [String]::IsNullOrEmpty($package)) {
        $PackagePath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($package)
    } else {
        if (-Not (Test-Path ($env:TEMP + "\prometheus_packages\"))) { New-Item ($env:TEMP + "\prometheus_packages\") -ItemType Directory | Out-Null }
        $PackagePath = $env:TEMP + "\prometheus_packages\" + $(get-date).ToString("dd-MM-yyyy") + "_" + $(get-date).ToString("hh_mm_ss") + "_" + $WntdProject.Name + "." + $Studio + "x"
    }
    $PackageWriter = New-Object Microsoft.Windows.Kits.Hardware.ObjectModel.Submission.PackageWriter $WntdProject
    if ($rph) { $PackageWriter.SetProgressActionHandler($action) }
    $PackageWriter.Save($PackagePath)
    $PackageWriter.Dispose()
    if (-Not $json) {
        Write-Output "Packaged to $($PackagePath)..."
    } else {
        @(New-ProjectPackage $WntdProject.Name $PackagePath) | ConvertTo-Json -Compress
    }
}
#
# Usage
function Usage {
    Write-Output "A shell-like tool set for HCK\HLK with various purposes which covers several actions as"
    Write-Output "explained in the usage section below."
    Write-Output "These tasks are done by using the HCK\HLK API provided with the Windows HCK\HLK Studio."
    Write-Output ""
    Write-Output "Usage:"
    Write-Output ""
    Write-Output "Command: <action> <actionsparameters> [json]"
    Write-Output ""
    Write-Output "Any parameter in [] is optional."
    Write-Output ""
    Write-Output "              json = Output in JSON format."
    Write-Output ""
    Write-Output "            action = The action you want to execute."
    Write-Output ""
    Write-Output " actionsparameters = The action's parameters as explained in the action's usage."
    Write-Output "                     NOTE: use -help to show action's usage."
    Write-Output ""
    Write-Output "Actions list:"
    Write-Output ""
    Write-Output "                   help : Shows the help message."
    Write-Output ""
    Write-Output "              listpools : Lists the pools info."
    Write-Output ""
    Write-Output "             createpool : Creates a pool."
    Write-Output ""
    Write-Output "             deletepool : Deletes a pool."
    Write-Output ""
    Write-Output "            movemachine : Moves a machine from one pool to another."
    Write-Output ""
    Write-Output "        setmachinestate : Sets the state of a machine to Ready or NotReady."
    Write-Output ""
    Write-Output "          deletemachine : Deletes a machine"
    Write-Output ""
    Write-Output "     listmachinetargets : Lists the target devices of a machine that are available to be tested."
    Write-Output ""
    Write-Output "           listprojects : Lists the projects info."
    Write-Output ""
    Write-Output "          createproject : Creates a project."
    Write-Output ""
    Write-Output "          deleteproject : Deletes a project."
    Write-Output ""
    Write-Output "    createprojecttarget : Creates a project's target."
    Write-Output ""
    Write-Output "    deleteprojecttarget : Delete a project's target."
    Write-Output ""
    Write-Output "              listtests : Lists a project target's tests."
    Write-Output ""
    Write-Output "            gettestinfo : Gets a project target's test info."
    Write-Output ""
    Write-Output "              queuetest : Queue's a test, use listtestresults to get the results."
    Write-Output ""
    Write-Output "    applyprojectfilters : Applies the filters on a project's test results."
    Write-Output ""
    Write-Output " applytestresultfilters : Applies the filters on a test result."
    Write-Output ""
    Write-Output "        listtestresults : Lists a test's results info."
    Write-Output ""
    Write-Output "      ziptestresultlogs : Zips a test result's logs."
    Write-Output ""
    Write-Output "   createprojectpackage : Creates a project's package."
    Write-Output ""
    Write-Output "NOTE: For more infromation about every action use action's -help parameter!"
    Write-Output "NOTE: Windows HCK\HLK Studio should be installed on the machine running the script!"
}

# ----------------------------------------------------------------- #
# Choosing which action to perform by parsing the called parameters #
# ----------------------------------------------------------------- #
$toolsHCKlist = New-Object System.Collections.ArrayList
$toolsHCKlist.AddRange( ("listpools",
                         "createpool",
                         "deletepool",
                         "movemachine",
                         "setmachinestate",
                         "deletemachine",
                         "listmachinetargets",
                         "listprojects",
                         "createproject",
                         "deleteproject",
                         "createprojecttarget",
                         "deleteprojecttarget",
                         "listtests",
                         "gettestinfo",
                         "queuetest",
                         "applyprojectfilters",
                         "applytestresultfilters",
                         "listtestresults",
                         "ziptestresultlogs",
                         "createprojectpackage") )

# -------------------------------------- #
# Trying to perform the requested action #
# -------------------------------------- #
$ConnectFileName = $env:WTTSTDIO + "connect.xml"
Write-Output "Opening connection file $ConnectFileName"
$ConnectFile = [xml](Get-Content $ConnectFileName)

$ControllerName = $ConnectFile.Connection.GetAttribute("Server")
$DatabaseName = $connectFile.Connection.GetAttribute("Source")

Write-Output "Connecting to $ControllerName..."
$Manager = New-Object Microsoft.Windows.Kits.Hardware.ObjectModel.DBConnection.DatabaseProjectManager $ControllerName, $DatabaseName

$RootPool = $Manager.GetRootMachinePool()
$DefaultPool = $RootPool.DefaultPool

if ($server) {
    Write-Output "Initializing input server's named pipe at \\.\pipe\toolsHCKIn"
    Write-Output "Waiting for input client connection..."
    $toolsHCKInPipe = New-Object System.IO.Pipes.NamedPipeServerStream "\\.\pipe\toolsHCKIn"
    $toolsHCKInPipe.WaitForConnection()
    Write-Output "Input client is connected"

    Write-Output "Connecting to output server's named pipe at \\.\pipe\toolsHCKOut"
    $toolsHCKOutPipe = New-Object System.IO.Pipes.NamedPipeClientStream "\\.\pipe\toolsHCKOut"
    while (-Not $toolsHCKOutPipe.IsConnected) {
        try {
            $toolsHCKOutPipe.Connect(500)
        } catch [TimeoutException] {
            # NOTHING
        }
    }
    Write-Output "Connected to output server"

    $toolsHCKOutPipeSw = New-Object System.IO.StreamWriter $toolsHCKOutPipe

    $toolsHCKInPipeSr = New-Object System.IO.StreamReader $toolsHCKInPipe

    $toolsHCKOutPipeSw.WriteLine($ControllerName)
    $toolsHCKOutPipeSw.Flush()
}

while($true) {
    if ($server) {
        $cmdline = $toolsHCKInPipeSr.ReadLine()
    } else {
        Write-Host -NoNewline "toolsHCK@$ControllerName> "
        $cmdline = Read-Host
    }

    [System.Collections.ArrayList]$cmdlinelist = $cmdline.Split(" ")
    $json = $false
    if ($cmdlinelist.Contains("json")) {
        $json = $true
        $cmdlinelist.Remove("json")
    }

    $cmd = $cmdlinelist[0]
    $cmdlinelist.RemoveAt(0)
    $cmdargs = $cmdlinelist -join " "

    if ([String]::IsNullOrEmpty($cmd) -or $cmd -eq "help") {
        $output = Usage
    } elseif ($cmd -eq "exit") {
        break
    } elseif ($cmd -eq "ping") {
        $output = "pong"
    } elseif ($toolsHCKlist.Contains($cmd)) {
        try {
            $actionoutput = Invoke-Expression "$cmd $cmdargs"
            if (-Not $json) {
                $output = $actionoutput
            } else {
                $output = @(New-ActionResult $actionoutput) | ConvertTo-Json -Depth $MaxJsonDepth -Compress
            }
        } catch {
            if (-Not $json) {
                if ([String]::IsNullOrEmpty($_.Exception.InnerException)) {
                    $output = "WARNING: $($_.Exception.Message)"
                } else {
                    $output = "WARNING: $($_.Exception.InnerException.Message)"
                }
            } else {
                $output = New-ActionResult $nil $_.Exception | ConvertTo-Json -Compress
            }
        }
    } else {
        $output = "No such action name, type help."
    }

    if ($server) {
        $toolsHCKOutPipeSw.WriteLine()
        $output | foreach {
            $toolsHCKOutPipeSw.WriteLine($_)
        }
        $toolsHCKOutPipeSw.Flush()
    } else {
        $output | foreach {
            Write-Host $_
        }
    }
}

if ($server) {
    $toolsHCKOutPipeSw.Dispose()
    $toolsHCKOutPipe.Dispose()
    $toolsHCKInPipeSr.Dispose()
    $toolsHCKInPipe.Dispose()
}
