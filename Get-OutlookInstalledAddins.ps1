<#
.SYNOPSIS
    Lists Outlook COM add-ins and their status.

.DESCRIPTION
    Reads the COM add-ins that classic Outlook shows on the Add-ins page.
    Each row is the add-in name and a status of Active, Inactive, or Disabled.
    The script uses cmdlets, operators, and PowerShell core types so it can run
    when Constrained Language mode is turned on.

.EXAMPLE
    .\Get-OutlookInstalledAddins.ps1
#>

Function New-CoreObject {
    Param(
        [Hashtable]$Property
    )

    # [PSCustomObject] is rejected in Constrained Language mode.
    Return (New-Object -TypeName PSObject -Property $Property)
}

Function Test-Blank {
    Param(
        [String]$Value
    )

    If ($Null -Eq $Value) {
        Return $True
    }

    Return ($Value.Trim().Length -Eq 0)
}

Function Get-LoadBehaviorDescription {
    Param(
        [Int]$LoadBehavior
    )

    # The Outlook Add-ins page lists an add-in as Active when LoadBehavior is 3
    # (connected and load at startup). Load-on-demand (9) and not loaded (2)
    # appear under Inactive. Disabled is a separate resiliency state.
    If (($LoadBehavior -Band 3) -Eq 3) {
        Return "Active"
    }

    Return "Inactive"
}

Function Get-TextPiecesFromBytes {
    Param(
        $Bytes
    )

    $ByteList = @($Bytes)
    $Length = $ByteList.Count
    If ($Length -Eq 0) {
        Return
    }

    $Seen = @{}
    $Modes = @(2, 1)
    ForEach ($Step in $Modes) {
        $Builder = ""
        For ($Index = 0; $Index -Lt $Length; $Index += $Step) {
            $Code = -1
            If ($Step -Eq 2) {
                If (($Index + 1) -Ge $Length) {
                    Break
                }
                $Code = [Int]$ByteList[$Index] + (256 * [Int]$ByteList[$Index + 1])
            }
            Else {
                $Code = [Int]$ByteList[$Index]
            }

            $IsPrintable = ($Code -Ge 32 -And $Code -Le 126)
            If ($IsPrintable) {
                $Builder += [Char]$Code
                Continue
            }

            If ($Builder.Length -Ge 4 -And -Not $Seen.ContainsKey($Builder)) {
                $Seen[$Builder] = $True
                Write-Output $Builder
            }
            $Builder = ""
        }

        If ($Builder.Length -Ge 4 -And -Not $Seen.ContainsKey($Builder)) {
            $Seen[$Builder] = $True
            Write-Output $Builder
        }
    }
}

Function Get-RegistryDefaultValue {
    Param(
        [String]$Path
    )

    $Properties = Get-ItemProperty -LiteralPath $Path -ErrorAction Stop
    $DefaultValue = $Properties.'(default)'
    If (Test-Blank -Value ([String]$DefaultValue)) {
        Return
    }

    Return [String]$DefaultValue
}

Function Get-ComServerPath {
    Param(
        [String]$ProgId
    )

    If (Test-Blank -Value $ProgId) {
        Return
    }

    $ClassRoots = @(
        "HKCU:\SOFTWARE\Classes",
        "HKLM:\SOFTWARE\Classes",
        "HKLM:\SOFTWARE\WOW6432Node\Classes",
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Classes"
    )

    ForEach ($Root in $ClassRoots) {
        $ClsidKey = "$Root\$ProgId\CLSID"
        If (-Not (Test-Path -LiteralPath $ClsidKey)) {
            Continue
        }

        $Clsid = ""
        Try {
            $Clsid = Get-RegistryDefaultValue -Path $ClsidKey
        }
        Catch {
            Continue
        }

        If (Test-Blank -Value $Clsid) {
            Continue
        }

        $ServerKeys = @(
            "$Root\CLSID\$Clsid\InprocServer32",
            "$Root\CLSID\$Clsid\LocalServer32"
        )

        ForEach ($ServerKey in $ServerKeys) {
            If (-Not (Test-Path -LiteralPath $ServerKey)) {
                Continue
            }

            $ServerPath = ""
            Try {
                $ServerPath = Get-RegistryDefaultValue -Path $ServerKey
            }
            Catch {
                Continue
            }

            If (-Not (Test-Blank -Value $ServerPath)) {
                Return [String]$ServerPath
            }
        }
    }
}

Function Get-AddinFileLeaf {
    Param(
        [String]$Path
    )

    If (Test-Blank -Value $Path) {
        Return
    }

    $Clean = $Path.Trim()
    If ($Clean -Like "file:*") {
        $Clean = $Clean -Replace '^file:[/\\]*', ''
        $Clean = ($Clean -Split '\|')[0]
        $Clean = $Clean -Replace '\\', '/'
    }
    Else {
        $Clean = $Clean -Replace '\\', '/'
    }

    $Parts = @($Clean -Split '/')
    $Leaf = [String]$Parts[$Parts.Count - 1]
    If (Test-Blank -Value $Leaf) {
        Return
    }

    Return $Leaf
}

Function Get-DisabledAddinMarkers {
    $DisabledPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems"
    If (-Not (Test-Path -LiteralPath $DisabledPath)) {
        Return
    }

    $Properties = $Null
    Try {
        $Properties = Get-ItemProperty -LiteralPath $DisabledPath -ErrorAction Stop
    }
    Catch {
        Write-Warning "[WARN] Could not read $DisabledPath. $($_.Exception.Message)"
        Return
    }

    $SkipNames = @(
        "PSPath",
        "PSParentPath",
        "PSChildName",
        "PSDrive",
        "PSProvider"
    )

    ForEach ($Property in @($Properties.PSObject.Properties)) {
        $Name = [String]$Property.Name
        $Skip = $False
        ForEach ($SkipName in $SkipNames) {
            If ($Name -Eq $SkipName) {
                $Skip = $True
            }
        }
        If ($Skip) {
            Continue
        }

        $Data = $Property.Value
        If ($Data -Is [String]) {
            Write-Output ([String]$Data)
            Continue
        }

        If ($Data -Is [Byte[]]) {
            Get-TextPiecesFromBytes -Bytes $Data
        }
    }
}

Function Get-OutlookComAddins {
    # Current-user keys are read first. Later keys only fill blanks, so a
    # per-user LoadBehavior wins over the machine and Click-to-Run copies.
    $Locations = @(
        "HKCU:\SOFTWARE\Microsoft\Office\Outlook\Addins",
        "HKLM:\SOFTWARE\Microsoft\Office\Outlook\Addins",
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\REGISTRY\MACHINE\Software\Microsoft\Office\Outlook\Addins",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\Outlook\Addins"
    )

    $Index = @{}

    ForEach ($RegistryPath in $Locations) {
        If (-Not (Test-Path -LiteralPath $RegistryPath)) {
            Continue
        }

        $AddinKeys = @()
        Try {
            $AddinKeys = @(Get-ChildItem -LiteralPath $RegistryPath -ErrorAction Stop)
        }
        Catch {
            Write-Warning "[WARN] Could not read $RegistryPath. $($_.Exception.Message)"
            Continue
        }

        ForEach ($Key in $AddinKeys) {
            $Properties = $Null
            Try {
                $Properties = Get-ItemProperty -LiteralPath $Key.PSPath -ErrorAction Stop
            }
            Catch {
                Write-Warning "[WARN] Could not read $($Key.PSPath). $($_.Exception.Message)"
                Continue
            }

            $ProgId = [String]$Key.PSChildName
            $KeyId = $ProgId.ToLower()
            $FriendlyName = [String]$Properties.FriendlyName
            $Description = [String]$Properties.Description
            $FileHint = [String]$Properties.FileName
            If (Test-Blank -Value $FileHint) {
                $FileHint = [String]$Properties.Manifest
            }

            $HasLoadBehavior = $False
            $LoadBehaviorNumber = 0
            $LoadBehaviorText = [String]$Properties.LoadBehavior
            If ($LoadBehaviorText -Match '^\d+$') {
                $HasLoadBehavior = $True
                $LoadBehaviorNumber = [Int]$LoadBehaviorText
            }

            If (-Not $Index.ContainsKey($KeyId)) {
                $Index[$KeyId] = New-CoreObject -Property @{
                    ProgId          = $ProgId
                    Name            = $FriendlyName
                    Description     = $Description
                    FileHint        = $FileHint
                    LoadBehavior    = $LoadBehaviorNumber
                    HasLoadBehavior = $HasLoadBehavior
                }
                Continue
            }

            $Existing = $Index[$KeyId]
            If ((Test-Blank -Value $Existing.Name) -And -Not (Test-Blank -Value $FriendlyName)) {
                $Existing.Name = $FriendlyName
            }
            If ((Test-Blank -Value $Existing.Description) -And -Not (Test-Blank -Value $Description)) {
                $Existing.Description = $Description
            }
            If ((Test-Blank -Value $Existing.FileHint) -And -Not (Test-Blank -Value $FileHint)) {
                $Existing.FileHint = $FileHint
            }
            If ((-Not $Existing.HasLoadBehavior) -And $HasLoadBehavior) {
                $Existing.HasLoadBehavior = $True
                $Existing.LoadBehavior = $LoadBehaviorNumber
            }
        }
    }

    $DisabledMarkers = @(Get-DisabledAddinMarkers)
    $Records = @()

    ForEach ($KeyId in @($Index.Keys)) {
        $Entry = $Index[$KeyId]
        If ((Test-Blank -Value $Entry.Name) -And (Test-Blank -Value $Entry.Description)) {
            Continue
        }

        $LoadBehaviorNumber = 0
        If ($Entry.HasLoadBehavior) {
            $LoadBehaviorNumber = [Int]$Entry.LoadBehavior
        }

        $Status = Get-LoadBehaviorDescription -LoadBehavior $LoadBehaviorNumber
        If ($DisabledMarkers.Count -Gt 0) {
            $FileLeaf = Get-AddinFileLeaf -Path $Entry.FileHint
            If (Test-Blank -Value $FileLeaf) {
                $FileLeaf = Get-AddinFileLeaf -Path (Get-ComServerPath -ProgId $Entry.ProgId)
            }

            ForEach ($Marker in $DisabledMarkers) {
                $MatchesProgId = $Marker -Like "*$($Entry.ProgId)*"
                $MatchesFile = (-Not (Test-Blank -Value $FileLeaf)) -And ($Marker -Like "*$FileLeaf*")
                If ($MatchesProgId -Or $MatchesFile) {
                    $Status = "Disabled"
                    Break
                }
            }
        }

        $Name = [String]$Entry.Name
        If (Test-Blank -Value $Name) {
            $Name = $Entry.ProgId
        }

        $Records += New-CoreObject -Property @{
            Name   = $Name
            Status = $Status
        }
    }

    $StatusRank = @{
        Active   = 0
        Inactive = 1
        Disabled = 2
    }
    Return @($Records | Sort-Object @{ Expression = { $StatusRank[$_.Status] } }, Name)
}

$ComAddins = @(Get-OutlookComAddins)
$ComAddins | Format-Table -Property Name, Status -AutoSize
