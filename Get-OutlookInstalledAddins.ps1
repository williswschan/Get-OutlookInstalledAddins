#Requires -Version 5.1
<#
.SYNOPSIS
    Lists Outlook COM add-ins and their status.

.DESCRIPTION
    Reads the COM add-ins that classic Outlook shows on the Add-ins page.
    Each row is the add-in name and a status of Active, Inactive, or Disabled.

.EXAMPLE
    .\Get-OutlookInstalledAddins.ps1
#>

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

Function Get-ReadableStringsFromBytes {
    Param(
        [Byte[]]$Bytes
    )

    $Results = New-Object System.Collections.Generic.List[String]
    If ($Null -Eq $Bytes -Or $Bytes.Length -Eq 0) {
        Return
    }

    $Decoded = @(
        [System.Text.Encoding]::Unicode.GetString($Bytes),
        [System.Text.Encoding]::ASCII.GetString($Bytes)
    )

    ForEach ($Text in $Decoded) {
        $Found = [Regex]::Matches($Text, "[A-Za-z0-9][A-Za-z0-9 \._\-\+\:\\]{3,}")
        ForEach ($Match in $Found) {
            $Value = $Match.Value.Trim()
            If ($Value.Length -Ge 4 -And -Not $Results.Contains($Value)) {
                $Results.Add($Value) | Out-Null
            }
        }
    }

    ForEach ($Value in $Results) {
        Write-Output $Value
    }
}

Function Get-ComServerPath {
    Param(
        [String]$ProgId
    )

    If ([String]::IsNullOrWhiteSpace($ProgId)) {
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

        Try {
            $Clsid = (Get-Item -LiteralPath $ClsidKey -ErrorAction Stop).GetValue([String]::Empty)
        }
        Catch {
            Continue
        }

        If ([String]::IsNullOrWhiteSpace([String]$Clsid)) {
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

            Try {
                $ServerPath = (Get-Item -LiteralPath $ServerKey -ErrorAction Stop).GetValue([String]::Empty)
            }
            Catch {
                Continue
            }

            If (-Not [String]::IsNullOrWhiteSpace([String]$ServerPath)) {
                Return [String]$ServerPath
            }
        }
    }
}

Function Get-AddinFileLeaf {
    Param(
        [String]$Path
    )

    If ([String]::IsNullOrWhiteSpace($Path)) {
        Return
    }

    $Clean = $Path.Trim()
    If ($Clean -Like "file:*") {
        $Clean = $Clean -Replace '^file:[/\\]*', ''
        $Clean = ($Clean -Split '\|')[0]
        $Clean = $Clean -Replace '/', '\'
    }

    $Leaf = [System.IO.Path]::GetFileName($Clean)
    If ([String]::IsNullOrWhiteSpace($Leaf)) {
        Return
    }

    Return $Leaf
}

Function Get-DisabledAddinMarkers {
    $DisabledPath = "HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems"
    If (-Not (Test-Path -LiteralPath $DisabledPath)) {
        Return
    }

    Try {
        $DisabledKey = Get-Item -LiteralPath $DisabledPath -ErrorAction Stop
    }
    Catch {
        Write-Warning "[WARN] Could not read $DisabledPath. $($_.Exception.Message)"
        Return
    }

    ForEach ($ValueName in @($DisabledKey.GetValueNames())) {
        If ([String]::IsNullOrWhiteSpace($ValueName)) {
            Continue
        }

        If ($DisabledKey.GetValueKind($ValueName) -Ne [Microsoft.Win32.RegistryValueKind]::Binary) {
            Write-Output ([String]$DisabledKey.GetValue($ValueName))
            Continue
        }

        $Strings = @(Get-ReadableStringsFromBytes -Bytes ([Byte[]]$DisabledKey.GetValue($ValueName)))
        ForEach ($Text in $Strings) {
            Write-Output $Text
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

        Try {
            $AddinKeys = @(Get-ChildItem -LiteralPath $RegistryPath -ErrorAction Stop)
        }
        Catch {
            Write-Warning "[WARN] Could not read $RegistryPath. $($_.Exception.Message)"
            Continue
        }

        ForEach ($Key in $AddinKeys) {
            Try {
                $Properties = Get-ItemProperty -LiteralPath $Key.PSPath -ErrorAction Stop
            }
            Catch {
                Write-Warning "[WARN] Could not read $($Key.PSPath). $($_.Exception.Message)"
                Continue
            }

            $ProgId = [String]$Key.PSChildName
            $KeyId = $ProgId.ToLowerInvariant()
            $FriendlyName = [String]$Properties.FriendlyName
            $Description = [String]$Properties.Description
            $FileHint = [String]$Properties.FileName
            If ([String]::IsNullOrWhiteSpace($FileHint)) {
                $FileHint = [String]$Properties.Manifest
            }

            $HasLoadBehavior = $False
            $LoadBehaviorNumber = 0
            $ParsedLoadBehavior = 0
            If ($Null -Ne $Properties.PSObject.Properties["LoadBehavior"]) {
                If ([Int]::TryParse([String]$Properties.LoadBehavior, [Ref]$ParsedLoadBehavior)) {
                    $HasLoadBehavior = $True
                    $LoadBehaviorNumber = $ParsedLoadBehavior
                }
            }

            If (-Not $Index.ContainsKey($KeyId)) {
                $Index[$KeyId] = [PSCustomObject]@{
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
            If ([String]::IsNullOrWhiteSpace($Existing.Name) -And -Not [String]::IsNullOrWhiteSpace($FriendlyName)) {
                $Existing.Name = $FriendlyName
            }
            If ([String]::IsNullOrWhiteSpace($Existing.Description) -And -Not [String]::IsNullOrWhiteSpace($Description)) {
                $Existing.Description = $Description
            }
            If ([String]::IsNullOrWhiteSpace($Existing.FileHint) -And -Not [String]::IsNullOrWhiteSpace($FileHint)) {
                $Existing.FileHint = $FileHint
            }
            If (-Not $Existing.HasLoadBehavior -And $HasLoadBehavior) {
                $Existing.HasLoadBehavior = $True
                $Existing.LoadBehavior = $LoadBehaviorNumber
            }
        }
    }

    $DisabledMarkers = @(Get-DisabledAddinMarkers)
    $Records = New-Object System.Collections.Generic.List[Object]

    ForEach ($KeyId in @($Index.Keys)) {
        $Entry = $Index[$KeyId]
        If ([String]::IsNullOrWhiteSpace($Entry.Name) -And [String]::IsNullOrWhiteSpace($Entry.Description)) {
            Continue
        }

        $LoadBehaviorNumber = 0
        If ($Entry.HasLoadBehavior) {
            $LoadBehaviorNumber = [Int]$Entry.LoadBehavior
        }

        $Status = Get-LoadBehaviorDescription -LoadBehavior $LoadBehaviorNumber
        If ($DisabledMarkers.Count -Gt 0) {
            $FileLeaf = Get-AddinFileLeaf -Path $Entry.FileHint
            If ([String]::IsNullOrWhiteSpace($FileLeaf)) {
                $FileLeaf = Get-AddinFileLeaf -Path (Get-ComServerPath -ProgId $Entry.ProgId)
            }

            ForEach ($Marker in $DisabledMarkers) {
                $MatchesProgId = $Marker -Like "*$($Entry.ProgId)*"
                $MatchesFile = (-Not [String]::IsNullOrWhiteSpace($FileLeaf)) -And ($Marker -Like "*$FileLeaf*")
                If ($MatchesProgId -Or $MatchesFile) {
                    $Status = "Disabled"
                    Break
                }
            }
        }

        $Name = [String]$Entry.Name
        If ([String]::IsNullOrWhiteSpace($Name)) {
            $Name = $Entry.ProgId
        }

        $Records.Add([PSCustomObject]@{
            Name   = $Name
            Status = $Status
        }) | Out-Null
    }

    $StatusRank = @{
        Active   = 0
        Inactive = 1
        Disabled = 2
    }
    $Sorted = @($Records | Sort-Object @{ Expression = { $StatusRank[$_.Status] } }, Name)
    ForEach ($Record in $Sorted) {
        Write-Output $Record
    }
}

$ComAddins = @(Get-OutlookComAddins)
$ComAddins | Format-Table -Property Name, Status -AutoSize
