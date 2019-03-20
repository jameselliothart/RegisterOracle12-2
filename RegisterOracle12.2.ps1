function Test-OraclePath {
    Param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [string]$Path
    )

    process {
        while ([string]::IsNullOrEmpty($Path) -or !(Test-Path $Path)) {
            $Path = Read-Host "Could not find Oracle client bin directory path '$Path'. Please enter here"
        }

        Write-Host "'$Path' is valid" -ForegroundColor Green
        return $Path
    }
}

Push-Location $PSScriptRoot

$oraclePaths = $env:path.split(';') | Where-Object {$_ -like '*oracle*client*bin'}
Write-Host "Found the following Oracle client directories:"
$oraclePaths | Write-Host

Write-Host "Attempting to determine 32bit client bin directory..."
$32BitInstallPath = $oraclePaths | Where-Object {$_ -like '*32*'} | Select-Object -First 1 | Test-OraclePath

Write-Host "Registering 32bit OraOLEDB12.dll..."
Set-Location c:\windows\system32
.\regsvr32.exe "$32BitInstallPath\OraOLEDB12.dll"

Write-Host "Attempting to determine 64bit client bin directory..."
$64BitInstallPath = $oraclePaths | Where-Object {$_ -like '*64*'} | Select-Object -First 1 | Test-OraclePath

Write-Host "Registering 64bit OraOLEDB12.dll..."
Set-Location c:\windows\syswow64
.\regsvr32.exe "$64BitInstallPath\OraOLEDB12.dll"

foreach ($path in $32BitInstallPath, $64BitInstallPath) {
    Write-Host "Attempting to register Oracle client via OraProvCfg.exe in ..\ODP.NET\bin\4 directory..."
    $odpBin = "$path\..\ODP.NET\bin\4" | Test-OraclePath
    Set-Location $odpBin
    Write-Host "Running: OraProvCfg.exe /action:gac /providerpath:""$odpBin\Oracle.DataAccess.dll"""
    .\OraProvCfg.exe /action:gac /providerpath:"$odpBin\Oracle.DataAccess.dll"
}

# Oracle DbProviderFactories information
$unmanaged = [ordered]@{
    name        = "ODP.NET, Unmanaged Driver"
    invariant   = "Oracle.DataAccess.Client"
    description = "Oracle Data Provider for .NET, Unmanaged Driver"
    type        = "Oracle.DataAccess.Client.OracleClientFactory, Oracle.DataAccess, Version=4.122.1.0, Culture=neutral, PublicKeyToken=89b483f429c47342"
}
$managed = [ordered]@{
    name        = "ODP.NET, Managed Driver"
    invariant   = "Oracle.ManagedDataAccess.Client"
    description = "Oracle Data Provider for .NET, Managed Driver"
    type        = "Oracle.ManagedDataAccess.Client.OracleClientFactory, Oracle.ManagedDataAccess, Version=4.122.1.0, Culture=neutral, PublicKeyToken=89b483f429c47342"
}

# Get machine.config paths
$machineConfigPathTemp = [System.Runtime.InteropServices.RuntimeEnvironment]::SystemConfigurationFile
if ($machineConfigPathTemp -like "*\Framework64\*") {
    $machineConfigPath32 = $machineConfigPathTemp -replace '\\Framework64\\', '\Framework\'
    $machineConfigPath64 = $machineConfigPathTemp
}
else {
    $machineConfigPath32 = $machineConfigPathTemp
    $machineConfigPath64 = $machineConfigPathTemp -replace '\\Framework\\', '\Framework64\'
}

Write-Host "Updating machine.config files with Oracle DbProviderFactories..."
foreach ($machineConfigPath in $machineConfigPath32, $machineConfigPath64) {
    # Make a backup on the machine.config
    $datetime = Get-Date -Format yyyyMMddhhmmss
    $backupDirectory = if ($machineConfigPath -like "*\Framework64\*") {'C:\Temp\machineConfig64'} else {'C:\Temp\machineConfig32'}
    $backupPath = "$backupDirectory\machine$datetime.config"
    Write-Host "Creating backup machine.config in $backupPath"
    if (-not (Test-Path $backupDirectory)) {
        New-Item -Path $backupDirectory -ItemType Directory
    }
    Copy-Item -Path $machineConfigPath -Destination $backupPath

    [xml]$machineConfig = get-content $machineConfigPath
    $dbProviderFactories = $machineConfig.SelectSingleNode("//system.data/DbProviderFactories")
    $invariants = $dbProviderFactories.add.invariant

    foreach ($driver in $managed, $unmanaged) {
        if ($driver.invariant -notin $invariants) {
            $factory = $machineConfig.CreateElement("add")
            foreach ($attribute in $driver.GetEnumerator()) {
                $factoryAtt = $machineConfig.CreateAttribute("$($attribute.Key)")
                $factoryAtt.Value = $attribute.Value
                $factory.Attributes.Append($factoryAtt)
            }
            $dbProviderFactories.AppendChild($factory)
        }
    }

    $machineConfig.Save("$machineConfigPath")
}

Pop-Location