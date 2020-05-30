$mas = Get-ADComputer -Filter {enabled -eq "true"} | Sort-Object name | Select-Object name 
Write-host "Found computers - " $mas.Count -ForegroundColor Cyan

$comp_off = @()
$not_installed = @()
$not_wmi = @()
$installed_zabbix = @()
$change = @()

$path_src = "\\dfs01\distrib$\zabbix"
$path_dst = "c$\ProgramData\zabbix"
$algoritm = "MD5"

Function get_architecture ($comp){
    $architecture = (Get-WmiObject Win32_ComputerSystem -computer $comp).systemtype
    if ($architecture -eq "x64-based PC"){
        $arch = "win64"
        }
    elseif ($architecture -eq "x86-based PC") {
        $arch = "win32"
        }
    return $arch
}

Function get_hash_src($comp, $file) {
    $arch = get_architecture -comp $comp 
    switch ($file) {
        "zabbix_agentd.win.conf" {
        $file_src = Get-FileHash -Path "$path_src\win64\zabbix_agentd.win.conf" -Algorithm $algoritm
        return $file_src
        break
        }
        "smartctl-disks-discovery.ps1" {
        $file_src = Get-FileHash -Path "$path_src\win64\smartctl-disks-discovery.ps1" -Algorithm $algoritm
        return $file_src
        break
        }
        "zabbix_agentd.exe" {
        $file_src = Get-FileHash -Path "$path_src\$arch\zabbix_agentd.exe" -Algorithm $algoritm
        return $file_src
        break
        }
        "zabbix_get.exe" {
        $file_src = Get-FileHash -Path "$path_src\$arch\zabbix_get.exe" -Algorithm $algoritm
        return $file_src
        break
        }
        "zabbix_sender.exe" {
        $file_src = Get-FileHash -Path "$path_src\$arch\zabbix_sender.exe" -Algorithm $algoritm
        return $file_src
        break
        }
        default {
        $error_msges = "Missing required `"file`" parameter for `"get_hash_src`" function"
        throw $error_msges
        }
    }
}

Function get_hash_dst($comp, $file) {
    switch ($file) {
        "zabbix_agentd.win.conf" {
        $file_dest = Get-FileHash -Path "\\$comp\$path_dst\zabbix_agentd.win.conf" -Algorithm $algoritm
        return $file_dest
        break
        }
        "smartctl-disks-discovery.ps1" {
        $file_dest = Get-FileHash -Path "\\$comp\$path_dst\smartctl-disks-discovery.ps1" -Algorithm $algoritm
        return $file_dest
        break
        }
        "zabbix_agentd.exe" {
        $file_dest = Get-FileHash -Path "\\$comp\$path_dst\zabbix_agentd.exe" -Algorithm $algoritm
        return $file_dest
        break
        }
        "zabbix_get.exe" {
        $file_dest = Get-FileHash -Path "\\$comp\$path_dst\zabbix_get.exe" -Algorithm $algoritm
        return $file_dest
        break
        }
        "zabbix_sender.exe" {
        $file_dest = Get-FileHash -Path "\\$comp\$path_dst\zabbix_sender.exe" -Algorithm $algoritm
        return $file_dest
        break
        }
        default {
        $error_msg = "Missing required `"file`" parameter for `"get_hash_dst`" function"
        throw $error_msg
        }
    }
}

Function Zabbix-service($zabbix, $comp, [switch]$start, [switch]$stop) {
    if($start){
        $zabbix = connect_wmi -comp $comp
        if($zabbix){
            if ($zabbix.State -eq "Stopped"){
                $zabbix.StartService()
                Write-Host "$comp - Starting the Zabbix Agent Service"
                }
            elseif ($zabbix.State -eq "Running") {
                Write-Host "$comp - Zabbix Agent service is already running"
                }
            }
        else {
            Write-Host "$comp - Service `"Zabbix Agent`" is missing" -ForegroundColor red
        }
        }
    if($stop){
        if ($zabbix.State -eq "Running"){
            $zabbix.StopService()
            Write-Host "$comp - Stopping the Zabbix Agent Service"
            }
        else {
            Write-Host "$comp - Zabbix Agent service is not running"
            }   
        }
}

Function connect_wmi ($comp){
    $WMI = Get-WmiObject Win32_Service -Filter "Name = 'Zabbix Agent'" -ComputerName $comp
    return $WMI
}

Function compare-file ($file_use, $comp) {
    $arch = get_architecture -comp $comp 
    If(!(test-path "\\$comp\$path_dst\$file_use")){
        Copy-Item -Path $path_src\$arch\$file_use -Destination \\$comp\$path_dst\$file_use
        Write-Host "$comp - file `"$file_use`" not exist ... Copy file" -ForegroundColor Yellow
    }
    else {
        $file_src = get_hash_src -comp $comp -file $file_use
        $file_dest = get_hash_dst -comp $comp -file $file_use
        if ($file_src.Hash -ne $file_dest.Hash){
            Zabbix-service -zabbix $zabbix -comp $comp -stop
            Copy-Item -Path $file_src.Path -Destination $file_dest.Path -Force        
            Write-Host "$comp - File updated: $file_dest" -ForegroundColor Yellow
            $change += $comp           
        }
    }     
}
Function smartmontools ($comp){
    if (!(Test-Path -Path "\\$comp\c$\Program Files\smartmontools")){
        Write-Host "$comp - Copy smartmontools" -ForegroundColor Yellow
        $arch = get_architecture -comp $comp
        Copy-Item -Path  "$path_src\smartmontools\$arch\smartmontools" -Destination "\\$comp\c$\Program Files\smartmontools" -Recurse -Force
        if ((Test-Path -Path "\\$comp\c$\Program Files\smartmontools\bin\smartctl.exe")){
            Write-Host "$comp - smartmontools copied to C:\Program Files\smartmontools" -ForegroundColor Yellow
        }
        else {
            Write-Host "$comp - copying smartmontools failed" -ForegroundColor red
        }
    }
    else {
        Write-Host "$comp - smartmontools is already installed"
    }
}

Function result ($array, $msg, $color, [switch]$total) {
    if ($array){
        Write-host ""
        Write-host $msg -ForegroundColor $color
        $array
        if ($total){
            Write-host "Total - " $array.Count -ForegroundColor Cyan
        }
    }
}

foreach ($comp in $mas.Name){
    if (Test-connection $comp -Count 2 -quiet){
        try
            {
            $zabbix = connect_wmi -comp $comp
            if ($zabbix -and (test-path "\\$comp\$path_dst")){
                Write-Host "$comp - OK"
                $installed_zabbix += $comp

                # Zabbix_agentd.win.conf
                compare-file -file_use "Zabbix_agentd.win.conf" -comp $comp
                 
                # smartctl-disks-discovery.ps1
                compare-file -file_use "smartctl-disks-discovery.ps1" -comp $comp

                # zabbix_agentd.exe
                compare-file -file_use "zabbix_agentd.exe" -comp $comp

                # zabbix_get.exe
                compare-file -file_use "zabbix_get.exe" -comp $comp
                              
                # zabbix_sender.exe
                compare-file -file_use "zabbix_sender.exe" -comp $comp
                 
                # smartmontools
                smartmontools -comp $comp
                    
                Zabbix-service -zabbix $zabbix -comp $comp -start
            }
            else {
                Write-Host "$comp - Zabbix is not installed!" -ForegroundColor Magenta
                

                # smartmontools
                smartmontools -comp $comp

                # zabbix dir
                if(!(test-path "\\$comp\$path_dst")){
                    New-Item -ItemType Directory -Force -Path "\\$comp\$path_dst"
                    }
                $arch = get_architecture -comp $comp
                Copy-Item -Path "$path_src\$arch\*" -Destination "\\$comp\$path_dst" -Recurse -Force
                Write-Host "Copying the zabbix directory" -ForegroundColor Magenta
                Start-Sleep -s 1

                # Install Zabbix Agent service 
                Invoke-Command -ErrorVariable error_var -ErrorAction SilentlyContinue -ComputerName $comp -ScriptBlock { 
                    C:\ProgramData\zabbix\zabbix_agentd.exe --config C:\ProgramData\zabbix\zabbix_agentd.win.conf --install
                }
                $error_var 
                Write-Host "$comp - Install zabbix" -ForegroundColor Magenta
                Zabbix-service -zabbix $zabbix -comp $comp -start  
                $not_installed += $comp                      
            }
        }
        catch [System.UnauthorizedAccessException]
        {
            Write-Host "$comp - Error! - Access denied" -ForegroundColor red
            #Write-Host $_
            $not_wmi += $comp
        }
        catch 
        {            
            Write-Host "$comp - Warning error!" -ForegroundColor red
            Write-Host $_.Exception -ForegroundColor red
            $not_wmi += $comp
        }      
    }
    else {
        Write-host "$comp - Not available (Ping)" -ForegroundColor red
        $comp_off += $comp
    }
}
    
result -array $comp_off -msg "List of not available computers:" -color red
result -array $not_wmi -msg "No WMI available" -color red
result -array $not_installed -msg "Installed service `"Zabbix Agent`":" -color Magenta -total
result -array $installed_zabbix -msg "List of computers with installed and running service `"Zabbix Agent`":" -color green -total
result -array $change -msg "Updated files:" -color Yellow -total
