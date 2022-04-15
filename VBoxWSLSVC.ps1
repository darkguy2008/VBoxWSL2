$VMNAME = "vmname"
$USER = "user"
$PASS = "pass"
$VBOXMANAGE = "C:\Program Files\Oracle\VirtualBox\VBoxManage.exe"

$restartCounter = 0
while($true)
{
    ##############################################################################################################################
    # Get existing rules
    ##############################################################################################################################
    $existingRules = @()
    $existingRulesOutput = & $VBOXMANAGE showvminfo $VMNAME
    $existingRulesMatches = [regex]::Matches($existingRulesOutput,'NIC \d Rule\(\d*?\):.*?\[(.*?)-(.*?)\]');
    foreach($ruleMatch in $existingRulesMatches)
    {
        $port = $ruleMatch.Groups[2].Value
        $proto = $ruleMatch.Groups[1].Value
        $existingRules += [PSCustomObject]@{ "Port"=$port; "Protocol"=$proto }
    }

    ##############################################################################################################################
    # Get currently open ports
    ##############################################################################################################################
    $newPorts = @()
    $netstatOutput = & $VBOXMANAGE guestcontrol $VMNAME run --username $USER --password $PASS "/usr/bin/sudo" -- "/usr/bin/netstat" "-ant"
    $portMatches = [regex]::matches($netstatOutput,'([a-z]{1,}?)\s*\d*\s*\d*\s*\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}\:(\d{1,5})\s*0\.0\.0\.0');
    foreach ($portMatch in $portMatches)
    {
        $port = $portMatch.Groups[2].Value;
        $proto =  $portMatch.Groups[1].Value;
        if ($port -eq 53) { continue; }
        if ($port -eq 41339) { continue; }
        $newPorts += [PSCustomObject]@{ "Port"=$port; "Protocol"=$proto }
    }

    ##############################################################################################################################
    # Prepare new rules
    ##############################################################################################################################
    $newRules = @{}
    $currentRules = @{}

    foreach($rule in $newPorts) {
        $port = $rule.Port
        $proto = $rule.Protocol
        $newRules["$proto-$port"] = $true
    }

    foreach($rule in $existingRules) {
        $port = $rule.Port
        $proto = $rule.Protocol
        $currentRules["$proto-$port"] = $true
    }

    ##############################################################################################################################
    # Remove unused ports
    ##############################################################################################################################
    foreach($key in $currentRules.Keys)
    {
        $rule = $newRules[$key]
        $split = $key.Split('-')
        $proto = $split[0]
        $port = $split[1]
        if($rule -eq $null) {
            & $VBOXMANAGE controlvm $VMNAME natpf1 delete "[$key]"
            Write-Host  "Port $port ($proto) closed" -ForegroundColor Red
        }
    }

    ##############################################################################################################################
    # Add new ports
    ##############################################################################################################################
    foreach($key in $newRules.Keys)
    {
        $split = $key.Split('-')
        $proto = $split[0]
        $port = $split[1]
        $rule = $currentRules[$key]
        if($rule -eq $null) {
            & $VBOXMANAGE controlvm $VMNAME natpf1 "[$key],$proto,,${port},,${port}"
            Write-Host "Port $port ($proto) opened" -ForegroundColor Green
        }
    }

    Start-Sleep -s 5
    $restartCounter = $restartCounter + 1
    if ($restartCounter -eq 6 || $restartCounter -gt 6) {
        $restartCounter = 0
        & $VBOXMANAGE guestcontrol $VMNAME run --username $USER --password $PASS --wait-stdout --wait-stderr --ignore-operhaned-processes --timeout 10000 "/usr/bin/sudo" -- "/usr/sbin/service" "vboxadd-service" "restart"
        Start-Sleep -s 1
        & $VBOXMANAGE guestcontrol $VMNAME run --username $USER --password $PASS --wait-stdout --wait-stderr --ignore-operhaned-processes --timeout 10000 "/usr/bin/sudo" -- "/usr/sbin/service" "vboxadd" "restart"
    }
}
