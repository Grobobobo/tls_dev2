[xml]$xmlContent = Get-Content "./base_files/SkillDefinitions_Items_Other"

# Load Loc_TLS file and create a hashtable for skill names
$locData = @{}
$locLines = Get-Content "./base_files/Loc_TLS"
foreach ($line in $locLines) {
    $parts = $line -split ','
    if ($parts[0] -match "^SkillName_") {
        $skillId = $parts[0] -replace "SkillName_", ""
        $englishName = $parts[1]
        $locData[$skillId] = $englishName
    }
}

$csvData = @()

# Get all skill definitions that have Buff or RegenStat
$skillsWithEffects = @{}

foreach ($skillDef in $xmlContent.SkillDefinitions.SkillDefinition) {
    $skillId = $skillDef.Id
    $skillName = $locData[$skillId]
    if (-not $skillName) {
        $skillName = $skillId
    }
    
    # Extract base skill name (without trailing numbers)
    $baseSkillName = $skillId -replace '\d+$', ''
    
    $hasBuff = $false
    $buffData = $null
    $hasRegenStat = $false
    $regenStatData = $null
    
    if ($skillDef.SkillAction -and $skillDef.SkillAction.Generic -and $skillDef.SkillAction.Generic.SkillEffects) {
        $skillEffects = $skillDef.SkillAction.Generic.SkillEffects
        if ($skillEffects.CasterEffect) {
            if ($skillEffects.CasterEffect.Buff) {
                $hasBuff = $true
                $buffData = $skillEffects.CasterEffect.Buff
            }
            if ($skillEffects.CasterEffect.RegenStat) {
                $hasRegenStat = $true
                $regenStatData = $skillEffects.CasterEffect.RegenStat
            }
        }
    }
    
    # Process only if skill has Buff or RegenStat
    if (($hasBuff -and $buffData) -or ($hasRegenStat -and $regenStatData)) {
        # Get costs
        $actionPointsCost = if ($skillDef.ActionPointsCost) { $skillDef.ActionPointsCost } else { "" }
        $manaCost = if ($skillDef.ManaCost) { $skillDef.ManaCost } else { "" }
        $movePointsCost = if ($skillDef.MovePointsCost) { $skillDef.MovePointsCost } else { "" }
        
        if ($hasBuff -and $buffData) {
            $statId = if ($buffData.Stat) { $buffData.Stat.Id } else { "" }
            $turnsCount = if ($buffData.TurnsCount) { $buffData.TurnsCount } else { "" }
            
            # If this is not the first occurrence of this skill, aggregate the bonuses
            if ($skillsWithEffects.ContainsKey($baseSkillName)) {
                $skillsWithEffects[$baseSkillName]["Bonuses"] += $buffData.Bonus
            } else {
                $skillsWithEffects[$baseSkillName] = @{
                    "Name" = $skillName
                    "Type" = "Buff"
                    "ActionPointsCost" = $actionPointsCost
                    "ManaCost" = $manaCost
                    "MovePointsCost" = $movePointsCost
                    "StatId" = $statId
                    "Bonuses" = @($buffData.Bonus)
                    "TurnsCount" = $turnsCount
                }
            }
        } elseif ($hasRegenStat -and $regenStatData) {
            $statId = if ($regenStatData.Stat) { $regenStatData.Stat.Id } else { "" }
            $bonus = if ($regenStatData.Bonus) { $regenStatData.Bonus } else { "" }
            
            # If this is not the first occurrence of this skill, aggregate the bonuses
            if ($skillsWithEffects.ContainsKey($baseSkillName)) {
                if ($bonus -ne "") {
                    $skillsWithEffects[$baseSkillName]["Bonuses"] += $bonus
                }
            } else {
                $skillsWithEffects[$baseSkillName] = @{
                    "Name" = $skillName
                    "Type" = "RegenStat"
                    "ActionPointsCost" = $actionPointsCost
                    "ManaCost" = $manaCost
                    "MovePointsCost" = $movePointsCost
                    "StatId" = $statId
                    "Bonuses" = if ($bonus -ne "") { @($bonus) } else { @() }
                    "TurnsCount" = ""
                }
            }
        }
    }
}

# Now create CSV data from aggregated skills
foreach ($baseSkillName in $skillsWithEffects.Keys | Sort-Object) {
    $skillData = $skillsWithEffects[$baseSkillName]
    $bonusString = @($skillData["Bonuses"] | Where-Object { $_ -ne "" }) -join "/"
    
    $csvData += [PSCustomObject]@{
        "InternalName" = $baseSkillName
        "Name" = $skillData["Name"]
        "Type" = $skillData["Type"]
        "ActionPointsCost" = $skillData["ActionPointsCost"]
        "ManaCost" = $skillData["ManaCost"]
        "MovePointsCost" = $skillData["MovePointsCost"]
        "StatName" = $skillData["StatId"]
        "Bonus" = $bonusString
        "TurnsCount" = $skillData["TurnsCount"]
    }
}

# Export to CSV
$csvData | Export-Csv -Path "./SkillDefinitions_Items_Other.csv" -NoTypeInformation

Write-Host "CSV file created successfully at ./SkillDefinitions_Items_Other.csv"
