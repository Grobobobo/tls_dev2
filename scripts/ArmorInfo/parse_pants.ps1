[xml]$xmlContent = Get-Content "./base_files/ItemDefinitions_Pants"

# Load ItemListDefinitions_ArmorItems to determine base vs special tags
[xml]$itemListContent = Get-Content "./base_files/ItemListDefinitions_ArmorItems"
$itemTags = @{}

# Process all ItemsListDefinition elements to find Base and Special items
foreach ($itemListDef in $itemListContent.ItemsListDefinitions.ItemsListDefinition) {
    $listId = $itemListDef.Id
    # Check if this is a "Base" list (ends with "Base")
    if ($listId -match "Base$") {
        foreach ($item in $itemListDef.Item) {
            $itemTags[$item.Id] = "Base"
        }
    }
    # Check if this is a "Special" list (ends with "Special")
    elseif ($listId -match "Special$") {
        foreach ($item in $itemListDef.Item) {
            $itemTags[$item.Id] = "Special"
        }
    }
}

# Load Loc_TLS file and create a hashtable for item names
$locData = @{}
$locLines = Get-Content "./base_files/Loc_TLS"
foreach ($line in $locLines) {
    $parts = $line -split ','
    if ($parts[0] -match "^ItemName_") {
        $itemId = $parts[0] -replace "ItemName_", ""
        $englishName = $parts[1]
        $locData[$itemId] = $englishName
    }
}

# First pass: find the maximum number of attributes across all items
$maxAttributes = 0
foreach ($itemDef in $xmlContent.ItemDefinitions.ItemDefinition) {
    for ($i = 0; $i -lt 6; $i++) {
        $levelNode = $itemDef.LevelVariations.Level | Where-Object { $_.Id -eq $i.ToString() }
        if ($levelNode -and $levelNode.BaseStatBonuses -and $levelNode.BaseStatBonuses.BaseStatBonus) {
            $count = @($levelNode.BaseStatBonuses.BaseStatBonus).Count
            if ($count -gt $maxAttributes) {
                $maxAttributes = $count
            }
        }
    }
}

$csvData = @()

foreach ($itemDef in $xmlContent.ItemDefinitions.ItemDefinition) {
    $itemId = $itemDef.Id
    $itemName = $locData[$itemId]
    
    # Get skill for level 0
    $levelZero = $itemDef.LevelVariations.Level | Where-Object { $_.Id -eq "0" }
    $skill0 = ""
    if ($levelZero -and $levelZero.Skills) {
        $skill0 = $levelZero.Skills.Skill
    }
    
    # Get MainStatBonus values for levels 0-5
    $mainStatBonusValues = @()
    $mainStatName = ""
    for ($i = 0; $i -lt 6; $i++) {
        $levelNode = $itemDef.LevelVariations.Level | Where-Object { $_.Id -eq $i.ToString() }
        if ($levelNode) {
            if ($mainStatName -eq "") {
                $mainStatName = $levelNode.MainStatBonus.Stat
            }
            $mainStatBonusValues += $levelNode.MainStatBonus.InnerText
        }
    }
    $mainStatBonusString = $mainStatBonusValues -join "/"
    
    # Get BasePrice values for levels 0-5
    $basePriceValues = @()
    for ($i = 0; $i -lt 6; $i++) {
        $levelNode = $itemDef.LevelVariations.Level | Where-Object { $_.Id -eq $i.ToString() }
        if ($levelNode) {
            $basePriceValues += $levelNode.BasePrice
        }
    }
    $basePriceString = $basePriceValues -join "/"
    
    # Create object with base properties
    $tag = $itemTags[$itemId]
    if (-not $tag) {
        $tag = ""
    }
    
    $row = [PSCustomObject]@{
        "Tag" = $tag
        "Name" = $itemName
        "Level0Skill" = $skill0
        "MainStatBonus_Name" = $mainStatName
        "MainStatBonusLevels0-5" = $mainStatBonusString
    }
    
    # Add attribute columns (separated into Name and Values)
    for ($attrIndex = 0; $attrIndex -lt $maxAttributes; $attrIndex++) {
        $attrValues = @()
        $attrName = ""
        for ($i = 0; $i -lt 6; $i++) {
            $levelNode = $itemDef.LevelVariations.Level | Where-Object { $_.Id -eq $i.ToString() }
            $attrValue = ""
            if ($levelNode -and $levelNode.BaseStatBonuses -and $levelNode.BaseStatBonuses.BaseStatBonus) {
                $bonuses = @($levelNode.BaseStatBonuses.BaseStatBonus)
                if ($attrIndex -lt $bonuses.Count) {
                    if ($attrName -eq "") {
                        $attrName = $bonuses[$attrIndex].Stat
                    }
                    $attrValue = $bonuses[$attrIndex].InnerText
                }
            }
            $attrValues += $attrValue
        }
        # Only join non-empty values with /
        $columnNameValue = ""
        $columnValuesValue = ""
        if ($attrName -ne "") {
            $nonEmptyValues = @($attrValues | Where-Object { $_ -ne "" })
            if ($nonEmptyValues.Count -gt 0) {
                $columnNameValue = $attrName
                $columnValuesValue = $nonEmptyValues -join "/"
            }
        }
        $columnName = "Attribute$($attrIndex + 1)_Name"
        $columnValues = "Attribute$($attrIndex + 1)_Values"
        $row | Add-Member -NotePropertyName $columnName -NotePropertyValue $columnNameValue
        $row | Add-Member -NotePropertyName $columnValues -NotePropertyValue $columnValuesValue
    }
    
    # Add BasePrice as the last column
    $row | Add-Member -NotePropertyName "BasePrice" -NotePropertyValue $basePriceString
    
    $csvData += $row
}

# Export to CSV
$csvData | Export-Csv -Path "./ItemDefinitions_Pants.csv" -NoTypeInformation

Write-Host "CSV file created successfully at ./ItemDefinitions_Pants.csv"
