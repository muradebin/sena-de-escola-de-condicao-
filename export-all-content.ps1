[CmdletBinding()]
param(
  [string]$AssetsRoot = '.\1\assets\flutter_assets\assets',
  [string]$OutDir = '.\content_export',
  [switch]$IncludeFileInventory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Web.Extensions

function Get-CategoryFromUnitName {
  param([string]$Unit)

  if ([string]::IsNullOrWhiteSpace($Unit)) { return 'unknown' }
  if ($Unit -match '^(revisao)$') { return 'revisao' }
  if ($Unit -match '^([a-zA-Z]+)\d+') { return $Matches[1].ToLowerInvariant() }
  if ($Unit -match '^([a-zA-Z]+)') { return $Matches[1].ToLowerInvariant() }
  return 'unknown'
}

function Get-TierFromBundle {
  param([string]$Bundle)

  switch ($Bundle.ToLowerInvariant()) {
    'demo' { return 'free' }
    'pro' { return 'paid' }
    'lp' { return 'paid' }
    default { return 'unknown' }
  }
}

function Get-QuizPayloadFromIndexHtml {
  param([string]$Path)

  $text = Get-Content -Raw -LiteralPath $Path
  $m = [regex]::Match($text, '\bvar\s+data\s*=\s*([\x27\x22])(?<b64>[A-Za-z0-9+/=]+)\1')
  if (-not $m.Success) { return $null }

  $bytes = [Convert]::FromBase64String($m.Groups['b64'].Value)
  return [Text.Encoding]::UTF8.GetString($bytes)
}

function Parse-JsonDeep {
  param([string]$Json)

  $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
  $ser.MaxJsonLength = [int]::MaxValue
  $ser.RecursionLimit = 4000
  return $ser.DeserializeObject($Json)
}

function Try-Get {
  param(
    [object]$Obj,
    [string]$Key
  )

  if ($null -eq $Obj) { return $null }

  if ($Obj -is [System.Collections.IDictionary]) {
    try { return $Obj[$Key] } catch { return $null }
  }

  if ($Obj.PSObject -and $Obj.PSObject.Properties.Name -contains $Key) {
    return $Obj.$Key
  }

  try { return $Obj[$Key] } catch { }
  return $null
}

function Get-Slides {
  param([object]$Root)

  $d = Try-Get $Root 'd'
  $sl = Try-Get $d 'sl'
  $groups = Try-Get $sl 'g'

  if (-not ($groups -is [System.Collections.IEnumerable])) { return @() }
  $firstGroup = @($groups)[0]
  $slides = Try-Get $firstGroup 'S'
  if (-not ($slides -is [System.Collections.IEnumerable])) { return @() }

  return @($slides)
}

function Extract-QuestionsFromSlides {
  param(
    [object[]]$Slides
  )

  $questions = New-Object System.Collections.Generic.List[object]
  $n = 0

  foreach ($slide in $Slides) {
    $slideType = [string](Try-Get $slide 'tp')

    $d = Try-Get $slide 'D'
    $questionText = [string](Try-Get $d 'd')
    $questionHtml = [string](Try-Get $d 'h')

    $at = Try-Get $slide 'at'
    $ati = Try-Get $at 'i'
    $imageUri = [string](Try-Get $ati 'i')
    if ([string]::IsNullOrWhiteSpace($imageUri)) { $imageUri = $null }

    $imageRelativePath = $null
    if ($imageUri -and $imageUri.StartsWith('storage://images/')) {
      $imageRelativePath = 'data/images/' + $imageUri.Substring('storage://images/'.Length)
    }

    $choicesContainer = Try-Get $slide 'C'
    $choices = Try-Get $choicesContainer 'chs'

    # Consider as a "question" if it has choices; otherwise we still keep it as content with empty options.
    $options = New-Object System.Collections.Generic.List[object]
    if ($choices -is [System.Collections.IEnumerable]) {
      foreach ($ch in @($choices)) {
        $choiceTextContainer = Try-Get $ch 't'
        $optionText = [string](Try-Get $choiceTextContainer 'd')
        $optionHtml = [string](Try-Get $choiceTextContainer 'h')
        $isCorrect = $false
        $isCorrectRaw = Try-Get $ch 'c'
        if ($null -ne $isCorrectRaw) { $isCorrect = [bool]$isCorrectRaw }

        $options.Add([pscustomobject]@{
            text    = $optionText
            html    = $optionHtml
            correct = $isCorrect
          })
      }
    }

    # Only increment question numbering for slides that look like real question slides
    if ($options.Count -gt 0 -or ($slideType -match 'Choice|Response|TrueFalse|Drag|Fill|Match')) {
      $n++
    } else {
      continue
    }

    $questions.Add([pscustomobject]@{
        n                 = $n
        slideId           = [string](Try-Get $slide 'i')
        type              = $slideType
        question          = $questionText
      questionHtml      = $questionHtml
        imageUri          = $imageUri
        imageRelativePath = $imageRelativePath
        options           = $options
      })
  }

  return $questions
}

function Get-FileInventory {
  param([string]$Dir)

  $files = Get-ChildItem -LiteralPath $Dir -File -Recurse
  $totalBytes = 0
  $entries = foreach ($f in $files) {
    $totalBytes += $f.Length

    $relPath = $f.FullName.Substring($Dir.Length)
    $relPath = $relPath.TrimStart([char[]]'\\/')
    $relPath = $relPath.Replace('\\', '/')

    [pscustomobject]@{
      path = $relPath
      bytes = [int64]$f.Length
    }
  }

  return [pscustomobject]@{
    fileCount = $files.Count
    totalBytes = [int64]$totalBytes
    files = $entries
  }
}

function Ensure-Dir {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) {
    New-Item -ItemType Directory -Path $Path | Out-Null
  }
}

$assetsRootFull = (Resolve-Path -LiteralPath $AssetsRoot).Path
Ensure-Dir -Path $OutDir
$outDirFull = (Resolve-Path -LiteralPath $OutDir).Path

$indexFiles = Get-ChildItem -LiteralPath $assetsRootFull -Filter index.html -File -Recurse |
  Where-Object { $_.FullName -match '\\www\\(index\.html|[^\\]+\\index\.html)$' }

$allPackages = New-Object System.Collections.Generic.List[object]
$indexAudit = New-Object System.Collections.Generic.List[object]
$byImage = @{}
$errors = New-Object System.Collections.Generic.List[object]

foreach ($file in $indexFiles) {
  try {
    $payload = Get-QuizPayloadFromIndexHtml -Path $file.FullName
    if (-not $payload) {
      $relIndex = $file.FullName.Substring($assetsRootFull.Length)
      $relIndex = $relIndex.TrimStart([char[]]'\\/')
        $parts = $relIndex -split '[\\/]'
        $bundle = if ($parts.Length -ge 1) { $parts[0] } else { 'unknown' }
        $unit = if ($parts.Length -ge 3) {
          if ($parts[2] -eq 'index.html') { '(www-root)' } else { $parts[2] }
        } else { 'unknown' }
      $indexAudit.Add([pscustomobject]@{
          bundle = $bundle
          unit = $unit
          indexHtml = $file.FullName
          hasPayload = $false
          exported = $false
          reason = 'no base64 var data'
        })
      continue
    }

    $root = Parse-JsonDeep -Json $payload
    $slides = Get-Slides -Root $root
    if ($slides.Count -eq 0) { continue }

    # Derive bundle + unit from path: <assetsRoot>\<bundle>\www\<unit>\index.html
    $rel = $file.FullName.Substring($assetsRootFull.Length)
    $rel = $rel.TrimStart([char[]]'\\/')
    $parts = $rel -split '[\\/]'
    $bundle = $parts[0]
    $unit = if ($parts[2] -eq 'index.html') { '(www-root)' } else { $parts[2] }

    if ($unit -eq '(www-root)') {
      $indexAudit.Add([pscustomobject]@{
          bundle = $bundle
          unit = $unit
          indexHtml = $file.FullName
          hasPayload = $true
          exported = $false
          reason = 'bundle landing page (www/index.html)'
        })
      continue
    }

    $tier = Get-TierFromBundle -Bundle $bundle
    $category = Get-CategoryFromUnitName -Unit $unit

    $d = Try-Get $root 'd'
    $title = [string](Try-Get $d 'T')

    $questions = Extract-QuestionsFromSlides -Slides $slides

    $usedImages = @($questions | Where-Object { $_.imageRelativePath } | ForEach-Object { $_.imageRelativePath } | Sort-Object -Unique)

    $unitDir = Split-Path -Parent $file.FullName
    $imagesDir = Join-Path $unitDir 'data\images'
    $imageFiles = @()
    if (Test-Path -LiteralPath $imagesDir) {
      $imageFiles = Get-ChildItem -LiteralPath $imagesDir -File | Select-Object -ExpandProperty Name | Sort-Object
    }

    foreach ($q in $questions) {
      if (-not $q.imageRelativePath) { continue }
      $img = $q.imageRelativePath
      if (-not $byImage.ContainsKey($img)) { $byImage[$img] = @() }
      $byImage[$img] += [pscustomobject]@{
        bundle = $bundle
        tier = $tier
        unit = $unit
        category = $category
        questionNumber = $q.n
        question = $q.question
      }
    }

    $record = [pscustomobject]@{
      bundle = $bundle
      tier = $tier
      category = $category
      unit = $unit
      title = $title
      indexHtml = ($file.FullName)
      questionCount = $questions.Count
      usedImages = $usedImages
      availableImages = $imageFiles
      questions = $questions
    }

    if ($IncludeFileInventory) {
      $record | Add-Member -NotePropertyName fileInventory -NotePropertyValue (Get-FileInventory -Dir $unitDir)
    }

    $outUnitDir = Join-Path $OutDir (Join-Path $bundle $unit)
    Ensure-Dir -Path (Join-Path $OutDir $bundle)
    Ensure-Dir -Path $outUnitDir

    $outFile = Join-Path $outUnitDir 'content.json'
    $record | ConvertTo-Json -Depth 30 | Set-Content -LiteralPath $outFile -Encoding UTF8 -NoNewline

    $outFileFull = (Resolve-Path -LiteralPath $outFile).Path
    $outFileRel = $outFileFull.Substring($outDirFull.Length)
    $outFileRel = $outFileRel.TrimStart([char[]]@('\','/'))
    $outFileRel = $outFileRel -replace '\\','/'

    $allPackages.Add([pscustomobject]@{
        bundle = $bundle
        tier = $tier
        category = $category
        unit = $unit
        title = $title
        questionCount = $questions.Count
        outFile = $outFileRel
      })

    $indexAudit.Add([pscustomobject]@{
        bundle = $bundle
        unit = $unit
        indexHtml = $file.FullName
        hasPayload = $true
        exported = $true
        questionCount = $questions.Count
        outFile = $outFileRel
      })
  }
  catch {
    $errors.Add([pscustomobject]@{
        file = $file.FullName
        error = $_.Exception.Message
      })

    try {
      $relIndex = $file.FullName.Substring($assetsRootFull.Length)
      $relIndex = $relIndex.TrimStart([char[]]'\\/')
      $parts = $relIndex -split '[\\/]'
      $bundle = if ($parts.Length -ge 1) { $parts[0] } else { 'unknown' }
      $unit = if ($parts.Length -ge 3) {
        if ($parts[2] -eq 'index.html') { '(www-root)' } else { $parts[2] }
      } else { 'unknown' }
      $indexAudit.Add([pscustomobject]@{
          bundle = $bundle
          unit = $unit
          indexHtml = $file.FullName
          hasPayload = $true
          exported = $false
          reason = $_.Exception.Message
        })
    } catch {
    }
  }
}

# Write master indexes
$masterIndex = [pscustomobject]@{
  generatedAt = (Get-Date).ToString('s')
  assetsRoot = $assetsRootFull
  indexHtmlCount = $indexFiles.Count
  exportedIndexHtmlCount = $indexAudit.Where({$_.exported}).Count
  packageCount = $allPackages.Count
  packages = $allPackages
  indexAudit = $indexAudit
  errors = $errors
}

$masterIndex | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $OutDir 'index.json') -Encoding UTF8 -NoNewline

# by-category index
$byCategory = $allPackages | Group-Object category | ForEach-Object {
  [pscustomobject]@{
    category = $_.Name
    packages = $_.Group
  }
}
$byCategory | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $OutDir 'by-category.json') -Encoding UTF8 -NoNewline

# by-tier index
$byTier = $allPackages | Group-Object tier | ForEach-Object {
  [pscustomobject]@{
    tier = $_.Name
    packages = $_.Group
  }
}
$byTier | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $OutDir 'by-tier.json') -Encoding UTF8 -NoNewline

# by-bundle index
$byBundle = $allPackages | Group-Object bundle | ForEach-Object {
  [pscustomobject]@{
    bundle = $_.Name
    packages = $_.Group
  }
}
$byBundle | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $OutDir 'by-bundle.json') -Encoding UTF8 -NoNewline

# by-image index
$byImageOut = foreach ($k in ($byImage.Keys | Sort-Object)) {
  [pscustomobject]@{
    image = $k
    usedIn = @($byImage[$k])
  }
}
$byImageOut | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $OutDir 'by-image.json') -Encoding UTF8 -NoNewline

"Exported $($allPackages.Count) packages"
"Wrote: $(Join-Path $OutDir 'index.json')"
"Wrote: $(Join-Path $OutDir 'by-category.json')"
"Wrote: $(Join-Path $OutDir 'by-tier.json')"
"Wrote: $(Join-Path $OutDir 'by-bundle.json')"
"Wrote: $(Join-Path $OutDir 'by-image.json')"
if ($errors.Count -gt 0) { "Errors: $($errors.Count) (see index.json)" }
