[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [string]$IndexHtml,

  [string]$OutJson
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $IndexHtml)) {
  throw "Index.html not found: $IndexHtml"
}

Add-Type -AssemblyName System.Web.Extensions

function Get-QuizPayloadFromIndexHtml {
  param([string]$Path)

  $text = Get-Content -Raw -LiteralPath $Path
  $m = [regex]::Match($text, '\bvar\s+data\s*=\s*([\x27\x22])(?<b64>[A-Za-z0-9+/=]+)\1')
  if (-not $m.Success) {
    throw "Could not find base64 payload in: $Path"
  }

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
    [object]$Dict,
    [string]$Key
  )

  if ($null -eq $Dict) { return $null }
  if ($Dict -is [System.Collections.IDictionary]) {
    try {
      return $Dict[$Key]
    } catch {
      return $null
    }
  }

  if ($Dict.PSObject -and $Dict.PSObject.Properties.Name -contains $Key) {
    return $Dict.$Key
  }

  try {
    return $Dict[$Key]
  } catch {
  }

  return $null
}

function Get-Slides {
  param([object]$Root)

  $d = Try-Get $Root 'd'
  $sl = Try-Get $d 'sl'
  $groups = Try-Get $sl 'g'

  if (-not ($groups -is [System.Collections.IEnumerable])) {
    throw "Unexpected JSON shape: d.sl.g is missing"
  }

  $firstGroup = @($groups)[0]
  $slides = Try-Get $firstGroup 'S'
  if (-not ($slides -is [System.Collections.IEnumerable])) {
    throw "Unexpected JSON shape: d.sl.g[0].S is missing"
  }

  return @($slides)
}

$json = Get-QuizPayloadFromIndexHtml -Path $IndexHtml
$root = Parse-JsonDeep -Json $json

$slides = Get-Slides -Root $root

$results = New-Object System.Collections.Generic.List[object]
$questionNumber = 0

foreach ($slide in $slides) {
  $slideType = [string](Try-Get $slide 'tp')
  $choicesContainer = Try-Get $slide 'C'
  $choices = Try-Get $choicesContainer 'chs'

  if (-not ($choices -is [System.Collections.IEnumerable])) {
    continue
  }

  $questionNumber++

  $d = Try-Get $slide 'D'
  $questionText = [string](Try-Get $d 'd')

  $imageUri = $null
  $at = Try-Get $slide 'at'
  $ati = Try-Get $at 'i'
  $imageUri = [string](Try-Get $ati 'i')
  if ([string]::IsNullOrWhiteSpace($imageUri)) { $imageUri = $null }

  $imageRelativePath = $null
  if ($imageUri -and $imageUri.StartsWith('storage://images/')) {
    $imageRelativePath = 'data/images/' + $imageUri.Substring('storage://images/'.Length)
  }

  $options = New-Object System.Collections.Generic.List[object]
  foreach ($ch in @($choices)) {
    $choiceTextContainer = Try-Get $ch 't'
    $optionText = [string](Try-Get $choiceTextContainer 'd')

    $isCorrectRaw = Try-Get $ch 'c'
    $isCorrect = $false
    if ($null -ne $isCorrectRaw) {
      $isCorrect = [bool]$isCorrectRaw
    }

    $options.Add([pscustomobject]@{
        text    = $optionText
        correct = $isCorrect
      })
  }

  $results.Add([pscustomobject]@{
      n                 = $questionNumber
      slideId           = [string](Try-Get $slide 'i')
      type              = $slideType
      question          = $questionText
      imageUri          = $imageUri
      imageRelativePath = $imageRelativePath
      options           = $options
    })
}

if (-not $OutJson) {
  $OutJson = Join-Path (Split-Path -Parent $IndexHtml) 'answers.json'
}

$results | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $OutJson -Encoding UTF8 -NoNewline

"Extracted $($results.Count) questions"
"Wrote: $OutJson"
