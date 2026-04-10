[CmdletBinding()]
param(
  [string]$ExportDir = '.\content_export',
  [string]$OutHtml = '.\quiz.html'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Web.Extensions
Add-Type -AssemblyName System.Web

function Parse-JsonDeep {
  param([string]$Json)

  $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
  $ser.MaxJsonLength = [int]::MaxValue
  $ser.RecursionLimit = 4000
  return $ser.DeserializeObject($Json)
}

function HtmlEncode {
  param([string]$Text)
  if ($null -eq $Text) { return '' }
  return [System.Web.HttpUtility]::HtmlEncode($Text)
}

function Normalize-Slash {
  param([string]$Path)
  if ($null -eq $Path) { return $null }
  return $Path -replace '\\','/'
}

function Get-CategoryInfo {
  param([object]$Pkg)

  $rawCategory = [string]$Pkg.category
  switch ($rawCategory.ToLowerInvariant()) {
    'cde' {
      return [pscustomobject]@{ key = 'codigo'; label = 'Codigo de estrada'; order = 1; raw = $rawCategory }
    }
    'sdt' {
      return [pscustomobject]@{ key = 'sinais'; label = 'Sinais de transito'; order = 2; raw = $rawCategory }
    }
    'exame' {
      return [pscustomobject]@{ key = 'exames'; label = 'Exames predefinidos'; order = 3; raw = $rawCategory }
    }
    'revisao' {
      return [pscustomobject]@{ key = 'revisao'; label = 'Revisao geral'; order = 4; raw = $rawCategory }
    }
    'mec' {
      return [pscustomobject]@{ key = 'revisao'; label = 'Revisao geral'; order = 4; raw = $rawCategory }
    }
    'cg' {
      return [pscustomobject]@{ key = 'revisao'; label = 'Revisao geral'; order = 4; raw = $rawCategory }
    }
    default {
      return [pscustomobject]@{ key = 'revisao'; label = 'Revisao geral'; order = 4; raw = $rawCategory }
    }
  }
}

$exportDirFull = (Resolve-Path -LiteralPath $ExportDir).Path
$indexPath = Join-Path $exportDirFull 'index.json'
if (-not (Test-Path -LiteralPath $indexPath)) {
  throw "Missing export index: $indexPath"
}

$indexObj = Parse-JsonDeep -Json (Get-Content -Raw -LiteralPath $indexPath)
$packages = @($indexObj.packages)

$packagesDecorated = foreach ($pkg in $packages) {
  $ci = Get-CategoryInfo -Pkg $pkg
  [pscustomobject]@{
    pkg = $pkg
    categoryKey = $ci.key
    categoryLabel = $ci.label
    categoryOrder = $ci.order
    rawCategory = $ci.raw
  }
}

# Order packages by requested category sequence, then tier/bundle/unit.
$packagesDecorated = $packagesDecorated | Sort-Object categoryOrder, @{Expression = { [string]$_.pkg.tier } }, @{Expression = { [string]$_.pkg.bundle } }, @{Expression = { [string]$_.pkg.unit } }

$sb = New-Object System.Text.StringBuilder

# Header
[void]$sb.AppendLine('<!doctype html>')
[void]$sb.AppendLine('<html lang="pt">')
[void]$sb.AppendLine('<head>')
[void]$sb.AppendLine('  <meta charset="utf-8">')
[void]$sb.AppendLine('  <meta name="viewport" content="width=device-width, initial-scale=1">')
[void]$sb.AppendLine('  <title>Quiz (Export)</title>')
[void]$sb.AppendLine('  <style>')
[void]$sb.AppendLine('    body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; margin: 16px; line-height: 1.35; background: #f8fafc; }')
[void]$sb.AppendLine('    .wrap { max-width: 1200px; margin: 0 auto; }')
[void]$sb.AppendLine('    .meta { color: #444; font-size: 14px; margin-bottom: 16px; }')
[void]$sb.AppendLine('    .actions { margin: 10px 0 16px; display: flex; gap: 10px; flex-wrap: wrap; }')
[void]$sb.AppendLine('    .btn { border: 1px solid #0f172a; background: #0f172a; color: #fff; border-radius: 8px; padding: 8px 12px; cursor: pointer; font-size: 14px; }')
[void]$sb.AppendLine('    .cat-nav { display: flex; gap: 8px; flex-wrap: wrap; margin: 8px 0 14px; }')
[void]$sb.AppendLine('    .cat-nav a { text-decoration: none; border: 1px solid #d1d5db; color: #111827; background: #fff; border-radius: 999px; padding: 6px 10px; font-size: 13px; }')
[void]$sb.AppendLine('    .category { margin: 18px 0 28px; border: 2px solid #e5e7eb; border-radius: 12px; padding: 12px; background: #ffffff; }')
[void]$sb.AppendLine('    .category h2 { margin: 2px 0 10px; font-size: 22px; }')
[void]$sb.AppendLine('    details { border: 1px solid #ddd; border-radius: 8px; padding: 10px 12px; margin: 10px 0; }')
[void]$sb.AppendLine('    summary { cursor: pointer; font-weight: 600; }')
[void]$sb.AppendLine('    .pkg-info { color: #555; font-weight: 400; }')
[void]$sb.AppendLine('    .q { padding: 10px 0; border-top: 1px dashed #e5e5e5; }')
[void]$sb.AppendLine('    .q:first-of-type { border-top: none; }')
[void]$sb.AppendLine('    .q-head { display: flex; gap: 12px; flex-wrap: wrap; align-items: flex-start; }')
[void]$sb.AppendLine('    .q-img { max-width: 360px; width: 100%; height: auto; border: 1px solid #eee; border-radius: 6px; }')
[void]$sb.AppendLine('    .q-text { min-width: 260px; flex: 1; }')
[void]$sb.AppendLine('    .q-num { color: #666; font-size: 12px; margin-bottom: 6px; }')
[void]$sb.AppendLine('    .opts { margin-top: 10px; }')
[void]$sb.AppendLine('    .opt { display: flex; gap: 8px; margin: 6px 0; align-items: flex-start; }')
[void]$sb.AppendLine('    .opt input { margin-top: 4px; }')
[void]$sb.AppendLine('    .opt-correct { border-radius: 6px; padding: 6px 8px; }')
[void]$sb.AppendLine('    .correct-tag { display: none; margin-left: 8px; padding: 2px 6px; border-radius: 999px; background: #166534; color: #fff; font-size: 11px; font-weight: 700; }')
[void]$sb.AppendLine('    .show-correct .opt-correct { background: #ecfdf3; border: 1px solid #bbf7d0; }')
[void]$sb.AppendLine('    .show-correct .correct-tag { display: inline-block; }')
[void]$sb.AppendLine('    .badge { display: inline-block; padding: 2px 8px; border: 1px solid #ddd; border-radius: 999px; font-size: 12px; margin-left: 8px; }')
[void]$sb.AppendLine('  </style>')
[void]$sb.AppendLine('</head>')
[void]$sb.AppendLine('<body class="show-correct">')
[void]$sb.AppendLine('<div class="wrap">')

$generatedAt = HtmlEncode([string]$indexObj.generatedAt)
$assetsRoot = HtmlEncode([string]$indexObj.assetsRoot)
$pkgCount = [int]$indexObj.packageCount

[void]$sb.AppendLine('<h1>Quiz (Export)</h1>')
[void]$sb.AppendLine(('<div class="meta">Gerado: {0}<br>Assets: {1}<br>Pacotes: {2}</div>' -f $generatedAt, $assetsRoot, $pkgCount))
[void]$sb.AppendLine('<div class="actions">')
[void]$sb.AppendLine('  <button id="toggleAnswers" class="btn" type="button">Ocultar respostas corretas</button>')
[void]$sb.AppendLine('</div>')

$presentCategoryKeys = @($packagesDecorated | Select-Object -ExpandProperty categoryKey -Unique)
[void]$sb.AppendLine('<div class="cat-nav">')
if ($presentCategoryKeys -contains 'codigo') { [void]$sb.AppendLine('  <a href="#cat-codigo">Codigo de estrada</a>') }
if ($presentCategoryKeys -contains 'sinais') { [void]$sb.AppendLine('  <a href="#cat-sinais">Sinais de transito</a>') }
if ($presentCategoryKeys -contains 'exames') { [void]$sb.AppendLine('  <a href="#cat-exames">Exames predefinidos</a>') }
if ($presentCategoryKeys -contains 'revisao') { [void]$sb.AppendLine('  <a href="#cat-revisao">Revisao geral</a>') }
[void]$sb.AppendLine('</div>')

$currentCategory = $null

foreach ($row in $packagesDecorated) {
  $pkg = $row.pkg
  $categoryKey = [string]$row.categoryKey
  $categoryLabel = [string]$row.categoryLabel

  if ($categoryKey -ne $currentCategory) {
    if ($null -ne $currentCategory) {
      [void]$sb.AppendLine('</section>')
    }
    $currentCategory = $categoryKey
    [void]$sb.AppendLine(('<section id="cat-{0}" class="category">' -f (HtmlEncode($categoryKey))))
    [void]$sb.AppendLine(('<h2>{0}</h2>' -f (HtmlEncode($categoryLabel))))
  }

  $bundle = [string]$pkg.bundle
  $tier = [string]$pkg.tier
  $unit = [string]$pkg.unit
  $title = [string]$pkg.title
  $qCount = [int]$pkg.questionCount
  $outFileRel = Normalize-Slash([string]$pkg.outFile)

  $contentPath = Join-Path $exportDirFull ($outFileRel -replace '/','\\')
  if (-not (Test-Path -LiteralPath $contentPath)) {
    continue
  }

  $contentObj = Parse-JsonDeep -Json (Get-Content -Raw -LiteralPath $contentPath)
  $questions = @($contentObj.questions)

  $sum = ('{0} <span class="badge">{1}</span> <span class="badge">{2}</span> <span class="pkg-info"> &mdash; {3} &mdash; {4} perguntas</span>' -f
    (HtmlEncode($title)),
    (HtmlEncode($tier)),
    (HtmlEncode($bundle)),
    (HtmlEncode($unit)),
    $qCount)

  [void]$sb.AppendLine('<details>')
  [void]$sb.AppendLine(('  <summary>{0}</summary>' -f $sum))

  foreach ($q in $questions) {
    $n = [int]$q.n
    $slideType = HtmlEncode([string]$q.type)

    $qId = "q_${bundle}_${unit}_${n}" -replace '[^A-Za-z0-9_]','_'

    $questionHtml = [string]$q.questionHtml
    $questionText = [string]$q.question

    $imgRel = [string]$q.imageRelativePath
    $imgSrc = $null
    if (-not [string]::IsNullOrWhiteSpace($imgRel)) {
      $imgSrc = "1/assets/flutter_assets/assets/$bundle/www/$unit/$imgRel"
      $imgSrc = Normalize-Slash $imgSrc
    }

    [void]$sb.AppendLine('  <div class="q">')
    [void]$sb.AppendLine('    <div class="q-head">')

    if ($imgSrc) {
      $imgAlt = HtmlEncode($questionText)
      [void]$sb.AppendLine(('      <img class="q-img" loading="lazy" src="{0}" alt="{1}">' -f $imgSrc, $imgAlt))
    }

    [void]$sb.AppendLine('      <div class="q-text">')
    [void]$sb.AppendLine(('        <div class="q-num">Pergunta {0} &mdash; {1}</div>' -f $n, $slideType))

    if (-not [string]::IsNullOrWhiteSpace($questionHtml)) {
      [void]$sb.AppendLine(('        <div class="q-html">{0}</div>' -f $questionHtml))
    } else {
      [void]$sb.AppendLine(('        <div class="q-plain">{0}</div>' -f (HtmlEncode($questionText))))
    }

    [void]$sb.AppendLine('        <div class="opts">')

    $options = @($q.options)
    $optIndex = 0
    foreach ($opt in $options) {
      $optIndex++
      $optHtml = [string]$opt.html
      $optText = [string]$opt.text
      $isCorrect = $false
      try {
        $isCorrect = [bool]$opt.correct
      }
      catch {
      }

      $inputId = "${qId}_opt$optIndex"
      $inputIdEnc = HtmlEncode($inputId)
      $qIdEnc = HtmlEncode($qId)
      $optClass = if ($isCorrect) { 'opt opt-correct' } else { 'opt' }

      [void]$sb.AppendLine(('          <div class="{0}">' -f $optClass))
      [void]$sb.AppendLine(('            <input type="checkbox" id="{0}" name="{1}" value="{2}">' -f $inputIdEnc, $qIdEnc, $optIndex))
      if (-not [string]::IsNullOrWhiteSpace($optHtml)) {
        [void]$sb.AppendLine(('            <label for="{0}">{1}</label>' -f $inputIdEnc, $optHtml))
      } else {
        [void]$sb.AppendLine(('            <label for="{0}">{1}</label>' -f $inputIdEnc, (HtmlEncode($optText))))
      }
      if ($isCorrect) {
        [void]$sb.AppendLine('            <span class="correct-tag">Correta</span>')
      }
      [void]$sb.AppendLine('          </div>')
    }

    [void]$sb.AppendLine('        </div>')
    [void]$sb.AppendLine('      </div>')
    [void]$sb.AppendLine('    </div>')
    [void]$sb.AppendLine('  </div>')
  }

  [void]$sb.AppendLine('</details>')
}

if ($null -ne $currentCategory) {
  [void]$sb.AppendLine('</section>')
}

[void]$sb.AppendLine('<script>')
[void]$sb.AppendLine('(function () {')
[void]$sb.AppendLine('  var btn = document.getElementById("toggleAnswers");')
[void]$sb.AppendLine('  if (!btn) return;')
[void]$sb.AppendLine('  btn.addEventListener("click", function () {')
[void]$sb.AppendLine('    document.body.classList.toggle("show-correct");')
[void]$sb.AppendLine('    if (document.body.classList.contains("show-correct")) {')
[void]$sb.AppendLine('      btn.textContent = "Ocultar respostas corretas";')
[void]$sb.AppendLine('    } else {')
[void]$sb.AppendLine('      btn.textContent = "Mostrar respostas corretas";')
[void]$sb.AppendLine('    }')
[void]$sb.AppendLine('  });')
[void]$sb.AppendLine('})();')
[void]$sb.AppendLine('</script>')

[void]$sb.AppendLine('</div>')
[void]$sb.AppendLine('</body>')
[void]$sb.AppendLine('</html>')

$outHtmlFull = $OutHtml
if (-not [System.IO.Path]::IsPathRooted($outHtmlFull)) {
  $outHtmlFull = Join-Path (Get-Location) $outHtmlFull
}
$sb.ToString() | Set-Content -LiteralPath $outHtmlFull -Encoding UTF8 -NoNewline

"Wrote: $outHtmlFull"
"Note: Open quiz.html in a browser. Images resolve relative to workspace root."