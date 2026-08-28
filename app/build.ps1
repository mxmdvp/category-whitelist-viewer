$ErrorActionPreference = 'Stop'

$appDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoDir = Split-Path -Parent $appDir
$distDir = Join-Path $repoDir 'dist'
$outputPath = Join-Path $distDir 'category-permissions.html'

$template = Get-Content -Raw -LiteralPath (Join-Path $appDir 'index.html')
$styles = Get-Content -Raw -LiteralPath (Join-Path $appDir 'styles.css')
$sheetJs = Get-Content -Raw -LiteralPath (Join-Path $appDir 'vendor\xlsx.full.min.js')
$application = Get-Content -Raw -LiteralPath (Join-Path $appDir 'app.js')

$sheetJs = $sheetJs -replace '</script', '<\/script'
$application = $application -replace '</script', '<\/script'

$html = $template.Replace('<link rel="stylesheet" href="styles.css">', "<style>`n$styles`n</style>")
$html = $html.Replace('<script src="vendor/xlsx.full.min.js"></script>', "<script>`n$sheetJs`n</script>")
$html = $html.Replace('<script src="app.js"></script>', "<script>`n$application`n</script>")

New-Item -ItemType Directory -Force -Path $distDir | Out-Null
[System.IO.File]::WriteAllText($outputPath, $html, [System.Text.UTF8Encoding]::new($false))

Write-Output "Built $outputPath"
