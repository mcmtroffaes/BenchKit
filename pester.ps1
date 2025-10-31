$config = New-PesterConfiguration
$config.Output.Verbosity = "Detailed"
$config.CodeCoverage.Path = "BenchKit.psm1"
$config.CodeCoverage.Enabled = $true
Invoke-Pester -Configuration $config
