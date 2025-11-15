BeforeAll {
    Import-Module .\BenchKit.psm1 -Force
}

Describe "Get-AverageStderr" {
    $testCases = @(
        @{ Numbers = @(1,2,3,4,5); ExpectedAverage = 3; ExpectedStdErr = 0.7071 },
        @{ Numbers = @(5,5,5,5); ExpectedAverage = 5; ExpectedStdErr = 0 },
        @{ Numbers = @(21,13,66,5); ExpectedAverage = 26.25; ExpectedStdErr = 13.6466 }
    )
    It -ForEach $testCases "Calculates correct Average and StdErr for <Numbers>" {
        $result = Get-AverageStdErr -Numbers $_.Numbers
        $result.Average | Should -BeGreaterOrEqual ($_.ExpectedAverage - 0.0001)
        $result.Average | Should -BeLessOrEqual ($_.ExpectedAverage + 0.0001)
        $result.StdErr | Should -BeGreaterOrEqual ($_.ExpectedStdErr - 0.0001)
        $result.StdErr | Should -BeLessOrEqual ($_.ExpectedStdErr + 0.0001)
    }
    It "Returns Infinity for StdErr when only one number is supplied" {
        $result = Get-AverageStdErr @(42)
        $result.StdErr | Should -Be ([Double]::PositiveInfinity)
        $result.Average | Should -Be 42
    }

    It -ForEach @() "Returns NaN for average and Infinity for stderr when empty array is supplied" {
        $result = Get-AverageStdErr @()
        [Double]::IsNaN($result.Average) | Should -BeTrue
        $result.StdErr  | Should -Be ([Double]::PositiveInfinity)
    }
}

Describe "Get-ConfidenceInterval" {
    It "Calculates correct 95% confidence interval for @(21,13,66,5)" {
        $ci = Get-ConfidenceInterval @(21,13,66,5)
        $expectedLower = 26.25 - 1.96 * 13.6466
        $expectedUpper = 26.25 + 1.96 * 13.6466
        $tol = 0.001
        ([Math]::Abs($ci[0] - $expectedLower) -le $tol) | Should -BeTrue
        ([Math]::Abs($ci[1] - $expectedUpper) -le $tol) | Should -BeTrue
    }

    It -Foreach @(42),@() "Returns correct confidence interval for <_>" {
        $ci = Get-ConfidenceInterval $_
        $ci[0] | Should -Be (-[Double]::PositiveInfinity)
        $ci[1] | Should -Be ([Double]::PositiveInfinity)
    }
}

Describe "Get-FormattedNumbers" {
    It -ForEach @(
        @{ Name = "Sample1"; Numbers = @(21,13,66,5); Expected = @("=== Sample1 ===", "21.000,13.000,66.000,5.000", "95% CI for sample mean: [-0.497, 52.997]") },
        @{ Name = "Sample2"; Numbers = @(1,2,3,4,5); Expected = @("=== Sample2 ===", "1.000,2.000,3.000,4.000,5.000", "95% CI for sample mean: [1.614, 4.386]") },
        @{ Name = "Single";   Numbers = @(42); Expected = @("=== Single ===", "42.000", "95% CI for sample mean: [-∞, ∞]") },
        @{ Name = "Empty";    Numbers = @(); Expected = @("=== Empty ===", "", "95% CI for sample mean: [-∞, ∞]") }
    ) "Formats numbers and CI correctly for <Name>" {
        $output = Get-FormattedNumbers -Name $_.Name -Numbers $_.Numbers
        $output | Should -Be $_.Expected
    }
}