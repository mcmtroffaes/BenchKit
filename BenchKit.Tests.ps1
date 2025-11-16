BeforeAll {
    Import-Module .\BenchKit.psm1 -Force
}

Describe "Get-ConfidenceInterval" {
    $testCases = @(
        @{ Numbers = @(-1,1); ExpectedLower = -1.96; ExpectedUpper = 1.96 },
        @{ Numbers = @(1,2,3,4,5); ExpectedLower = 3 - 1.96 * 0.7071; ExpectedUpper = 3 + 1.96 * 0.7071 },
        @{ Numbers = @(5,5,5,5); ExpectedLower = 5; ExpectedUpper = 5 },
        @{ Numbers = @(21,13,66,5); ExpectedLower = 26.25 - 1.96 * 13.6466; ExpectedUpper = 26.25 + 1.96 * 13.6466 }
        @{ Numbers = @(42); ExpectedLower = -[Double]::PositiveInfinity; ExpectedUpper = [Double]::PositiveInfinity }
        @{ Numbers = @(); ExpectedLower = -[Double]::PositiveInfinity; ExpectedUpper = [Double]::PositiveInfinity }
    )
    It -ForEach $testCases "Calculates correct 95% confidence interval for <Numbers>" {
        $result = Get-ConfidenceInterval -Numbers $_.Numbers
        $result[0] | Should -BeGreaterOrEqual ($_.ExpectedLower - 0.0001)
        $result[0] | Should -BeLessOrEqual ($_.ExpectedLower + 0.0001)
        $result[1] | Should -BeGreaterOrEqual ($_.ExpectedUpper - 0.0001)
        $result[1] | Should -BeLessOrEqual ($_.ExpectedUpper + 0.0001)
    }
}

Describe "Get-FormattedNumbers" {
    It -ForEach @(
        @{ Name = "Sample1"; Numbers = @(21,13,66,5); Expected = @("=== Sample1 ===", "21.000,13.000,66.000,5.000", "95% CI for mean: [-0.497, 52.997]") },
        @{ Name = "Sample2"; Numbers = @(1,2,3,4,5); Expected = @("=== Sample2 ===", "1.000,2.000,3.000,4.000,5.000", "95% CI for mean: [1.614, 4.386]") },
        @{ Name = "Single";   Numbers = @(42); Expected = @("=== Single ===", "42.000", "95% CI for mean: [-∞, ∞]") },
        @{ Name = "Empty";    Numbers = @(); Expected = @("=== Empty ===", "", "95% CI for mean: [-∞, ∞]") }
    ) "Formats numbers and CI correctly for <Name>" {
        $output = Get-FormattedNumbers -Name $_.Name -Numbers $_.Numbers
        $output | Should -Be $_.Expected
    }
}