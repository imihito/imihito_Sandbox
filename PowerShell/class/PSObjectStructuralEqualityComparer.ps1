using namespace System.Collections
using namespace System.Management.Automation

Set-StrictMode -Version Latest
class PSObjectStructuralEqualityComparer : IEqualityComparer, Generic.IEqualityComparer[psobject] {
    # PSObjectStructuralEqualityComparer () { }

    [bool]Equals([System.Object]$x, [System.Object]$y) { 
        return [System.Object]::Equals($x, $y)
    }
    [bool]Equals([psobject]$x, [psobject]$y) {
        [PSPropertyInfo[]]$xProps = $x.psobject.Properties | Sort-Object -Property Name
        [PSPropertyInfo[]]$yProps = $y.psobject.Properties | Sort-Object -Property Name
        
        if ($xProps.Length -ne $yProps.Length) { return $false }

        for ($i = 0; $i -lt $xProps.Length; ++$i) {
            if ($xProps[$i].Name -ne $yProps[$i].Name) { return $false }
            
            if ($xProps[$i].Value -is [IStructuralEquatable]) {
                if (-not [StructuralComparisons]::StructuralEqualityComparer.Equals($xProps[$i].Value, $yProps[$i].Value)) {
                    return $false
                }
                # continue
            } elseif ($xProps[$i].Value -is [IEnumerable]) {
                if (-not [System.Object]::Equals($xProps[$i].Value, $yProps[$i].Value)) { return $false }
                # continue
            } else {
                if ($xProps[$i].Value -ne $yProps[$i].Value) { return $false }
                # continue
            }
        }
        
        return $true
    }

    [int]GetHashCode([System.Object]$obj) { return $obj.GetHashCode() }
    [int]GetHashCode([psobject]$obj)      { return $obj.GetHashCode() }
}


<#
.SYNOPSIS
   短い説明
.DESCRIPTION
   詳しい説明
.EXAMPLE
   このコマンドレットの使用方法の例
.EXAMPLE
   このコマンドレットの使用方法の別の例
.INPUTS
   このコマンドレットへの入力 (存在する場合)
.OUTPUTS
   このコマンドレットからの出力 (存在する場合)
.NOTES
   全般的な注意
.COMPONENT
   このコマンドレットが属するコンポーネント
.ROLE
   このコマンドレットが属する役割
.FUNCTIONALITY
   このコマンドレットの機能
#>

function New-PSObjectStructuralEqualityComparer {
    [CmdletBinding()]
    [OutputType([Generic.IEqualityComparer[psobject]])]
    param (
    )
    begin {
        return [PSObjectStructuralEqualityComparer]::new()
    }
}

$comp = [PSObjectStructuralEqualityComparer]::new()

$a = [pscustomobject]@{
    a = $null
    b = "b"
    c = 1..2
}
$b = [pscustomobject]@{
    b = "b"
    a = 1
    c = 1..2
}
$c = [pscustomobject]@{
    a = "1"
    b = "a"
    c = 1..2
}

$comp.Equals($a, $b)
#$comp.Equals($a, $c)