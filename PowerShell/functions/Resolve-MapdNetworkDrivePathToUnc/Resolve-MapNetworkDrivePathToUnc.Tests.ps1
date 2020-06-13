$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"

Describe "Resolve-MapNetworkDrivePathToUnc" {
    It "LiteralPath" {
        Resolve-MapNetworkDrivePathToUnc -LiteralPath 'Z:\' | Should Be '\\nasne-5c5761\share2\'
        Resolve-MapNetworkDrivePathToUnc -LiteralPath 'z:\' | Should Be '\\nasne-5c5761\share2\'
        'Z:\' | Resolve-MapNetworkDrivePathToUnc | Should Be '\\nasne-5c5761\share2\'
        'z:\' | Resolve-MapNetworkDrivePathToUnc | Should Be '\\nasne-5c5761\share2\'
    }
    It "Path" {
        Resolve-MapNetworkDrivePathToUnc -Path 'Z:\B*' | Should Be '\\nasne-5c5761\share2\Book1.xlsx'
        Resolve-MapNetworkDrivePathToUnc -Path 'z:\b*' | Should Be '\\nasne-5c5761\share2\Book1.xlsx'
    }


    It "RelativeLiteralPath" {
        In 'Z:\' {
            Resolve-MapNetworkDrivePathToUnc -LiteralPath '.' | Should Be '\\nasne-5c5761\share2\'
            '.' | Resolve-MapNetworkDrivePathToUnc | Should Be '\\nasne-5c5761\share2\'
        }
    }

    It "RelativePath" {
        In 'Z:\' {
            Resolve-MapNetworkDrivePathToUnc -Path '.\B*' | Should Be '\\nasne-5c5761\share2\Book1.xlsx'
            Resolve-MapNetworkDrivePathToUnc -Path 'B*' | Should Be '\\nasne-5c5761\share2\Book1.xlsx'
        }
    }

    It "NoChange" {
        $path = 'C:\Windows'
        Resolve-MapNetworkDrivePathToUnc -LiteralPath $path | Should Be $path
        Resolve-MapNetworkDrivePathToUnc -LiteralPath $path.ToLower() | Should Be $path
        $path = 'C:\Windows\System3*'
        Resolve-MapNetworkDrivePathToUnc -Path $path.ToLower() | Should Be 'C:\Windows\System32'
    }
    
    It "InvalidDrive" {
        'C:\Windows\System3*' | Resolve-MapNetworkDrivePathToUnc | Should Throw
        Resolve-MapNetworkDrivePathToUnc -LiteralPath 'Env:\ALLUSERSPROFILE' |
            Should Throw
    }
}
