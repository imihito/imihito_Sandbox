function Search-Type {
<#
.SYNOPSIS
Search type from loaded assemblies.
ロード済みアセンブリー内から型を検索します。
.DESCRIPTION
Search type by root type from Loaded assemblies.
ロード済みアセンブリー内から、指定した型及びサブクラスを検索します。
.EXAMPLE
[System.IFormatProvider] | Search-Type

IsPublic IsSerial Name               BaseType     
-------- -------- ----               --------     
True     False    IFormatProvider                 
True     True     CultureInfo        System.Object
True     True     DateTimeFormatInfo System.Object
True     True     NumberFormatInfo   System.Object

.EXAMPLE
Search-Type -RootType System.Collections.ICollection -Namespace System.Collections*

IsPublic IsSerial Name                             BaseType                                                           
-------- -------- ----                             --------                                                           
True     True     CollectionBase                   System.Object                                                      
True     True     DictionaryBase                   System.Object                                                      
True     True     ReadOnlyCollectionBase           System.Object                                                      
True     True     Queue                            System.Object                                                      
True     True     ArrayList                        System.Object                                                      
True     True     BitArray                         System.Object                                                      
True     True     Stack                            System.Object                                                      
True     True     Hashtable                        System.Object                                                      
True     False    ICollection

.INPUTS
System.Type
.OUTPUTS
System.Type
By default, return all public type in loaded assemblies.
#>
    [CmdletBinding()]
    [OutputType([type])]
    param (
        [Parameter(ValueFromPipeline=$true)]
        [type]$RootType = [System.Object]
        ,
        [SupportsWildcards()]
        [string]$Name
        ,
        [SupportsWildcards()]
        [string]$Namespace
        ,
        [SupportsWildcards()]
        [string]$FullName
        ,
        # Include not public types.
        [bool]$Force
    )
    begin {
        # Declare foreach variables for IntelliSense.
        [System.Reflection.Assembly]$asm = [type]$t = $null
    }
    process {
        foreach ($asm in [System.AppDomain]::CurrentDomain.GetAssemblies()) {
            foreach ($t in $asm.GetTypes()) {
                # Public only othewise Force switch.
                if (-not ($Force -or $t.IsPublic)) { continue }

                if (-not $RootType.IsAssignableFrom($t)) { continue }

                # Name check.
                if (-not [string]::IsNullOrEmpty($Name)      -and ($t.Name      -notlike $Name)      ) { continue }
                if (-not [string]::IsNullOrEmpty($Namespace) -and ($t.Namespace -notlike $Namespace) ) { continue }
                if (-not [string]::IsNullOrEmpty($FullName)  -and ($t.FullName  -notlike $FullName)  ) { continue }

                Write-Output -InputObject $t
            }
        }
    }
}