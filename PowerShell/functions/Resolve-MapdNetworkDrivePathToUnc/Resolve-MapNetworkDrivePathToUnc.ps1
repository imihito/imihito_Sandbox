using namespace System.Collections.Generic
using namespace System.Management.Automation

function Resolve-MapNetworkDrivePathToUnc {
<#
.SYNOPSIS
割り当てられたネットワークドライブのパスをUNCパスに解決します。
.DESCRIPTION
Long description

.EXAMPLE
An example

.NOTES
General notes
#>
    [CmdletBinding(
        DefaultParameterSetName = 'LiteralPath',
        HelpUri = 'https://qiita.com/nukie_53/items/673e9581624fc8f201bf'
    )]
    [OutputType([string])]
    param (
        # 
        [Parameter(
            Mandatory = $true, 
            Position = 0,
            ValueFromPipeline = $true, 
            ValueFromPipelineByPropertyName = $true,
            ParameterSetName = 'LiteralPath'
        )]
        [ValidateNotNullOrEmpty()]
        [Alias('PSPath', 'ProviderPath', 'FullName', 'LocalPath')]
        [string[]]$LiteralPath
        ,
        # Specifies a path to one or more locations.
        [Parameter(
            Mandatory = $true,
            ParameterSetName = 'Path'
        )]
        [ValidateNotNullOrEmpty()]
        [SupportsWildcards()]
        [string[]]$Path
    )
    begin {
        # 割り当てられたネットワークドライブの一覧をDictionaryに格納。
        [Dictionary[string, string]]$driveUncDic = 
            # 念のため大文字・小文字無視の辞書とする
            [Dictionary[string, string]]::new([System.StringComparer]::InvariantCultureIgnoreCase)
        # WMI 経由でドライブ情報を取得。
        # DriveType が 4 のものがネットワークドライブで、
        # DeviceID が ドライブレター、ProviderName が 実際のパスとなる。
        Get-CimInstance -Query '
            SELECT 
                DeviceID, ProviderName 
            FROM Win32_LogicalDisk 
                WHERE DriveType = 4' | 
            ForEach-Object -Process {
                $driveUncDic.Add($_.DeviceID, $_.ProviderName)
                $_.Dispose()
            }
    }
    process {
        [hashtable]$resolvePathParam = @{
            $($PSCmdlet.ParameterSetName) = $PSBoundParameters[$PSCmdlet.ParameterSetName]
        }
        [PathInfo]$p = $null
        foreach ($p in Resolve-Path @resolvePathParam -ErrorAction Stop) {
            if ($p.Provider.Name -ne 'FileSystem') {
                # ファイル以外のものが指定された場合はエラーを出してスキップ。
                $PSCmdlet.WriteError([ErrorRecord]::new(
                    [System.ArgumentException]::new(('${0} には実際に存在するファイルのパスを指定してください。' -f $PSCmdlet.ParameterSetName)),
                    'InvalidArgument',
                    [ErrorCategory]::InvalidArgument,
                    $p.ProviderPath
                ))
                continue
            }

            [uri]$u = [uri]($p.ProviderPath)
            [string]$localPath = $u.LocalPath
            if (-not $u.IsFile) {
                $PSCmdlet.WriteObject($localPath)
                continue
            }
            if ($u.IsUnc) {
                $PSCmdlet.WriteObject($localPath)
                continue
            }
            
            $driveUncDic.GetEnumerator() |
                ForEach-Object -Process {
                    if ($localPath.StartsWith($_.Key)) {
                        [string]$resolvedPath = $_.Value + $localPath.Substring($_.Key.Length)
                        $PSCmdlet.WriteObject($resolvedPath)
                        continue
                    }
                }
            
            $PSCmdlet.WriteObject($localPath)
        }
    }
}