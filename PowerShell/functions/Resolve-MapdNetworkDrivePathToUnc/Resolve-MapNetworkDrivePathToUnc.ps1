using namespace System.Collections.Generic
using namespace System.Management.Automation

function Resolve-MapNetworkDrivePathToUnc {
<#
.SYNOPSIS
割り当てられたネットワークドライブのパスをUNCパスに解決します。
.DESCRIPTION
Resolve-MapNetworkDrivePathToUnc 関数は、割り当てられたネットワークドライブのパスをUNCパスに解決します。
指定されたパスが相対パスの場合は、絶対パスに変換した上で解決が行われます。
すでにUNCパスの場合、あるいはローカルパスの場合は絶対パスを返します。
存在しているパスを指定する必要があります。
.INPUTS
System.String
.OUTPUTS
System.String
#>
    [CmdletBinding(
        DefaultParameterSetName = 'LiteralPath',
        HelpUri = 'https://qiita.com/nukie_53/items/673e9581624fc8f201bf'
    )]
    [OutputType([string])]
    param (
        # 解決するパスを指定します。LiteralPath の値は、入力した内容のまま使用されます。ワイルドカードとして解釈される文字はありません。
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
        # 解決するパスを指定します。ワイルドカードを使用できます。
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
        [Dictionary[string, string]]$driveUncDic = [Dictionary[string, string]]::new()
        
        # WMI 経由でドライブ情報を取得。
        # DeviceID が ドライブレター、ProviderName が 実際のパスとなる。
        [string]$Wql = 'SELECT DeviceID, ProviderName FROM Win32_MappedLogicalDisk'
        Get-CimInstance -Query $Wql | 
            ForEach-Object -Process {
                $driveUncDic.Add($_.DeviceID, $_.ProviderName)
                $_.Dispose()
            }
    }
    process {
        # Resolve-Path コマンドレットに、
        # LiteralPath が指定されたときには $LiteralPath、Path が指定されたときは $Path を渡したいため、スプラッティングを使用する。
        [hashtable]$resolvePathParam = @{
            $($PSCmdlet.ParameterSetName) = $PSBoundParameters[$PSCmdlet.ParameterSetName]
        }
        [PathInfo]$p = $null
        foreach ($p in Resolve-Path @resolvePathParam) {
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
            [string]$localPath = $u.LocalPath # この後よく使うので一旦変数に格納。

            if ($u.IsUnc) {
                # すでにUNCパスならそのままでOK。
                $PSCmdlet.WriteObject($localPath)
                continue
            }
            
            # continue の対象がわかりにくくになるため、foreach ではなく、ForEach-Object を使用。
            $driveUncDic.GetEnumerator() |
                ForEach-Object -Process {
                    # Resolve-Path によりパスの大文字小文字は正規化されているので、そのまま比較してOK。
                    if ($localPath.StartsWith($_.Key)) {
                        # パスの先頭がドライブレターなら、ドライブレター以降を切り出して実際のパスと結合。
                        [string]$resolvedPath = $_.Value + $localPath.Substring($_.Key.Length)
                        $PSCmdlet.WriteObject($resolvedPath)
                        continue
                    }
                }
            
            # ローカルのパス(C ドライブとか)
            $PSCmdlet.WriteObject($localPath)
        }
    }
}