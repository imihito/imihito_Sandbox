# 使用するアセンブリや名前空間の指定(PowerShell 5.1以降の機能)
using namespace System

# UIAutomation 関連のアセンブリ群
using assembly  UIAutomationClient
using assembly  UIAutomationTypes
using assembly  UIAutomationClientSideProviders
using namespace System.Windows.Automation

# SendKeys 用のアセンブリ
using assembly  System.Windows.Forms
using namespace System.Windows.Forms

function Send-UIAKeys {
<#
.SYNOPSIS
対象の要素にキーストロークを送信します。
.DESCRIPTION
$InputObjectで指定された要素にキーストロークを送信します。
System.Windows.Forms.SendKeys.SendWaitを使用するため、アクティブなウィンドウが変更されます。
#>
    [CmdletBinding()]
    [OutputType([System.Windows.Automation.AutomationElement])]
    Param(
        
        # キーストロークを送信する要素を指定します。
        # キーボードフォーカスを受け取ることが出来ればフォーカスし、そうでなければ直近の親ウィンドウを最前面にします。
        # このパラメーターは必須です。
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [AutomationElement]$InputObject
        ,
        # 送信するキーストロークを指定します。
        # System.Windows.Forms.SendKeys クラスと同じ形式で文字列を指定します。
        # https://docs.microsoft.com/ja-jp/dotnet/api/system.windows.forms.sendkeys?view=netframework-4.8
        # このパラメーターは必須です。
        [Parameter(Mandatory = $true)]
        [string]$Keys
        ,
        # キーストローク送信後、待機する時間をミリ秒単位で指定します。
        # $RestoreFocus スイッチを指定する場合は意図した動作になるよう調整が必要です。
        [ValidateRange(0, [int]::MaxValue)]
        [Alias('ms')]
        [int]$WaitMilliseconds = 0
        ,
        # キーストローク送信後、フォーカスを直前の要素に戻します。
        # 既定では、キーストロークを送信した要素が最前面となります。
        # 指定する場合、$WaitMilliseconds の値も適切な値に変更する必要があります。
        # この関数を連続して実行する場合、期待した結果が得られないことがあります。
        [switch]$RestoreFocus
        ,
        # $InputObject を再度パイプラインに出力します。
        # 既定では、この関数による出力はありません。
        [switch]$PassThru
    )

    Process {
        # 現在のフォーカスを取得。
        [AutomationElement]$currentFocus =　[AutomationElement]::FocusedElement
        

        # 親ウィンドウを取得するため、WindowPatternを実装している要素を探すTreeWalkerを作成。
        [TreeWalker]$windowWalker = [TreeWalker]::new(
            [PropertyCondition]::new(
                [AutomationElement]::IsWindowPatternAvailableProperty, 
                $true
            )
        )

        # 親ウィンドウ取得。
        [AutomationElement]$parentWin = $windowWalker.Normalize($InputObject)
        if ($null -eq $parentWin) {
            # 取得できなかった場合は強制停止。
            $PSCmdlet.ThrowTerminatingError([Management.Automation.ErrorRecord]::new(
                [InvalidOperationException]::new('親ウィンドウを取得できません。'),
                'ParentWindowNotFound',
                [Management.Automation.ErrorCategory]::NotEnabled,
                $InputObject
            ))
        }

        # フォーカスの変更。
        # WindowPattern.SetWindowVisualState(最大化・最小化などの変更)を行うと、
        # 現在の状態にかかわらずそのウィンドウが最前面になることを利用。
        [WindowPattern]$winPtn = $parentWin.GetCurrentPattern([WindowPattern]::Pattern)
        [WindowVisualState]$visState = $winPtn.Current.WindowVisualState
        if ($visState -eq [WindowVisualState]::Minimized) {
            # 最小化されている場合は通常に戻す。
            $visState = [WindowVisualState]::Normal
        }
        $winPtn.SetWindowVisualState($visState)

        if ($parentWin.Current.IsKeyboardFocusable) {
            $parentWin.SetFocus()
        }
        if ($InputObject.Current.IsKeyboardFocusable) {
            $InputObject.SetFocus()
        }
        
        # キーストローク送信。
        [SendKeys]::SendWait($Keys)

        # 送信後の待機。
        while (-not $winPtn.WaitForInputIdle($WaitMilliseconds)) {
        }
        Start-Sleep -Milliseconds $WaitMilliseconds


        if ($RestoreFocus) {
            # フォーカスを戻す。
            $currentFocus.SetFocus()
        }
    }
}