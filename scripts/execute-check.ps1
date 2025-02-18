Param(
    [String]$Arg1
)

. "$PSScriptRoot\_automation-utils.ps1"
. "$PSScriptRoot\_config.ps1"

#=======================================================================
# チェック対象ファイルリストの取得
#=======================================================================
$targetFileList = @()

if ($Arg1 -eq "ALL") {
    $targetFileList = $(Get-ChildItem -LiteralPath $CHECK_TARGET_ROOT_DIR -Directory -Recurse -Include $CHECK_TARGET_INCLUDE_DIR_NAMES -Exclude $CHECK_TARGET_EXCLUDE_DIR_NAMES | `
            ForEach-Object { Get-ChildItem -LiteralPath ($_.FullName + "\*") -File -Include $CHECK_TARGET_INCLUDE_FILE_NAMES -Exclude $CHECK_TARGET_EXCLUDE_FILE_NAMES })
}
else {
    $targetFileList = $(Get-ChildItem -Path $Arg1)
}

#=======================================================================
# チェック対象ファイルの内容を読み込み
#=======================================================================
class OutputInfo {
    [string]$Heading
    [string]$Content
    OutputInfo([string]$heading, [string]$content) {
        $this.Heading = $heading
        $this.Content = $content
    }
}

$outputInfos = @()
foreach ($file in $targetFileList) {
    #---------------
    # 見出し
    #---------------
    $heading = $file.Name

    #---------------
    # 内容
    #---------------
    $content = Get-Content -Path $file.FullName -Encoding UTF8

    # 行番号の最大桁数
    $maxLength = $content.Length.ToString().Length

    # 行番号をファイルの各行の先頭に付与
    $lineNumber = 1
    $content2 = $content | ForEach-Object { "{0,$maxLength}: {1}" -f $lineNumber++, $_ }
    $content2 = $content2 -join "`r`n"

    $formattedContent = ""
    $formattedContent += "・ファイルパス　　  ：" + $file.FullName + "`r`n"
    $formattedContent += "・ファイル更新日時：" + $file.LastWriteTime.ToString("yyyy/MM/dd HH:mm:ss") + "`r`n"
    $formattedContent += "・ファイルサイズ　  ：" + $file.Length + " bytes`r`n"
    $formattedContent += $content2

    $outputInfos += [OutputInfo]::new($heading, $formattedContent)
}

#=======================================================================
# チェック用Wordファイルの作成
#=======================================================================
$DATETIME = $(Get-Date -Format "yyyyMMddHHmmss")
$CHECK_TARGET_WD_IN_DIR = $(Join-Path $CHECK_TOOL_IN_DIR   $CHECK_TARGET_WD_BASE_NAME | Join-Path -ChildPath $DATETIME)
$CHECK_TARGET_WD_IN_FILE = $(Join-Path $CHECK_TARGET_WD_IN_DIR ($CHECK_TARGET_WD_BASE_NAME + "_" + $DATETIME + ".docx"))
$CHECK_TARGET_WD_OUT_DIR = $(Join-Path $CHECK_TOOL_OUT_DIR $CHECK_TARGET_WD_BASE_NAME | Join-Path -ChildPath $DATETIME)
$CHECK_TARGET_WD_OUT_FILE_PATTERN = $(Join-Path $CHECK_TARGET_WD_OUT_DIR ($CHECK_TARGET_WD_BASE_NAME + "_" + $DATETIME + "*.docx")) # 後続処理で使用

# inフォルダを作成
New-Item -ItemType Directory -Path $CHECK_TARGET_WD_IN_DIR

# outフォルダを作成
New-Item -ItemType Directory -Path $CHECK_TARGET_WD_OUT_DIR

# チェック用Wordファイルの作成
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    $doc = $word.Documents.Add()

    # 余白を狭く
    $doc.PageSetup.TopMargin = 10
    $doc.PageSetup.BottomMargin = 10
    $doc.PageSetup.LeftMargin = 10
    $doc.PageSetup.RightMargin = 10

    # 行間を最小限に
    $style = $doc.Styles.Item("標準")
    $style.ParagraphFormat.SpaceBefore = 0
    $style.ParagraphFormat.SpaceAfter = 0
    $style.ParagraphFormat.LineUnitBefore = 0
    $style.ParagraphFormat.LineUnitAfter = 0
    $style.ParagraphFormat.DisableLineHeightGrid = $true
    $style.ParagraphFormat.LineSpacingRule = 4 #4:wdLineSpaceExactly（固定値）
    $style.ParagraphFormat.LineSpacing = 20

    # フォントをVSCode初期設定に合わせる
    $style.Font.NameFarEast = "Meiryo UI"
    $style.Font.NameAscii = "Consolas"
    $style.Font.NameOther = "Consolas"
    $style.Font.Name = "Consolas"

    $firstFile = $true

    foreach ($info in $outputInfos) {
        $selection = $word.Selection
            
        # 最初のファイルでなければ改ページを挿入
        if (-not $firstFile) {
            $selection.InsertBreak(1)  # 1:wdPageBreak (ページ区切り)
        }
        $firstFile = $false
        
        # 見出し
        $selection.TypeText($info.Heading)
        $selection.Style = $doc.Styles.Item("見出し 1")
        $selection.TypeParagraph()

        # 内容
        $selection.TypeText($info.Content)
        $selection.TypeParagraph()
        $selection.TypeParagraph()
    }

    # 保存
    $CHECK_TARGET_WD_IN_FILE = $CHECK_TARGET_WD_IN_FILE -as [string] # refで渡すため、psobject -> string
    $doc.SaveAs([ref]$CHECK_TARGET_WD_IN_FILE, [ref]16) # 16:wdFormatDocumentDefault（docx）

}
finally {
    $doc.Close($false) # false:確認ダイアログを出さずに閉じる
    $word.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($doc) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
    Remove-Variable doc, word
}

#=======================================================================
# チェックツールの起動
#=======================================================================
# チェックツールの起動
Start-Process $CHECK_TOOL_EXE -Wait # -Wait:起動するまで待機

# アプリのウィンドウを取得
$window = $(Get-Window $APP_WINDOW_TITLE)
if ($null -eq $window) {
    Write-Error "アプリのウィンドウが取得できませんでした。指定されたウィンドウ名=[$APP_WINDOW_TITLE]"
    exit
}

# すべてのコントロールを取得
$controls = $(Get-AllControls $window)

# 各コントロールの連想配列を作成
$controlsByName = @{} # 名前がキー
$controlsByAutomationId = @{} # IDがキー

foreach ($control in $controls) {
    $name = $control.Current.Name
    $automationId = $control.Current.AutomationId

    # Name をキーにした連想配列
    if ($name -and !$controlsByName.ContainsKey($name)) {
        $controlsByName[$name] = $control
    }
    # AutomationId をキーにした連想配列
    if ($automationId -and !$controlsByAutomationId.ContainsKey($automationId)) {
        $controlsByAutomationId[$automationId] = $control
    }
}

#=======================================================================
# 各項目の入力
#=======================================================================
$name = "TextBox_Text"
Set-TextBoxValue $controlsByAutomationId[$name] $CHECK_TARGET_WD_OUT_DIR

# # inputフォルダ（テキストボックス）
# Set-TextBoxValue $controlsByAutomationId["xxx"] $CHECK_TARGET_WD_IN_DIR

# # 仕様（ドロップダウン）
# Select-ListBoxItem $controlsByName["xxx"] "せっけい"

# # チェック対象（チェックボックス）本文, 図形
# Select-CheckBox $controlsByName["本文"] $true
# Select-CheckBox $controlsByName["図形"] $true

# # outputフォルダ（テキストボックス）
# Set-TextBoxValue $controlsByAutomationId["xxx"] $CHECK_TARGET_WD_OUT_DIR

#=======================================================================
# チェック実行
#=======================================================================
# Click-Button $controlsByName["実行"]

# # ダイアログが2回表示されるので、OKまで進める
# Start-Sleep -Milliseconds 300
# [Windows.Forms.SendKeys]::SendWait("{ ENTER }")
# Start-Sleep -Milliseconds 300
# [Windows.Forms.SendKeys]::SendWait("{ ENTER }")

#=======================================================================
# チェック完了判定
#=======================================================================
# チェックツールの実行ボタンの活性/非活性を監視
# 実行開始時に実行ボタンが非活性になるため、活性化されたら、実行完了と判断する。

$timeout = 5
$elapsed = 0

do {
    $isEnabled = Get-IsEnabled $controlsByName["実行"]
    if ($isEnabled) {
        Write-Host "ボタン '$buttonName' は有効です"
        break
    }
    else {
        Write-Host "ボタン '$buttonName' は無効です"
    }

    Start-Sleep -Milliseconds 300
    $elapsed++

} while ($elapsed -lt $timeout)

# チェック結果Wordファイルの存在チェック
$outFile = $(Get-ChildItem -Path $CHECK_TARGET_WD_OUT_FILE_PATTERN)

if ( $outFile.Count -eq 0 ) {
    $wshell = New-Object -ComObject WScript.Shell
    $wshell.Popup("チェック結果Wordファイルが出力されていません。VsCodeのターミナルに出力されているメッセージを確認してください。", 0, "エラー", 16)
}
else {
    # チェック結果のWordファイルを開く
    Invoke-Item $outFile[0].FullName
}

#=======================================================================
# チェックツールを閉じる
#=======================================================================
Close-Window $window

Write-Host "処理終了"

# Add-Type -TypeDefinition @"
# using System;
# using System.Runtime.InteropServices;
# using System.Windows.Automation;

# public class UIAutomationHelper {
#     public static bool IsButtonEnabled(string buttonName) {
#         AutomationElement root = AutomationElement.RootElement;
#         if (root == null) return false;

#         Condition condition = new PropertyCondition(AutomationElement.NameProperty, buttonName);
#         AutomationElement button = root.FindFirst(TreeScope.Descendants, condition);

#         if (button == null) return false;

#         return !(bool)button.GetCurrentPropertyValue(AutomationElement.IsEnabledProperty);
#     }
# }
# "@ -Language CSharp

# # チェックしたいボタンの名前を指定
# $buttonName = "OK"

# # ボタンの状態を判定
# if ([UIAutomationHelper]::IsButtonEnabled($buttonName)) {
#     Write-Host "ボタン '$buttonName' は有効です"
# } else {
#     Write-Host "ボタン '$buttonName' は無効です"
# }


# function Get-MessageBox {
#     param(
#         [String]$MessageBoxTitle
#     )
#     # ルート要素を取得
#     $desktop = [System.Windows.Automation.AutomationElement]:: RootElement

#     # メッセージボックスを取得
#     $messageBox = $desktop.FindFirst([System.Windows.Automation.TreeScope]:: Children,
#         (New - Object System.Windows.Automation.PropertyCondition(
#             [System.Windows.Automation.AutomationElement]:: NameProperty, $MessageBoxTitle)))

#     return [System.Windows.Automation.AutomationElement]$messageBox
# }

# function Get- MessageBox - Buttons {
#     param(
#         [System.Windows.Automation.AutomationElement]$MessageBox
#     )

#     if ($MessageBox - ne $null) {
#         $buttons = $MessageBox.FindAll([System.Windows.Automation.TreeScope]:: Descendants,
#             (New - Object System.Windows.Automation.PropertyCondition(
#                 [System.Windows.Automation.AutomationElement]:: ControlTypeProperty,
#                 [System.Windows.Automation.ControlType]:: Button)))

#         return $buttons
#     } else {
#         Write - Host "メッセージボックスが見つかりません。"
#         return $null
#     }
# }

