. "$PSScriptRoot\_automation-utils.ps1"

# アプリを起動
# calc
# $WINDOW_NAME = "電卓";

# Start-Process "C:\Users\roron431\Downloads\df141\DF.exe"
# $WINDOW_NAME = "DF";

Start-Process "C:\Users\roron431\Downloads\TresGrep_1.23_20210604\TresGrep\TresGrep.exe"
$WINDOW_NAME = "TresGrep  ファイル・テキスト検索";

# Start-Process "c:\Users\roron431\Downloads\FileToText_20130623\FileToText.exe"
# $WINDOW_NAME = "FileToText*";

Start-Sleep -Seconds 2  # 少し待ってウィンドウが開くのを待つ

# アプリのウィンドウを取得
$window = $(Get-Window $WINDOW_NAME)
if ($null -eq $window) {
    Write-Error "アプリのウィンドウが取得できませんでした。指定されたウィンドウ名=[$WINDOW_NAME]"
    exit
}

# すべてのコントロールを取得
$controls = $(Get-AllControls $window)

# 各コントロールの連想配列を作成（後続処理でキー指定でコントロールを参照できるように。）
$controlsByName = @{} # 名前がキー
$controlsByAutomationId = @{} # IDがキー（通常は使わない）

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

$SENDMAIL_INI_FILE = "${PJ_DIR}/tools/sendmail/sendmail.ini"
$content = Get-Content -Path $SENDMAIL_INI_FILE -Encoding UTF8 -Raw

＜１　ファイルリスト収集＞
# 1. チェック対象のファイルリストを収集

＜２　チェック用一時Wordファイルの作成＞
# 2. 1の各ファイルの内容をメモリに読み込み
# 3. inフォルダを作成
　　　「チェックツール.exe/in/forVsCode_YYYYMMDDHHMMSS/」
# 4. outフォルダを作成
　　　「チェックツール.exe/out/forVsCode_YYYYMMDDHHMMSS/」
# 5. 4配下にWordファイルを作成
　　　「adoc_YYYYMMDDHHMMSS.word」
# 6. 5にファイル名を見出しに、ファイル内容を行番号付きで書き出し

＜３　チェックツールの実行＞
# 7. アプリを起動
# 8. 各入力項目を設定し、実行ボタンを押下
# 9. ダイアログが2回表示されるので、OKまで進める

＜４　チェックツールの実行完了の判定＞
# 10. (悩み中)
　　- ファイルのサイズ増加を監視して、サイズ増加が止まったら出力完了とみなす
　　- チェックツールの実行ボタンの活性/非活性を監視


# inputフォルダ（テキストボックス）
# 仕様（ドロップダウン）
# チェック対象（チェックボックス）本文, 図形
# outputフォルダ（テキストボックス）


# $name = "TextBox_Text"
# Set-TextBoxValue $controlsByAutomationId[$name] "das,dlsal;kdsdsadsadsaa"

# Start-Sleep -Milliseconds 500
# Select-CheckBox $controlsByName[$name] $true
# Write-Host "-----------------------------------------"
# Write-Host $(Get-CheckBoxIsSelected $controlsByName[$name])
# Write-Host "-----------------------------------------"

# Start-Sleep -Milliseconds 500
# Select-CheckBox $controlsByName[$name] $false
# Write-Host "-----------------------------------------"
# Write-Host $(Get-CheckBoxIsSelected $controlsByName[$name])
# Write-Host "-----------------------------------------"

# Start-Sleep -Milliseconds 500
# Select-CheckBox $controlsByName[$name] $true
# Write-Host "-----------------------------------------"
# Write-Host $(Get-CheckBoxIsSelected $controlsByName[$name])
# Write-Host "-----------------------------------------"

# Start-Sleep -Milliseconds 500
# Select-CheckBox $controlsByName[$name] $false
# Write-Host "-----------------------------------------"
# Write-Host $(Get-CheckBoxIsSelected $controlsByName[$name])
# Write-Host "-----------------------------------------"

Set-TextBoxValue $controlsByAutomationId["1001"] "dasdas421312grfgr"

# Start-Sleep -Milliseconds 1000
# # Select-CheckBox $controlsByName["リアルタイム検索(8)"] $true
# Select-CheckBox-Multiple $controlsByName["リアルタイム検索(8)"] @("リアルタイム検索(8)")
# Write-Host "-----------------------------------------"
# Write-Host $(Get-CheckBoxIsSelected $controlsByName["リアルタイム検索(8)"])
# Write-Host "-----------------------------------------"

# Start-Sleep -Milliseconds 1000
# # Select-CheckBox $controlsByName["リアルタイム検索(8)"] $false
# Select-CheckBox-Multiple $controlsByName["リアルタイム検索(8)"] @()
# Write-Host "-----------------------------------------"
# Write-Host $(Get-CheckBoxIsSelected $controlsByName["リアルタイム検索(8)"])
# Write-Host "-----------------------------------------"

# Start-Sleep -Milliseconds 1000
# # Select-CheckBox $controlsByName["リアルタイム検索(8)"] $true
# Select-CheckBox-Multiple $controlsByName["リアルタイム検索(8)"] @("リアルタイム検索(8)")
# Write-Host "-----------------------------------------"
# Write-Host $(Get-CheckBoxIsSelected $controlsByName["リアルタイム検索(8)"])
# Write-Host "-----------------------------------------"

# Click-Button $controlsByName["1"]
# Click-Button $controlsByName["2"]
# Click-Button $controlsByName["3"]
# Click-Button $controlsByName["4"]




# # アプリのウィンドウを閉じる
# Close-Window $window
