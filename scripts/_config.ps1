#==============================================================================
# チェック対象ファイル設定
#==============================================================================
# 検索対象のルートディレクトリ（このディレクトリ配下を再帰検索）
# ※各自ローカルのパスに変更
$global:CHECK_TARGET_ROOT_DIR = "C:\Users\roron431\Music\my-extension\bk"

# 以下4つは、大文字小文字は区別されない。また、ワイルドカード(*)が使用可能。
# (*) https://learn.microsoft.com/ja-jp/powershell/module/microsoft.powershell.core/about/about_wildcards?view=powershell-7.5

# 検索対象の抽出条件（フォルダ名）
$global:CHECK_TARGET_INCLUDE_DIR_NAMES = @()

# 検索対象の除外条件（フォルダ名）
$global:CHECK_TARGET_EXCLUDE_DIR_NAMES = @("*bk*", "*backup*", "*old*", "*backup*")

# 検索対象の抽出条件（ファイル名）
$global:CHECK_TARGET_INCLUDE_FILE_NAMES = @("*.adoc")

# 検索対象の除外条件（ファイル名）
$global:CHECK_TARGET_EXCLUDE_FILE_NAMES = @()

#==============================================================================
# チェックツール設定
#==============================================================================
# 使用するチェックツールの実行ファイル
# ※各自ローカルのパスに変更
$global:CHECK_TOOL_EXE = "C:\Users\roron431\Downloads\FileToText_20130623\FileToText.exe"

# チェックツールのディレクトリ情報（自動設定）
# 実行ファイルのディレクトリ
$global:CHECK_TOOL_DIR = $(Split-Path -Path $CHECK_TOOL_EXE)
# 入力ファイル用ディレクトリ  
$global:CHECK_TOOL_IN_DIR = $(Join-Path $CHECK_TOOL_DIR "in")
# 出力ファイル用ディレクトリ
$global:CHECK_TOOL_OUT_DIR = $(Join-Path $CHECK_TOOL_DIR "out")

# チェックツールのアプリケーションウィンドウのタイトル
$global:APP_WINDOW_TITLE = "*FileToText*"

#==============================================================================
# チェック用Wordファイル設定
#==============================================================================
# outフォルダ配下に出力する際のルートフォルダ名および出力ファイル名に利用する名前
$global:CHECK_TARGET_WD_BASE_NAME = "forVSCode"
