Add-Type -AssemblyName UIAutomationClient

# アプリのウィンドウを取得
function Get-Window {
    [OutputType([System.Windows.Automation.AutomationElement])]
    param (
        [String]$WindowNamePattern
    )
    # デスクトップのルート要素を取得
    $desktop = [System.Windows.Automation.AutomationElement]::RootElement

    $timeout = 5
    $elapsed = 0

    $windows = @()
    $window = $null
    do {
        # アプリのウィンドウを取得
        $windows = $desktop.FindAll([System.Windows.Automation.TreeScope]::Children,
            (New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::ControlTypeProperty, 
                [System.Windows.Automation.ControlType]::Window)))

        foreach ($w in $windows) {
            if ($w.Current.Name -like $WindowNamePattern) {
                $window = $w
                break
            }
        }
        if ($null -ne $window) {
            break
        }

        Start-Sleep -Milliseconds 300
        $elapsed++

    } while ($elapsed -lt $timeout)
    
    Write-Host "--- Found Windows --------------------------"
    foreach ($w in $windows) {
        $name = $w.Current.Name
        $automationId = $w.Current.AutomationId
        Write-Host "[Window] Name: $name, AutomationId: $automationId"
    }
    Write-Host "--------------------------------------------"

    return $window
}

# アプリのウィンドウを閉じる
function Close-Window {
    [OutputType([void])]
    param (
        [System.Windows.Automation.AutomationElement]$Window
    )
    $closePattern = $Window.GetCurrentPattern([System.Windows.Automation.WindowPattern]::Pattern)
    $closePattern.Close()
}

# すべてのコントロールの取得
function Get-AllControls {
    [OutputType([System.Windows.Automation.AutomationElement[]])]
    param (
        [System.Windows.Automation.AutomationElement]$Window
    )
    $controls = $Window.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)

    Write-Host "--- Found Controls -------------------------"
    foreach ($control in $controls) {
        $name = $control.Current.Name
        $automationId = $control.Current.AutomationId
        $controlType = $control.Current.ControlType.ProgrammaticName

        Write-Host "[Control] Name: $name, AutomationId: $automationId, ControlType: $controlType"
        $supportedPatterns = $control.GetSupportedPatterns()
        if ($supportedPatterns.Count -gt 0) {
            foreach ($pattern in $supportedPatterns) {
                $patternName = $pattern.ProgrammaticName
                Write-Host "  Supported Pattern: $patternName"
            }
        }
        else {
            Write-Host "  No supported patterns."
        }   
    }
    Write-Host "--------------------------------------------"
    return $controls
}

# Button のクリック
function Click-Button {
    [OutputType([void])]
    param (
        [System.Windows.Automation.AutomationElement]$Button
    )
    $invokePattern = $Button.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)
    $invokePattern.Invoke()
}

# TextBox の値を取得
function Get-TextBoxValue {
    [OutputType([string])]
    param (
        [System.Windows.Automation.AutomationElement]$TextBox
    )
    $valuePattern = $null
    if ($TextBox.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)) {
        # 1行のテキストボックス
        return $valuePattern.Current.Value
    }
    else {
        # 複数行のテキストボックス
        $textPattern = $TextBox.GetCurrentPattern([System.Windows.Automation.TextPattern]::Pattern)
        return $textPattern.DocumentRange.GetText(-1) # -1:文字数制限なし(＝全ての文字列を取得)
    }
}

Add-Type -AssemblyName System.Windows.Forms

# TextBox に値を設定
function Set-TextBoxValue {
    [OutputType([void])]
    param (
        [System.Windows.Automation.AutomationElement]$TextBox,
        [string]$Value
    )
    $valuePattern = $null
    if ($TextBox.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)) {
        # 1行のテキストボックス
        $valuePattern.SetValue($Value)
    }
    else {
        # 複数行のテキストボックス
        # https://learn.microsoft.com/ja-jp/dotnet/framework/ui-automation/add-content-to-a-text-box-using-ui-automation
        $TextBox.SetFocus();
        Start-Sleep -Milliseconds 100
        
        # 既に入力されている内容をクリア
        [Windows.Forms.SendKeys]::SendWait("^{HOME}") # Ctrl + Home
        [Windows.Forms.SendKeys]::SendWait("^+{END}") # Ctrl + Shift + Home
        [Windows.Forms.SendKeys]::SendWait("{DEL}")   # Delete

        # ↓ クリップボード経由で貼り付け（早い）
        [System.Windows.Forms.Clipboard]::SetText($Value)
        [Windows.Forms.SendKeys]::SendWait("^v") # Ctrl + V

        # ↓ 遅い
        # [Windows.Forms.SendKeys]::SendWait($Value)
    }
}

# ListBox / CombComboBox の選択中アイテムのテキストを取得
function Get-ListBoxItemText {
    [OutputType([string])]
    param (
        [System.Windows.Automation.AutomationElement]$ListBox
    )
    $selectionPattern = $ListBox.GetCurrentPattern([System.Windows.Automation.SelectionPattern]::Pattern)
    return $selectionPattern.Current.GetSelection()[0].Current.Name
}

# ListBox / CombComboBox のアイテムを選択
function Select-ListBoxItem {
    [OutputType([void])]
    param (
        [System.Windows.Automation.AutomationElement]$ListBox,
        [string]$SelectedItemText
    )
    # $expandCollapsePattern = $ListBox.GetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern)
    # $expandCollapsePattern.Expand()
    # Start-Sleep -Milliseconds 500  # 一時停止（展開を待つ）

    # すべての ドロップダウンのアイテムを取得
    $items = __getSubtreeByType $ListBox ([System.Windows.Automation.ControlType]::ListItem)

    # 指定したアイテムを選択
    foreach ($item in $items) {
        if ($item.Current.Name -eq $SelectedItemText) {
            $selectionPattern = $item.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
            $selectionPattern.Select()
            break
        }
    }
    # $expandCollapsePattern.Collapse()
}

# CheckBox の現在の選択状態(true/false)を取得
# (1つのCheckBox のコントロールを指定する場合)
function Get-CheckBoxIsSelected {
    [OutputType([bool])]
    param (
        [System.Windows.Automation.AutomationElement]$CheckBox
    )
    $togglePattern = $CheckBox.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern)
    $currentState = $togglePattern.Current.ToggleState
    return ($currentState -eq [System.Windows.Automation.ToggleState]::On)
}

# CheckBox の選択状態(true/false)を変更
# (1つのCheckBox のコントロールを指定する場合)
function Select-CheckBox {
    [OutputType([void])]
    param (
        [System.Windows.Automation.AutomationElement]$CheckBox,
        [bool]$Selected
    )
    $togglePattern = $CheckBox.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern)
    $currentState = $togglePattern.Current.ToggleState
    if ($Selected -and ($currentState -ne [System.Windows.Automation.ToggleState]::On)) {
        $togglePattern.Toggle()
    }
    elseif (-not $Selected -and ($currentState -ne [System.Windows.Automation.ToggleState]::Off)) {
        $togglePattern.Toggle()
    }
}

# CheckBox の現在の選択状態(選択中アイテムのテキストのリスト)を取得
# (複数のCheckBox を包含する 親コントロールを指定する場合)
function Get-CheckBoxSelectionTexts {
    [OutputType([string[]])]
    param (
        [System.Windows.Automation.AutomationElement]$ParentElement
    )
    $checkBoxes = __getSubtreeByType $ParentElement ([System.Windows.Automation.ControlType]::CheckBox)
    $selectedItemTexts = @()

    foreach ($checkBox in $checkBoxes) {
        $togglePattern = $checkBox.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern)
        if ($togglePattern.Current.ToggleState -eq [System.Windows.Automation.ToggleState]::On) {
            $selectedItemTexts += $checkBox.Current.Name
        }
    }

    return $selectedItemTexts
}

# CheckBox の選択状態を変更（指定テキストを選択し、それ以外を解除）
# (複数のCheckBox を包含する 親コントロールを指定する場合)
function Select-CheckBox-Multiple {
    [OutputType([void])]
    param (
        [System.Windows.Automation.AutomationElement]$ParentElement,
        [string[]]$SelectedItemTexts
    )
    $checkBoxes = __getSubtreeByType $ParentElement ([System.Windows.Automation.ControlType]::CheckBox)

    foreach ($checkBox in $checkBoxes) {
        $togglePattern = $checkBox.GetCurrentPattern([System.Windows.Automation.TogglePattern]::Pattern)
        $currentState = $togglePattern.Current.ToggleState
        $shouldBeChecked = $checkBox.Current.Name -in $SelectedItemTexts
        
        # チェックボックスの状態を変更
        if ($shouldBeChecked -and ($currentState -ne [System.Windows.Automation.ToggleState]::On)) {
            $togglePattern.Toggle()
        }
        elseif (-not $shouldBeChecked -and ($currentState -ne [System.Windows.Automation.ToggleState]::Off)) {
            $togglePattern.Toggle()
        }
    }
}

# RadioButton の現在の選択状態(true/false)を取得
# (1つのRadioButton のコントロールを指定する場合)
function Get-RadioButtonIsSelected {
    [OutputType([bool])]
    param (
        [System.Windows.Automation.AutomationElement]$RadioButton
    )
    $selectionPattern = $RadioButton.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
    return $selectionPattern.Current.IsSelected
}

# RadioButton を選択
# (1つのRadioButton のコントロールを指定する場合)
function Select-RadioButton {
    [OutputType([void])]
    param (
        [System.Windows.Automation.AutomationElement]$RadioButton
    )
    $selectionPattern = $RadioButton.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
    $selectionPattern.Select()
}

# RadioButton の選択中アイテムの名前を取得
# (複数のRadioButton を包含する 親コントロールを指定する場合)
function Get-RadioButtonSelectedItemText {
    [OutputType([string])]
    param (
        [System.Windows.Automation.AutomationElement]$ParentElement
    )
    $radioButtons = __getSubtreeByType $ParentElement ([System.Windows.Automation.ControlType]::RadioButton)

    foreach ($radioButton in $radioButtons) {
        $selectionPattern = $radioButton.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
        if ($selectionPattern.Current.IsSelected) {
            return $radioButton.Current.Name
        }
    }
    return $null
}

# RadioButton の選択状態を変更
# (複数のRadioButton を包含する 親コントロールを指定する場合)
function Select-RadioButton-Multiple {
    [OutputType([void])]
    param (
        [System.Windows.Automation.AutomationElement]$ParentElement,
        [string]$SelectedItemText
    )
    $radioButtons = __getSubtreeByType $ParentElement ([System.Windows.Automation.ControlType]::RadioButton)

    foreach ($radioButton in $radioButtons) {
        if ($radioButton.Current.Name -eq $SelectedItemText) {
            $selectionPattern = $radioButton.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern)
            $selectionPattern.Select()
            break
        }
    }
}

# コントロールの活性状態(true/false)を取得
function Get-IsEnabled {
    [OutputType([bool])]
    param (
        [System.Windows.Automation.AutomationElement]$Element
    )
    return $Element.GetCurrentPropertyValue([System.Windows.Automation.AutomationElement]::IsEnabledProperty)
}

function __getSubtreeByType {
    param (
        [System.Windows.Automation.AutomationElement]$ParentElement,
        [System.Windows.Automation.ControlType]$ControlType
    )
    return $ParentElement.FindAll([System.Windows.Automation.TreeScope]::Subtree, 
        (New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty, 
            $ControlType))
    )
}
