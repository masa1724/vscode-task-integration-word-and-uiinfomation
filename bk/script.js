
var sh = new ActiveXObject("WScript.Shell");
sh.Run("calc.exe");

// 起動してすぐにキー送信すると失敗する可能性があるので、3秒停止
WScript.Sleep(2000);

//  A～Zキーを0.1秒ごとに送信
sh.SendKeys("1");

WScript.Sleep(100);

sh.SendKeys("2");

WScript.Sleep(100);

sh.SendKeys("3");

// https://learn.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/sendkeys-statement

WScript.Sleep(200);
sh.SendKeys("{TAB}");
WScript.Sleep(200);
sh.SendKeys("{TAB}");
WScript.Sleep(200);
sh.SendKeys("{TAB}");
WScript.Sleep(200);
sh.SendKeys("{TAB}");
WScript.Sleep(200);
sh.SendKeys("{TAB}");
WScript.Sleep(200);
sh.SendKeys("{TAB}");





function Get-MessageBox {
    param(
        [String]$MessageBoxTitle
    )

    # ルート要素を取得
    $desktop = [System.Windows.Automation.AutomationElement]:: RootElement

    # メッセージボックスを取得
    $messageBox = $desktop.FindFirst([System.Windows.Automation.TreeScope]:: Children,
        (New - Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]:: NameProperty, $MessageBoxTitle)))

    return [System.Windows.Automation.AutomationElement]$messageBox
}

-------

    function Get- MessageBox - Buttons {
    param(
        [System.Windows.Automation.AutomationElement]$MessageBox
    )

    if ($MessageBox - ne $null) {
        $buttons = $MessageBox.FindAll([System.Windows.Automation.TreeScope]:: Descendants,
            (New - Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]:: ControlTypeProperty,
                [System.Windows.Automation.ControlType]:: Button)))

        return $buttons
    } else {
        Write - Host "メッセージボックスが見つかりません。"
        return $null
    }
}

