class ENavi
{
    $ie
    $doc

    # 初期化
    ENavi($URL)
    {
        $this.ie = New-Object -ComObject InternetExplorer.Application
        $this.ie.Visible = $true
        $this.ie.Navigate($URL)
        while ($this.ie.busy -or $this.ie.readystate -ne 4) { Start-Sleep -Milliseconds 100 }
        $this.doc = $this.ie.document
    }

    # ログイン画面
    Login($USER_ID,$USER_PASSWORD)
    {
        $this.doc.IHTMLDocument3_getElementById("TextStaffNo").value = $USER_ID
        $this.doc.IHTMLDocument3_getElementById("TextPassword").value = $USER_PASSWORD
        $this.doc.IHTMLDocument3_getElementById("BtnOk").click()
        while ($this.ie.busy -or $this.ie.readystate -ne 4) { Start-Sleep -Milliseconds 100 }
    }

    # トップ画面
    Top()
    {
        while ($this.ie.busy -or $this.ie.readystate -ne 4) { Start-Sleep -Milliseconds 100 }
        $this.doc.IHTMLDocument3_getElementById("ImgBtnMenuMonth").click()
        while ($this.ie.busy -or $this.ie.readystate -ne 4) { Start-Sleep -Milliseconds 100 }
    }

    # 月次勤怠画面
    Attendance()
    {
        $today_btn = "LinkBtnDate" + [string](Get-Date).AddDays(-1).ToString("dd")
        $this.doc.IHTMLDocument3_getElementById($today_btn).click()
        while ($this.ie.busy -or $this.ie.readystate -ne 4) { Start-Sleep -Milliseconds 100 }
    }

    # 勤怠入力画面
    InputScreen($ATTENDANCE_STATUS, $BEGIN_TIME_HOUR, $BEGIN_TIME_MIN, $END_TIME_HOUR, $END_TIME_MIN, $COMMENT)
    {
        $this.doc.IHTMLDocument3_getElementById("CmbStatus").SelectedIndex = $ATTENDANCE_STATUS
        $this.doc.IHTMLDocument3_getElementById("CmbBeginTimeHour").SelectedIndex = $BEGIN_TIME_HOUR
        $this.doc.IHTMLDocument3_getElementById("CmbBeginTimeMin").SelectedIndex = $BEGIN_TIME_MIN
        $this.doc.IHTMLDocument3_getElementById("CmbEndTimeHour").SelectedIndex = $END_TIME_HOUR
        $this.doc.IHTMLDocument3_getElementById("CmbEndTimeMin").SelectedIndex = $END_TIME_MIN
        $this.doc.IHTMLDocument3_getElementById("TextComment").value = $COMMENT
        $this.doc.IHTMLDocument3_getElementById("BtnOkSigndayedit").click()
        while ($this.ie.busy -or $this.ie.readystate -ne 4) { Start-Sleep -Milliseconds 100 }
    }

    # 確認画面
    Confirm()
    {
        $this.doc.IHTMLDocument3_getElementById("BtnOk").click()
        sleep 30
        $this.ie.Quit()
    }
}

$INPUT_CSV_PATH
Add-Type -AssemblyName System.Windows.Forms
# 入力ファイル選択
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = 'C:\Users\AWVZXS010\Documents\chromedriver_win32\e-navi' 
    Filter = 'すべてのファイル|*'
    Title = 'select file'
}
if($FileBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
    $INPUT_CSV_PATH = $FileBrowser.FileName
}

# configファイル読込
$lines = get-content $INPUT_CSV_PATH
$data = @()
foreach($line in $lines){
    if($line -match "^$"){ continue }
    if($line -match "^\s*;"){ continue }

    $param = $line.split("=",2)
    write-host($param[1])
    $data += $param[1]
}

# 定数定義
$URL = $data[0]
$USER_ID = $data[1]
$USER_PASSWORD = $data[2]
$ATTENDANCE_STATUS = $data[3]
$BEGIN_TIME_HOUR = $data[4]
$BEGIN_TIME_MIN = $data[5]
$END_TIME_HOUR = $data[6]
$END_TIME_MIN = $data[7]
$COMMENT = $data[8]

# 主処理
$enavi = New-Object ENavi($URL)
$enavi.Login($USER_ID,$USER_PASSWORD)
$enavi.Top()
$enavi.Attendance()
$enavi.InputScreen($ATTENDANCE_STATUS,$BEGIN_TIME_HOUR,$BEGIN_TIME_MIN,$END_TIME_HOUR,$END_TIME_MIN,$COMMENT)
$enavi.Confirm()
