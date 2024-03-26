# 加载Excel组件
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

# 创建Excel应用程序对象

$excel = New-Object -ComObject Excel.Application
# 隐藏Excel窗口
$excel.Visible = $false

# 打开工作簿
(Get-Location).Path
$workbook = $excel.Workbooks.Open("C:\Users\sakura\Desktop\ps\options.xlsx")

# 获取第一个工作表
$worksheet = $workbook.Sheets.Item(1)

# 获取A列和C列的最大行数
$countRowA = ($worksheet.UsedRange.Columns.Item('A').Cells | Where-Object { $_.Value2 -ne $null }).Count
$countRowE = ($worksheet.UsedRange.Columns.Item('E').Cells | Where-Object { $_.Value2 -ne $null }).Count
$lengthA = $countRowA | ForEach-Object {if ($_ % 2 -eq 0) {$_} else {$_ + 1}}
$lengthE = $countRowE | ForEach-Object {if ($_ % 2 -eq 0) {$_} else {$_ + 1}}

$quotient = [math]::floor(100 / $lengthA)
$remainder = 100 % $lengthA
$arrayA = @()
for ($i = 0; $i -lt $lengthA; $i++) {
    $value = if ($i -le $remainder) { $quotient + 1 } else { $quotient }
    $arrayA += $value*2
}
#$arrayA

$quotient = [math]::floor(100 / $lengthE)
$remainder = 100 % $lengthE
$arrayE = @()
for ($i = 0; $i -lt $lengthE; $i++) {
    $value = if ($i -le $remainder) { $quotient + 1 } else { $quotient }
    $arrayE += $value*2
}
#$arrayE

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# 创建窗口
$form = New-Object System.Windows.Forms.Form
$form.Text = "Dynamic GUI"
$form.Size = New-Object System.Drawing.Size(600,400)

# 创建 TableLayoutPanel
$tableLayoutPanel = New-Object System.Windows.Forms.TableLayoutPanel
$tableLayoutPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayoutPanel.ColumnCount = 2
$tableLayoutPanel.RowCount = 2
$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$tableLayoutPanel.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 80)))
$tableLayoutPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 20)))
$form.Controls.Add($tableLayoutPanel)

# 创建左上组的 RadioButtons
$groupBox1 = New-Object System.Windows.Forms.GroupBox
$groupBox1.Text = "Group 1"
$groupBox1.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayoutPanel.Controls.Add($groupBox1, 0, 0)

# 创建 TableLayoutPanel 用于放置 GroupBox1 内的 RadioButton
$radioTable1 = New-Object System.Windows.Forms.TableLayoutPanel
$radioTable1.Dock = [System.Windows.Forms.DockStyle]::Fill
$radioTable1.ColumnCount = 2
$radioTable1.RowCount = $lengthA/2
$radioTable1.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$radioTable1.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
for ($i = 1; $i -le $lengthA; $i++) {
    $radioTable1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, $arrayA[($i)-1])))
}
# $radioTable1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
# $radioTable1.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$groupBox1.Controls.Add($radioTable1)

for ($i = 1; $i -le $countRowA; $i++) {
    $radioButton = New-Object System.Windows.Forms.RadioButton
    $radioButton.Text = $worksheet.Cells.Item($i, 'A').Value2
    $radioButton.AutoSize = $true
    $radioTable1.Controls.Add($radioButton)
}

# 创建右上组的 RadioButtons
$groupBox2 = New-Object System.Windows.Forms.GroupBox
$groupBox2.Text = "Group 2"
$groupBox2.Dock = [System.Windows.Forms.DockStyle]::Fill
$tableLayoutPanel.Controls.Add($groupBox2, 1, 0)

# 创建 TableLayoutPanel 用于放置 GroupBox2 内的 RadioButton
$radioTable2 = New-Object System.Windows.Forms.TableLayoutPanel
$radioTable2.Dock = [System.Windows.Forms.DockStyle]::Fill
$radioTable2.ColumnCount = 2
$radioTable2.RowCount = $lengthE/2
$radioTable2.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$radioTable2.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
for ($i = 1; $i -le $lengthE; $i++) {
    $radioTable2.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, $arrayE[($i)-1])))
}
# $radioTable2.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
# $radioTable2.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$groupBox2.Controls.Add($radioTable2)

for ($i = 1; $i -le $countRowE; $i++) {
    $radioButton = New-Object System.Windows.Forms.RadioButton
    $radioButton.Text = $worksheet.Cells.Item($i, 'E').Value2
    $radioButton.AutoSize = $true
    $radioTable2.Controls.Add($radioButton)
}

# 创建按钮
$button = New-Object System.Windows.Forms.Button
$button.Text = "Click me!"
$button.MaximumSize = New-Object System.Drawing.Size(100, 50)  # 设置最大尺寸
$tableLayoutPanel.Controls.Add($button, 0, 1)
$tableLayoutPanel.SetColumnSpan($button, 2)
# 设置按钮居中
$button.Anchor = [System.Windows.Forms.AnchorStyles]::None
$button.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

$button.Add_Click({
    $selectedLastName = $radioTable1.Controls | Where-Object { $_.Checked }
    $selectedFirstName = $radioTable2.Controls | Where-Object { $_.Checked }
    $firstName = $selectedFirstName.Text
    [System.Windows.Forms.MessageBox]::Show("cpoyed$firstName","hint")
    Set-Clipboard -Value $firstName
    Start-Process notepad.exe -ArgumentList "$($selectedLastName.Text) $($selectedFirstName.Text)"
    $form.Close()
})

# 关闭工作簿
$workbook.Close()

# 退出Excel应用程序
$excel.Quit()

# 释放COM对象
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

# 显示窗口
$form.ShowDialog() | Out-Null

