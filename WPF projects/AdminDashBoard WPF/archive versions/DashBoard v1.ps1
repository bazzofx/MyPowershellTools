Add-Type -AssemblyName System.Windows.Forms

$FormObject = [System.Windows.Forms.Form]
$LabelObject = [System.Windows.Forms.Label]
$FileDialogObject = [System.Windows.Forms.OpenFileDialog]
$ButtonObject = [System.Windows.Forms.Button]
$RadioObject = [System.Windows.Forms.RadioButton]

$titleFont = "Cambria,11,style=bold"
$radioFont = "Verdana,9"
$fontBasic = "Verdana,10,style=bold"
$window = New-Object $FormObject

$window.ClientSize="350,450"
$window.Text="Sniper Action"
$window.backcolor = "lightgray"
    #Define Main Title of Program
$lbl_title = New-Object $LabelObject
$lbl_title.text = "User Admin Dashboard v1.0"
$lbl_title.Autosize = $true
$lbl_title.Font="Cambria,12,style=bold" ### --------------------------------------------font here
$lbl_title.Location = New-Object system.drawing.point (35,5)

    #Define FilePath Label info
$lbl_filePath = New-Object $LabelObject
$lbl_filePath.text = "File not selected yet"
$lbl_filePath.AutoEllipsis= $true
$lbl_filePath.AutoSize=$false
$lbl_filePath.BorderStyle="Fixed3D"
$lbl_filePath.width ="220"
$lbl_filePath.Height="20"
$lbl_filePath.Font="Verdana,8,style=bold"
$lbl_filePath.ForeColor="gray"
$lbl_filePath.Location = New-Object system.drawing.point (20,85)
    #Define Search Button
$btn_Search = New-Object $ButtonObject
$btn_Search.text="Search..."
$btn_Search.AutoSize=$true
$btn_Search.Location = New-Object System.Drawing.Point(250,80)

    #Define Onboarding Label
$lbl_Onboarding = New-Object $LabelObject
$lbl_Onboarding.text = "_____________Onboarding_____________"
$lbl_Onboarding.Autosize = $true
$lbl_Onboarding.Font=$titleFont ### --------------------------------------------font here
$lbl_Onboarding.Location = New-Object system.drawing.point (20,120)

    #Define Radio options for Onboarding
$rdn_Button = New-Object $RadioObject
$rdn_Button.Text ="Reset Password"
$rdn_Button.Font=$radioFont
$rdn_Button.Location= New-Object System.Drawing.Point(20,150)
$rdn_Button.AutoSize=$true
$rdn_Button.Parent="onboarding"

$rdn_Button2 = New-Object $RadioObject
$rdn_Button2.Text ="Unlock Account"
$rdn_Button2.Font=$radioFont
$rdn_Button2.Location= New-Object System.Drawing.Point(20,175)
$rdn_Button2.AutoSize=$true
$rdn_Button2.Parent="onboarding"

$rdn_Button3 = New-Object $RadioObject
$rdn_Button3.Text ="Add MS License"
$rdn_Button3.Font=$radioFont
$rdn_Button3.Location= New-Object System.Drawing.Point(20,200)
$rdn_Button3.AutoSize=$true
$rdn_Button3.Parent="onboarding"
    
    #Define Onboarding Label
$lbl_Offboarding = New-Object $LabelObject
$lbl_Offboarding.text = "_____________Offboarding_____________"
$lbl_Offboarding.Autosize = $true
$lbl_Offboarding.Font=$titleFont ### --------------------------------------------font here
$lbl_Offboarding.Location = New-Object system.drawing.point (20,230)

    #Define Radio options for Offboarding
$rdn_Button4 = New-Object $RadioObject
$rdn_Button4.Text ="Lock Account"
$rdn_Button4.Font=$radioFont
$rdn_Button4.Location= New-Object System.Drawing.Point(20,260)
$rdn_Button4.AutoSize=$true
$rdn_Button4.Parent="offboarding"

$rdn_Button5 = New-Object $RadioObject
$rdn_Button5.Text ="Remove All Licenses"
$rdn_Button5.Font=$radioFont
$rdn_Button5.Location= New-Object System.Drawing.Point(20,285)
$rdn_Button5.AutoSize=$true
$rdn_Button5.Parent="offboarding"

$btn_Login = New-Object $ButtonObject
$btn_Login.text="Login"
$btn_Login.AutoSize=$true
$btn_Login.Font="Arial,9,style=bold"
$btn_Login.ForeColor="green"
$btn_Login.Location = New-Object System.Drawing.Point(250,40)

$btn_exec = New-Object $ButtonObject
$btn_exec.text="Execute"
$btn_exec.Font="Arial,13,style=bold"
$btn_exec.ForeColor="green"
$btn_exec.AutoSize=$true
$btn_exec.Location = New-Object System.Drawing.Point(20,400)

$btn_cancel = New-Object $ButtonObject
$btn_cancel.text="Cancel"
$btn_cancel.Font="Arial,11,style=bold"
$btn_cancel.ForeColor="red"
$btn_cancel.AutoSize=$true
$btn_cancel.Location = New-Object System.Drawing.Point(250,400)
#Functions
function SearchDialog{
        $OpenFileDialog = New-Object $FileDialogObject
        $OpenFileDialog.InitialDirectory = $Path
        $OpenFileDialog.CheckPathExists = $true
        $OpenFileDialog.CheckFileExists = $true
        $OpenFileDialog.Title = "Please feed me"
        $OpenFileDialog.Filter = "csv files (*.csv)|*.csv" #|All files (*.*)|*.*"

        $OpenFileDialog.ShowDialog()

            if ($openFileDialog.CheckFileExists -eq $true) {
                $filePath = $openFileDialog.FileName
                Write-Host $filePath }
        $lbl_filePath.Text = $filePath
        $lbl_filePath.ForeColor="green"
}


##Add Functions to the Form
$btn_Search.Add_Click({SearchDialog})

$window.Controls.AddRange(@($lbl_title,$btn_Search,$lbl_filePath,$lbl_Onboarding,$rdn_Button,$rdn_Button2,$rdn_Button3,$lbl_Offboarding,
$rdn_Button4,$rdn_Button5,$btn_Login,$btn_exec,$btn_cancel))

#creates form
$window.ShowDialog()
#cleans up
$window.Dispose()