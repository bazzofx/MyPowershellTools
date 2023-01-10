Add-Type -AssemblyName System.Windows.Forms

$FormObject = [System.Windows.Forms.Form]
$LabelObject =[System.Windows.Forms.Label]
$ComboBoxObject = [System.Windows.Forms.ComboBox]
$RadioObject = [System.Windows.Forms.RadioButton]
$ButtonObject = [System.Windows.Forms.Button]
$DefaultFont = "Verdana,10"

#Set up base form
$AppForm = New-Object $FormObject
$AppForm.ClientSize="500,300"
$AppForm.Text = "Programmer - Service Inspector"
$AppForm.BackColor = "lightblue"
$AppForm.Font=$DefaultFont

#Build the form
$lblService = New-Object $LabelObject
$lblService.Text="Services :"
$lblService.Autosize=$true
$lblService.Location=New-Object System.Drawing.Point(20,20)

$ddlService = New-object $ComboBoxObject
$ddlService.Width="300"
$ddlService.Location=New-Object System.Drawing.Point(120,20)
$ddlService.Text = "Pick a service"
#Load the dropdown list for services
Get-Service | ForEach-Object {$ddlService.Items.Add($_.Name)}


$lblForName = New-Object $LabelObject
$lblForName.Text="Service Friendly Name:"
$lblForName.Autosize=$true
$lblForName.Location=New-Object System.Drawing.Point(20,120)

$lblForStatus = New-Object $LabelObject
$lblForStatus.Text="Label Status:"
$lblForStatus.Autosize=$true
$lblForStatus.Location=New-Object System.Drawing.Point(20,220)

$lblStatus = New-Object $LabelObject
$lblStatus.Text=""
$lblStatus.Autosize=$true
$lblStatus.Location=New-Object System.Drawing.Point(160,220)

$lblName = New-Object $LabelObject
$lblName.Text=""
$lblName.Autosize=$true
$lblName.Location=New-Object System.Drawing.Point(250,120)

#RadioButton options
$rdnButton = New-Object $RadioObject
$rdnButton.Text ="Husband"
$rdnButton.Location= New-Object System.Drawing.Point(300,80)
$rdnButton.AutoSize=$true
$rdnButton.Parent="dad"


$rdnButton2 = New-Object $RadioObject
$rdnButton2.Text ="Wife"
$rdnButton2.Location= New-Object System.Drawing.Point(300,120)
$rdnButton2.AutoSize=$true
$rdnButton2.Parent="dad"

$Button = New-Object $ButtonObject
$Button.Text ="Execute"
$Button.Location= New-Object System.Drawing.Point(400,120)
$Button.AutoSize=$true



$AppForm.Controls.AddRange(@($lblService,$ddlService,$lblName,$lblForName,$lblForStatus,$lblStatus,$rdnButton,$rdnButton2,$Button))

#Add some functions
function GetServiceDetails {
$ServiceSelected = $ddlService.SelectedItem
$details = Get-Service $ServiceSelected | Select *
$lblName.Text = $details.name
$lblStatus.Text = $details.status
}
function Testing {Write-Host "working" -ForegroundColor green}

function radio {

if ($rdnButton.Checked -eq $true){
    Write-Host $rdnButton.Checked}
else{Write-Host "not selected" -ForegroundColor red}


}


####### Add functions to the controls
$ddlService.Add_SelectedIndexChanged({GetServiceDetails})

$Button.Add_Click({radio})

#Show the form
$AppForm.ShowDialog()

#Garbage Collection
$AppForm.Dispose()