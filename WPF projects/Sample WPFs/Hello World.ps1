Add-Type -AssemblyName System.Windows.Forms #Bring in the windows forms to we can build forms

$FormObject = [System.Windows.Forms.Form]
$LabelObject = [System.Windows.Forms.Label]
$ButtonObject = [System.Windows.Forms.Button]

$Window = New-Object $FormObject
$Window.ClientSize='500,300'
$Window.MinimumSize='500,300'
$Window.MaximumSize='500,300'
$Window.Text = 'Helo World - Tutorial'
$Window.BackColor = 'lightblue'

$lbtitle = New-Object $LabelObject
$lbTitle.text ="Hello World"
$lbTitle.AutoSize = $True #makes label to fit the word
$lbTitle.Font="Verdana,20,style=Bold"
$lbTitle.Location = New-Object System.Drawing.Point(120,100)

$btnHello = New-Object $ButtonObject
$btnHello.Text="Say Hello"
$btnHello.AutoSize = $True
$btnHello.Font = "Verdana,14"
$btnHello.Location = New-Object System.Drawing.Point(172,180)

$window.Controls.AddRange(@($lbTitle,$btnHello))

### Logic Section/Functions
function SayHello{
    if($lbtitle.Text -eq ""){
    $lbtitle.Text = "Hello World"}
    Else{$lbtitle.Text=""}
}
    

### Add Functions to Form
$btnHello.Add_Click({SayHello})

#Display the form
$Window.ShowDialog()

#Cleans up dialog
$Window.Dispose()