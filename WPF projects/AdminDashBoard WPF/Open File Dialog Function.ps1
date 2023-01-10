$userProfile = $env:USERPROFILE
$Path ="$userProfile\Downloads"
        
        $FileDialogObject = [System.Windows.Forms.OpenFileDialog]
 # Create Object and add properties

        
        $OpenFileDialog = New-Object $FileDialogObject
        $OpenFileDialog.InitialDirectory = $Path
        $OpenFileDialog.CheckPathExists = $true
        $OpenFileDialog.CheckFileExists = $true
        $OpenFileDialog.Title = "Please feed me"
        $OpenFileDialog.Filter = "csv files (*.csv)|*.csv" #|All files (*.*)|*.*"

        $OpenFileDialog.ShowDialog()

        if ($openFileDialog.CheckFileExists -eq $true) {
            $filePath = $openFileDialog.FileName
            Write-Host $filePath}