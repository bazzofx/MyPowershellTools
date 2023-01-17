<#
Why are you using a Try and Catch inside the switch?You might as well nest that too and build your logic inside the FunctionForOption2

Also make sure to add a Default value on your switch, as I am sure there will be somene pressing something other than 1 or 2.

I am not sure exactly what you building but maybe something like this might help somehow?
#>

Function MyMenu
{
	Write-host "My Menu with Options"
	Write-host "1. Do something"
	Write-host "2. Do something which requires Admin to continue"
	$answer=Read-host "Selection:"
    $answer
	switch($answer)
	{
		1 {FunctionForOption1}
		
		2 {FunctionForOption2}
        Default{
        Clear-Host
        Write-Host "Back to Menu"
        MyMenu}
	}
}
Function FunctionForOption1{
    #Do Stuff
    Start-Process notepad
}
Function FunctionForOption2{
			try {
                #Do Something that requires Admin credentials
				Start-Process notepad -verb runas
			}
			catch {
				#Invalid password message or password required message returned from clicking on cancel
        		Write-Host "Password wrong please try again" -ForegroundColor Red		
                MyMenu
			}
}
MyMenu