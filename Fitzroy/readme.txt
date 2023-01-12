Change the $wd variable to the location of where you extract the files

This function will get the first and last row of date based if a row named "Name"

1stLine() runs first, it export the first record Employee A START Sickness then it triggers the function loopCheck()

loopCheck()
then checks if next row has the same name as previous role
if has same name, skip and check next row until it finds a row with a diferent name.
When it finds a row with a different name, export previous row, Employee A END Sickness
Then it pushes back to 1stLine function where it will export the first record Employee B Start Sickness,

The loop runs according to the .csv file size