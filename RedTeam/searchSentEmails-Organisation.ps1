## https://office365itpros.com/2022/04/13/message-tracing-email-activity/
$Mbx = Get-EXOMailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited -identity 'denise.harkin@fitzroy.org' -Properties Office
$Report = [System.Collections.Generic.List[Object]]::new() 

   [string]$SenderAddress = $Mbx.PrimarySmtpAddress
   [array]$Messages = Get-MessageTrace -StartDate $StartDate -EndDate $EndDate -SenderAddress $SenderAddress -Status Delivered | ? {$_.Subject -like "*We want your opinion*"}

      ForEach ($M in $Messages) {
     $ReportLine = [PSCustomObject][Ordered]@{
          Date      = Get-Date($M.Received) -format g 
          User      = $M.SenderAddress
          Recipient = $M.RecipientAddress
          Subject   = $M.Subject
          MessageId = $M.MessageId }
     $Report.Add($ReportLine)
   } #End Foreach messages

   $Report