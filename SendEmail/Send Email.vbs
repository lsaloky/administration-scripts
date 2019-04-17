Set objEmail = CreateObject("CDO.Message")
 
With objEmail 
  .From = "noreply@domain.com"
  .To = "recipient@domain.com"
  .Subject = "Test" 
  .Textbody = "Test."
  .Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
  .Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "SERVERNAME" 
  .Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
  .Configuration.Fields.Update
  .Send
End With
