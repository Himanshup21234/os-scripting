$Username = "hpant30@gmail.com";
$Password = "***********";
$path = "E:\Time Tracker.xlsm";

function Send-ToEmail([string]$email, [string]$attachmentpath){

    $message = new-object Net.Mail.MailMessage;
    $message.From = "YourName@gmail.com";
    $message.To.Add($email);
    $message.Subject = "Schedule Tracker";
    $message.Body = "PFA";
    $attachment = New-Object Net.Mail.Attachment($attachmentpath);
    $message.Attachments.Add($attachment);

    $smtp = new-object Net.Mail.SmtpClient("smtp.gmail.com", "587");
    $smtp.EnableSSL = $true;
    $smtp.Credentials = New-Object System.Net.NetworkCredential($Username, $Password);
    $smtp.send($message);
    write-host "Mail Sent" ; 
    $attachment.Dispose();
 }
Send-ToEmail  -email "************" -attachmentpath $path;
