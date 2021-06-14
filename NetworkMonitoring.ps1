###################################################################
# Name: Network Device Pingtest                                   #
# Creator: FNLeaf                                                 #
# CreationDate: 05.15.2021                                        #
# LastModified: 05.15.2021                                        #
# Version: 1.0                                                    #
#                                                                 #
# Description: Checks PING availability of the network devices    #
#                                                                 #
###################################################################

#Declare files
Add-PSSnapin Microsoft.Exchange.Management.Powershell.Admin -erroraction silentlyContinue
$attachment = Get-Item 'D:\temp\HXBor7JW_400x400.jpg'                            #Insert your attachment logo
$ServerListFile = "D:\temp\IPAddresses.txt"                              #Insert PC/IP list to file
$ServerList = Get-Content $ServerListFile -ErrorAction SilentlyContinue  #Do not change
$LogTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"                        #Do not change
$outfile = 'D:\temp\Output.htm'                                          #Change depending on area
$area = "Area1"                                                         #Change depending on area
$recipients = "test1@gmail.com, testemail2@gmail.com"        

#EMAIL function
function sendMail($datestamp){
    $smtpServer = "smtp.gmail.com"                                       #Change depending on smtp server
    $srvPort = 587                                                       #Change depending on port required
    $smtpFrom = "noreply@yahoo.com"
    $smtpTo = $recipients                                                #Do not change
    #pass $datestamp into email subject
    $messageSubject = "PC Network Status Monitoring - $datestamp"
    $message = New-Object System.Net.Mail.MailMessage $smtpfrom, $smtpto 
    $message.Subject = $messageSubject 
    $message.IsBodyHTML = $true 
    $message.Body = "<head><pre>$style</pre></head>"
    $message.Body += '<br /><img src="{0}" />' -f ($attachment.Name)
    $message.Attachments.Add($attachment.FullName)   
    $message.Body += Get-Content $outfile 
    $smtp = New-Object Net.Mail.SmtpClient($smtpServer, $srvPort)
    $smtp.EnableSsl = $true
    $smtp.Credentials = New-Object System.Net.NetworkCredential("<yourusername>", "<yourpassword>") #Change depending on user account needed
    ##Send message
    $smtp.Send($message)
}


$failCount = 0

$Outputreport = "<h3 font-family:Arial>$area</h3>"
$Outputreport += "<table BORDER=5 BORDERCOLOR=DeepSkyBlue border-radius=10px><tr style='background-color: DeepSkyBlue ; color: black;-moz-border-radius:10px;-webkit-border-radius:10px;'><td><b>HOSTNAME</b></td><td><b>STATUS</b></td></tr>"

foreach ($Server in $ServerList) { 
    if (test-Connection -ComputerName $Server -Count 2 -Quiet ) {  
        #show some status in console
        echo "$Server is pingable"
        $Outputreport += "<tr><td>$Server</td><td style='background-color: green; color: white'>ONLINE</td></tr>"
    } else { 
        #show some status in console
        echo "$Server is offline" 
        $Outputreport += "<tr><td>$Server</td><td style='background-color: red'>OFFLINE</td></tr>"
        $failCount = $failCount + 1
    }     
} 

$Outputreport += "</table>"
##Save results to .html file for e-mail
$Outputreport | out-file $outfile


##If the counter is greater than 1, send failure email

if ($failCount -ge 1) {
    sendMail "$LogTime"
}