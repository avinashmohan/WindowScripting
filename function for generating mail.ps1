function PrepareEmail($packageVersion,$PackagesArray,$ConfigurationsArray, $log_file) { 
             
    
    $packagesDelivered = ""
    $configurationDelivered = ""
    
    if ($PackagesArray -eq $null) {
        $packagesDelivered = "<li>N/A</li>"
    } else {
       foreach ($element in $PackagesArray) {
              $packagesDelivered = $packagesDelivered + "<li>"+$element.toString() +"</li>"
                     #logMessage "inside package"
       }
    }
       
    if ($ConfigurationsArray -eq $null) { 
        $configurationDelivered = "<li>N/A</li>"
    } else {
        foreach ($element in $ConfigurationsArray) {
              $configurationDelivered = $configurationDelivered + "<li>"+$element.toString() +"</li>"
       }
    }

       
    $objOutlook = New-Object -comObject Outlook.Application
    $packageVersionWith_ = $packageVersion.Replace('.','_')
    $mail = $objOutlook.CreateItem(0)
       $mail.Recipients.Add('bat.man@xyz.com')
       $mail.CC = "grtgaz-dsi-igeco-io@xyz.com; grtgaz-dsi-etr-func@xyz.com; grtgaz-dsi-etr-idc@xyz.com";  
  #$mail.Recipients.Add('avinash.mohan@xyz.com')
   #$mail.CC = "ben.affleck@xyz.com" 
       $mail.Subject = "$packageVersion"
       
       $mail.Attachments.Add($log_file)
    $mail.HTMLBody = 
        "<font face='Calibri' size='10pt'>
                     Hello,
         <br><br>
         Testing!! .
         <br><br>
                     packages delivered: 
                      <ul>$packagesDelivered</ul>
        <br><br>

            

         
         Cordialement,<br>
                     __________________________________________________________<br>
         <table>
                     <tr><td>Team</td><td rowspan=2 style='padding-left:25px'></td></tr>
         <tr><td>New delhi</td></tr>
                     </table>
        </font>
                     "
    #$mail.save()

    $inspector = $mail.GetInspector
    $inspector.Display()
        
}
