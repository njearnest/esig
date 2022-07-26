# generate esignature - pull data from Azure AD
# search by email address

#Get-ChildItem Env: | Sort Name
$UserAppData = $Env:APPDATA

$DateStamp = (Get-Date -format ddMMMyyyy)
#$eSignaturePath = "C:\Users\nate.earnest\OneDrive\Development\DoubleStar\esig vs PowerShell\Generated\"
$eSignaturePath = "$UserAppData\Microsoft\Signatures\"


######################################################
# CONNECT TO OFFICE 365 via PowerShell
######################################################
write-host "Connecting to Office 365"
$LiveCred = Get-Credential 
Connect-AzureAD -Credential $LiveCred
######################################################

Write-Host "Gathering basic information..." -ForegroundColor Yellow
$UPNorEmail = Read-host "Enter the email address of the user?"

Get-AzureADUser -ObjectId $UPNorEmail


if (Get-AzureADUser -Filter "UserPrincipalName eq '$UPNorEmail'")
{
    $AzureADUser = Get-AzureADUser -ObjectId $UPNorEmail

    $DisplayName = $AzureADUser.GivenName + " " + $AzureADUser.Surname
    $Title = $AzureADUser.JobTitle
    $Phone = $AzureADUser.TelephoneNumber
    $Email = $AzureADUser.Mail

    $FileName = "esig_" + $DisplayName + "_"+ $Title + "_" + $DateStamp
    $FileName = $FileName -replace(" ","")

    $MyeSignaturePath = $eSignaturePath + $FileName 
    if (!(Test-Path $eSignaturePath))
    {
        #path doesnt exist, create
        New-Item -Path $eSignaturePath -ItemType directory
    }
    else
    {
        "all good"
    }

    ###################################################################
    #Generate TXT file
    ###################################################################
    $txtFilePath = $eSignaturePath + $FileName + ".txt"
    Add-Content -Path $txtFilePath -Value $DisplayName
    Add-Content -Path $txtFilePath -Value $Title
    Add-Content -Path $txtFilePath -Value $Phone
    Add-Content -Path $txtFilePath -Value $Email

    ###################################################################
    #Generate HTML file
    ###################################################################
    $htmlFilePath = $eSignaturePath + $FileName + ".htm"
    Add-Content -Path $htmlFilePath -Value "<html><head>"
    Add-Content -Path $htmlFilePath -Value "<style type=`"text/css`">"
    Add-Content -Path $htmlFilePath -Value "body"
    Add-Content -Path $htmlFilePath -Value "{"
    Add-Content -Path $htmlFilePath -Value "font-family:Calibri;"
    Add-Content -Path $htmlFilePath -Value "padding:0px;"
    Add-Content -Path $htmlFilePath -Value "margin:0px;"
    Add-Content -Path $htmlFilePath -Value "}"
    Add-Content -Path $htmlFilePath -Value "table"
    Add-Content -Path $htmlFilePath -Value "{"
    Add-Content -Path $htmlFilePath -Value "width:500px;"
    Add-Content -Path $htmlFilePath -Value "}"
    Add-Content -Path $htmlFilePath -Value "td"
    Add-Content -Path $htmlFilePath -Value "{"
    Add-Content -Path $htmlFilePath -Value "padding:0px;"
    Add-Content -Path $htmlFilePath -Value "margin:0px;"
    Add-Content -Path $htmlFilePath -Value "}"
    Add-Content -Path $htmlFilePath -Value "a, img"
    Add-Content -Path $htmlFilePath -Value "{"
    Add-Content -Path $htmlFilePath -Value "border: 0px;"
    Add-Content -Path $htmlFilePath -Value "text-decoration:none;"
    Add-Content -Path $htmlFilePath -Value "}"
    Add-Content -Path $htmlFilePath -Value "</style>"
    Add-Content -Path $htmlFilePath -Value "</head><body>"
    Add-Content -Path $htmlFilePath -Value "<table>"
    Add-Content -Path $htmlFilePath -Value "<tr><td style=`"font-size:12pt;`"><b>$DisplayName</b></td></tr>"
    Add-Content -Path $htmlFilePath -Value "<tr><td style=`"font-size:10pt;`">$Title</td></tr>"
    Add-Content -Path $htmlFilePath -Value "<tr><td style=`"font-size:10pt;`">P $Phone</td></tr>"
    Add-Content -Path $htmlFilePath -Value "<tr><td style=`"font-size:10pt;`"><a href=`"mailto:$Email`">$Email</a></td></tr>"
    Add-Content -Path $htmlFilePath -Value "<tr><td><table style=`"width:500px;`">"
    Add-Content -Path $htmlFilePath -Value "<tr><td colspan=`"5`"><hr /></td></tr>"
    Add-Content -Path $htmlFilePath -Value "<td style=`"width:320px;`"><a href=`"http://bit.ly/DoubleStar`"><img src=`"D-Star-Logo-all-rgb.png`" alt=`"Visit the DoubleStar website`" style=`"height:87px;width:200px;`"></a></td>"
    Add-Content -Path $htmlFilePath -Value "<td style=`"width:30px;`"><a href=`"http://bit.ly/DSWelcome`"><img src=`"32x32-youtube.png`" alt=`"Watch us on YouTube`" style=`"height:32px;width:32px;`"></a></td>"
    Add-Content -Path $htmlFilePath -Value "<td style=`"width:30px;`"><a href=`"http://bit.ly/DoubleStarFacebook`"><img src=`"32x32-facebook.png`" alt=`"Follow us on Facebook`" style=`"height:32px;width:32px;`"></a></td>"
    Add-Content -Path $htmlFilePath -Value "<td style=`"width:30px;`"><a href=`"http://bit.ly/DoubleStarTwitter`"><img src=`"32x32-twitter.png`" alt=`"Follow us on Twitter`" style=`"height:32px;width:32px;`"></a></td>"
    Add-Content -Path $htmlFilePath -Value "<td style=`"width:30px;`"><a href=`"http://bit.ly/DoubleStarLinkedIn`"><img src=`"32x32-linkedin.png`" alt=`"Link us on LinkedIn`" style=`"height:32px;width:32px;`"></a></td>"
    Add-Content -Path $htmlFilePath -Value "</tr></table></td></tr>"

    Add-Content -Path $htmlFilePath -Value "</table>"
    Add-Content -Path $htmlFilePath -Value "</body></html>"

    Invoke-WebRequest -Uri "https://raw.githubusercontent.com/njearnest/esig/main/32x32-facebook.png" -OutFile $($eSignaturePath + "32x32-facebook.png")
    Invoke-WebRequest -Uri "https://raw.githubusercontent.com/njearnest/esig/main/32x32-linkedin.png" -OutFile $($eSignaturePath + "32x32-linkedin.png")
    Invoke-WebRequest -Uri "https://raw.githubusercontent.com/njearnest/esig/main/32x32-twitter.png" -OutFile $($eSignaturePath + "32x32-twitter.png")
    Invoke-WebRequest -Uri "https://raw.githubusercontent.com/njearnest/esig/main/32x32-youtube.png" -OutFile $($eSignaturePath + "32x32-youtube.png")
    Invoke-WebRequest -Uri "https://raw.githubusercontent.com/njearnest/esig/main/D-Star-Logo-all-rgb.png" -OutFile $($eSignaturePath + "D-Star-Logo-all-rgb.png")
}

######################################################
# END / CLOSE ALL SESSIONS!
######################################################
write-host "Disconnecting from Office 365"
Remove-PSSession $Session