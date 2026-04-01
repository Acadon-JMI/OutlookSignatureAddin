Write-Progress -Activity "ACADON Signature Generator" -Status "Start" -PercentComplete 2 -Id 1

$scriptPath = $PSScriptRoot
$SignaturePath = Join-Path -Path $env:APPDATA  -ChildPath "\Microsoft\Signatures\" 

function Show-InputBox( [string] $title, [string] $message, [string] $Default ) 
 {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 

    $userForm = New-Object System.Windows.Forms.Form
    $userForm.Text = "$title"
    $userForm.Size = New-Object System.Drawing.Size(290,150)
    $userForm.StartPosition = "CenterScreen"
        $userForm.AutoSize = $False
        $userForm.MinimizeBox = $False
        $userForm.MaximizeBox = $False
        $userForm.SizeGripStyle= "Hide"
        $userForm.WindowState = "Normal"
        $userForm.FormBorderStyle="Fixed3D"
     
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(115,80)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.Add_Click({$value=$objTextBox.Text;$userForm.Close()})
    $userForm.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(195,80)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({$userForm.Close()})
    $userForm.Controls.Add($CancelButton)

    $userLabel = New-Object System.Windows.Forms.Label
    $userLabel.Location = New-Object System.Drawing.Size(10,20)
    $userLabel.Size = New-Object System.Drawing.Size(280,20)
    $userLabel.Text = "$message"
    $userForm.Controls.Add($userLabel) 

    $objTextBox = New-Object System.Windows.Forms.TextBox
    $objTextBox.Location = New-Object System.Drawing.Size(10,40)
    $objTextBox.Size = New-Object System.Drawing.Size(260,20)
    $objTextBox.Text="$Default"
    $userForm.Controls.Add($objTextBox) 

    $userForm.Topmost = $True
    $userForm.Opacity = 0.91
        $userForm.ShowIcon = $False

    $userForm.Add_Shown({$userForm.Activate()})
    [void] $userForm.ShowDialog()

    $value=$objTextBox.Text 

    return $value
 }

Write-Progress -Activity "ACADON Signature Generator" -Status "Fetch Information" -PercentComplete 10 -Id 1
Write-Progress -Activity "Get Username" -Status "Request User from System"-PercentComplete 33 -Id 2 -ParentId 1

$UserID    = $env:UserName
Write-Progress -Activity "Get Username" -Status "Get Confirmation from User"-PercentComplete 66 -Id 2 -ParentId 1
$UserID    = Show-InputBox  -Title "User Name" -Message "Please check your name." -default "$UserID"

Write-Progress -Activity "ACADON Signature Generator" -Status "Fetch Information" -PercentComplete 20 -Id 1
Write-Progress -Activity "Get User Data" -Status "LDAP Request to Server"-PercentComplete 2 -Id 2 -ParentId 1
$ADserver = "AC12" #
$DS = New-Object System.DirectoryServices.DirectoryEntry( "LDAP://$($ADserver):389/DC=acadon,DC=acadon,DC=de" ) 
$LDAPrequest = New-Object System.DirectoryServices.DirectorySearcher( $DS )

Write-Progress -Activity "Get User Data" -Status "LDAP Request to Server"-PercentComplete 33 -Id 2 -ParentId 1
Write-Progress -Activity "Sending Request and ... " -Status "Waiting for Answer" -SecondsRemaining 6 -Id 3 -ParentId 2
$LDAPrequest.filter = "((sAMAccountName=$UserID))"
$LDAPrequest.SearchScope = "subtree"
[void]$LDAPrequest.PropertiesToLoad.Add("company");
[void]$LDAPrequest.PropertiesToLoad.Add("department");
[void]$LDAPrequest.PropertiesToLoad.Add("title");
[void]$LDAPrequest.PropertiesToLoad.Add("givenname"); # alternativ: msexchshadowgivenname
[void]$LDAPrequest.PropertiesToLoad.Add("sn"); # alternativ: msexchshadowsn
[void]$LDAPrequest.PropertiesToLoad.Add("telephonenumber");
[void]$LDAPrequest.PropertiesToLoad.Add("mail");
[void]$LDAPrequest.PropertiesToLoad.Add("physicaldeliveryofficename");
Write-Progress -Activity "Sending Request and ... " -Status "Waiting for Answer" -SecondsRemaining 5 -Id 3 -ParentId 2
$LDAPresult = $LDAPrequest.FindOne()
Write-Progress -Activity "Sending Request and ... " -Status "Waiting for Answer" -SecondsRemaining 1 -Id 3 -ParentId 2
<# Show Results
  $LDAPresult.Properties
#>
Write-Progress -Activity "Sending Request and ... " -Status "Waiting for Answer" -Completed -Id 3 -ParentId 2
if( -not $LDAPresult ){
    Throw "Unable to connect to DC. Exiting"
}

Write-Progress -Activity "Get User Data" -Status "LDAP Answer from Server" -PercentComplete 66 -Id 2 -ParentId 1
$GivenName    = $LDAPresult.Properties.givenname
$SurName      = $LDAPresult.Properties.sn
$phone        = "+49 2151 96 96 0"
if( $LDAPresult.Properties.telephonenumber )
{ $phone      = $LDAPresult.Properties.telephonenumber }
$title        = "developer"
if( $LDAPresult.Properties.title )
{ $title      = $LDAPresult.Properties.title.ToLower() }
$office       = "Krefeld"
if( $LDAPresult.Properties.physicaldeliveryofficename )
{ $office     = $LDAPresult.Properties.physicaldeliveryofficename.ToLower() }
switch( $office ) 
{
    "Bergeijk"     { $address = "Stokskesweg 9, 5571TJ Bergeijk "  }
    "Braunschweig" { $address = "Berliner Str. 52 j, 38104 Braunschweig, Germany"  }
    "Bremen"       { $address = "Im Hollergrund 3, 28357 Bremen, Germany"  }
    "Dänemark"     { $address = "Königsberger Str. 115, 47809 Krefeld, Germany"  }
    "Dortmund"     { $address = "Rodenbergstraße 47, 44287 Dortmund, Germany"  }
    "Krefeld"      { $address = "Königsberger Str. 115, 47809 Krefeld, Germany"  }

    "Schweiz"      { $address = "acadon (Schweiz) GmbH, General-Guisan-Str. 6, CH-6300 Zug" }
    "Österreich"   { $address = "acadon GmbH, Am Euro Platz 2, AT-1120 Wien" }

    Default        { $address = "Königsberger Str. 115, 47809 Krefeld, Germany"  }
}

Write-Progress -Activity "ACADON Signature Generator" -Status "Fetch Information" -PercentComplete 30 -Id 1
Write-Progress -Activity "Confirm User Data" -Status "Name"-PercentComplete 10 -Id 2 -ParentId 1
$GivenName = Show-InputBox  -Title "GivenName"    -Message "Please check your name." -default "$GivenName"
Write-Progress -Activity "Confirm User Data" -Status "Name"-PercentComplete 20 -Id 2 -ParentId 1
$SurName   = Show-InputBox  -Title "SurName"      -Message "Please check your name." -default "$SurName"
Write-Progress -Activity "Confirm User Data" -Status "Phonenumber"-PercentComplete 30 -Id 2 -ParentId 1
$Phone     = Show-InputBox  -Title "Phonenumber"  -Message "Please check your phone number." -default "$phone"
Write-Progress -Activity "Confirm User Data" -Status "Title"-PercentComplete 40 -Id 2 -ParentId 1
$title     = Show-InputBox  -Title "Title"        -Message "Please Check your title." -default "$title"

Write-Progress -Activity "ACADON Signature Generator" -Status "Fetch Information" -PercentComplete 40 -Id 1
Write-Progress -Activity "Select Languages" -Status "Ask User"-PercentComplete 50 -Id 2 -ParentId 1

$languages = Get-ChildItem $scriptPath -Directory -Filter '??' -Name
$Selections = $languages | Out-GridView -OutputMode Multiple -Title 'Please select one or more language(s)...' 
    if( -not $Selections ) {
        Throw "No language selected. Exiting"
    }

Write-Progress -Activity "ACADON Signature Generator" -Status "Fetch Information" -PercentComplete 50 -Id 1
Write-Progress -Activity "Select Profiles" -Status "Fetch Profiles"-PercentComplete 60 -Id 2 -ParentId 1
$Profiles = ( Get-ChildItem -Path "$env:LOCALAPPDATA\Microsoft\Outlook\*.ost" -File ).BaseName       # Fetch all *.ost Files in Outlook Client Directory
$Profiles = $Profiles | ForEach-Object{ if( $_ -notmatch "^Outlook Data File$" ){ $_ } }             # Remove a local Outlook Profile from list
$Profiles = $Profiles | ForEach-Object{ if( $_ -like "*@acadon.*" ){
                                if( $_ -clike "*$env:UserName*"){  }else{ $_ -ireplace( "^$env:UserName@", "$env:UserName@" ) }
                            }
                            $_
            }

switch ($Profiles.count) {
        0 { 
            Throw "No Outlook Profile found. Exiting" 
        }
        1 {
        }
        Default {
            Write-Progress -Activity "Select Profiles" -Status "Ask User"-PercentComplete 65 -Id 2 -ParentId 1
            $SuggestProfile = ( $Profiles -cmatch "$env:UserName" )[0]
            if( -not $SuggestProfile ){ $SuggestProfile = " NOTHING MATCHING FOUND" }
            $Profiles = $Profiles | Out-GridView -OutputMode Multiple -Title "Please select one or more Outlook profiles. SUGGESTED Profile: $SuggestProfile"
        }
    }
    if( -not $Profiles ){
        Throw "No Outlook Profile selected. Exiting"
    }


Write-Progress -Activity "ACADON Signature Generator" -Status "Write Files" -PercentComplete 60 -Id 1
$Pass2=0
Write-Verbose " - Found $($Selections.count) Languages"
foreach( $Selection in $Selections )
{   
    $Pass2++
    Write-Verbose " -- About to process Language $Selection"
    Write-Progress -Activity "Language:" -Status "$Selection" -PercentComplete (100*( $Pass2 / $($Selections.count) ) ) -Id 2 -ParentId 1
    $FileSet = Get-ChildItem -Path $scriptPath\$Selection -File -ErrorAction SilentlyContinue
    Write-Verbose "  --> Found $($FileSet.count) Items"
    if( $FileSet.count -eq 0 ){
        Write-Error -Message " No Items found to process for language $Selection "
    }
    $Pass3=0
    foreach( $file in $FileSet )
    {
        $Pass3++
        Write-Progress -Activity "File:" -Status "$File" -PercentComplete ( 100*( $Pass3 / $($FileSet.count) ) ) -Id 3 -ParentId 2
        Write-Verbose " --- About to read from $($file.FullName)"
        $content = Get-Content -Path $file.FullName -Encoding Default 
        $content = $content.Replace( '@@MailAddress@@', "$UserID@acadon.net"  )
        $content = $content.Replace( '@@FirstName@@', $GivenName )
        $content = $content.Replace( '@@LastName@@', $SurName )
        $content = $content.replace( '@@TelefonNumber@@', $phone )
        $content = $content.replace( '@@JobTitle@@', $Title )
        $content = $content.replace( '@@Address@@', $address )
        if( $($file.extension) -eq "rtf") 
        {
            Write-Verbose "Substitude UTF with RTF encoding"
            $content = $content.replace( 'Ä', "\'c4" )
            $content = $content.replace( 'ä', "\'e4" )
            $content = $content.replace( 'Ö', "\'d6" )
            $content = $content.replace( 'ö', "\'f6" )
            $content = $content.replace( 'Ü', "\'DC" )
            $content = $content.replace( 'ü', "\'fc" )
            $content = $content.replace( 'ß', "\'df" )
        }
        foreach( $Profil in $Profiles )
        {
            $Target  =  Join-Path -Path $SignaturePath -ChildPath ( "$($file.BaseName)_$($Selection) ($Profil)$($file.Extension)")
            Write-Verbose "About to write to $Target"
            If( Test-Path -Path $Target -PathType Leaf ){ Remove-Item -Path $Target -Force -ErrorAction SilentlyContinue }
            $content | Set-Content -Path $Target -Encoding Default -Force
        }
    }
    
}
Write-Progress -Activity "ACADON Signature Generator" -Status "Finished writing Files" -PercentComplete 100 -Id 1
Write-Host "Finished Script"


# SIG # Begin signature block
# MIIZ3QYJKoZIhvcNAQcCoIIZzjCCGcoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUxlpK5HacqO4V/vYQWhYP4Wrk
# OtmgghPoMIIF+zCCA+OgAwIBAgITGAAAQ2+wlX5eJN8J2QACAABDbzANBgkqhkiG
# 9w0BAQsFADBYMRIwEAYKCZImiZPyLGQBGRYCZGUxFjAUBgoJkiaJk/IsZAEZFgZh
# Y2Fkb24xFjAUBgoJkiaJk/IsZAEZFgZhY2Fkb24xEjAQBgNVBAMTCWFjYWRvbiBB
# RzAeFw0yMjEwMTAxMDE0NDdaFw0yNDEwMDkxMDE0NDdaMIGBMRIwEAYKCZImiZPy
# LGQBGRYCZGUxFjAUBgoJkiaJk/IsZAEZFgZhY2Fkb24xFjAUBgoJkiaJk/IsZAEZ
# FgZhY2Fkb24xEDAOBgNVBAsTB0tyZWZlbGQxEDAOBgNVBAsTB1RlY2huaWsxFzAV
# BgNVBAMTDlRvcnN0ZW4gTmV0emVsMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIB
# CgKCAQEA1j5RLE2Rff+DKHpr4nH2PoiEFvl1UUFG0YCcON0hKflq9c3bWlvhkueY
# uAx4L0CnkUd5oPM9pe1n90usG/OV0fKGNJuFAROhRQKnoXuiSN1hxCi1RkU6jX24
# T1ty8cL+1JZx9H++v73ZItUwAJYlrpm6SNtW6u3fefASwglL6kMO88ooOk0xzXQJ
# OgbHcjVyV8HcUjNfkuyx5mOz2Oq8AdUenKggDRl4lI3qc7VAO2koEbZ+QvEaz7h4
# cfGBiyhKgGPgGGwiD8Mjmj/kuw2yvdqRRSOnhIYS0ve+MkKXNvzWwhBaj4vPPD7c
# vj9jPRSzXnsHuofNoLVhhbFblUeOTQIDAQABo4IBkjCCAY4wPgYJKwYBBAGCNxUH
# BDEwLwYnKwYBBAGCNxUIhN3xZ4Omk3GHjYkyhq6WUoLbiTqBGIP9oxCHvaFwAgFk
# AgEFMBMGA1UdJQQMMAoGCCsGAQUFBwMDMAsGA1UdDwQEAwIHgDAbBgkrBgEEAYI3
# FQoEDjAMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQEP/1GhKrfHU30StLvIbsNvOXy
# cTAfBgNVHSMEGDAWgBSVPiCUBZoNVSVyI1OLZEWgaenKQzBEBgNVHR8EPTA7MDmg
# N6A1hjNodHRwOi8vd2VibWFpbC5hY2Fkb24uZGUvQ2VydEVucm9sbC9hY2Fkb24l
# MjBBRy5jcmwwNwYDVR0RBDAwLqAdBgorBgEEAYI3FAIDoA8MDXRuZUBhY2Fkb24u
# ZGWBDXRuZUBhY2Fkb24uZGUwTgYJKwYBBAGCNxkCBEEwP6A9BgorBgEEAYI3GQIB
# oC8ELVMtMS01LTIxLTI0OTgyMzY5OTktMTUwMDA0Nzk4My0zNjczMDg4MzYtMTEx
# ODANBgkqhkiG9w0BAQsFAAOCAgEATV3nSczzw7DdLYXs8dbAFYQi2SozlbbNgcl/
# WgYHfp3u2ZxNcsYIW+Rz4qs3fqLzEki15uxftzxLRfPhGRTFdHIQVcgx+7x5q1LE
# f56x8ZupIT7zBHVLGwudNZutw8rLsjmT4p62vaCocU7peRSeTOmCbXUsS8YnM54V
# Du2AmDY0l31tPoGi+fwOnosepKc31A7VaD/E/p3vRm85K9gNG4ssgdClpgwxBhLB
# CEks9OO5K0No/l/eOiox/IelqAipgnZMlzBx7L4zHzgC4Wlvgb+z1jc4PDjbkCaQ
# pZ3gvCV7piGtCOflm7ukeDZ0BeLco/HW9J3lWz96CURZ1aMLfvNuBpGohCZwVw6/
# LeSX4RAOLaQ1La1oldM1IHlLnFvg+muiIukAX8kHVwoZ0po+rkClYhbRiDbLRkcR
# FZMgsQPXIZrChqzUD9xpDl5EcFe+j4syURMAi+OhfoFcMwEbVIEicxJOUj57qBx+
# NP03vi8+d6wmvudZ9jqdX/QsneI4Ggnje33w2eIqfZzLkfuV3q5BLhIPqUrdIoSr
# 5nhN+2HG7kFBRMEBeWWHyghHHC3ocA+klxbQRmGk36kqDWcYdsWKpJIvHe5lgTiT
# mOkSegzJq60ELtdaOcsKwJbQm/D+/9tBrlspLOo5hQzP5Hc1+Th77xob42kkUHXV
# xQPlvw4wggbsMIIE1KADAgECAhAwD2+s3WaYdHypRjaneC25MA0GCSqGSIb3DQEB
# DAUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMKTmV3IEplcnNleTEUMBIGA1UE
# BxMLSmVyc2V5IENpdHkxHjAcBgNVBAoTFVRoZSBVU0VSVFJVU1QgTmV0d29yazEu
# MCwGA1UEAxMlVVNFUlRydXN0IFJTQSBDZXJ0aWZpY2F0aW9uIEF1dGhvcml0eTAe
# Fw0xOTA1MDIwMDAwMDBaFw0zODAxMTgyMzU5NTlaMH0xCzAJBgNVBAYTAkdCMRsw
# GQYDVQQIExJHcmVhdGVyIE1hbmNoZXN0ZXIxEDAOBgNVBAcTB1NhbGZvcmQxGDAW
# BgNVBAoTD1NlY3RpZ28gTGltaXRlZDElMCMGA1UEAxMcU2VjdGlnbyBSU0EgVGlt
# ZSBTdGFtcGluZyBDQTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAMgb
# Aa/ZLH6ImX0BmD8gkL2cgCFUk7nPoD5T77NawHbWGgSlzkeDtevEzEk0y/NFZbn5
# p2QWJgn71TJSeS7JY8ITm7aGPwEFkmZvIavVcRB5h/RGKs3EWsnb111JTXJWD9zJ
# 41OYOioe/M5YSdO/8zm7uaQjQqzQFcN/nqJc1zjxFrJw06PE37PFcqwuCnf8DZRS
# t/wflXMkPQEovA8NT7ORAY5unSd1VdEXOzQhe5cBlK9/gM/REQpXhMl/VuC9RpyC
# vpSdv7QgsGB+uE31DT/b0OqFjIpWcdEtlEzIjDzTFKKcvSb/01Mgx2Bpm1gKVPQF
# 5/0xrPnIhRfHuCkZpCkvRuPd25Ffnz82Pg4wZytGtzWvlr7aTGDMqLufDRTUGMQw
# mHSCIc9iVrUhcxIe/arKCFiHd6QV6xlV/9A5VC0m7kUaOm/N14Tw1/AoxU9kgwLU
# ++Le8bwCKPRt2ieKBtKWh97oaw7wW33pdmmTIBxKlyx3GSuTlZicl57rjsF4VsZE
# Jd8GEpoGLZ8DXv2DolNnyrH6jaFkyYiSWcuoRsDJ8qb/fVfbEnb6ikEk1Bv8cqUU
# otStQxykSYtBORQDHin6G6UirqXDTYLQjdprt9v3GEBXc/Bxo/tKfUU2wfeNgvq5
# yQ1TgH36tjlYMu9vGFCJ10+dM70atZ2h3pVBeqeDAgMBAAGjggFaMIIBVjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQUGqH4YRkgD8NB
# d0UojtE1XwYSBFUwDgYDVR0PAQH/BAQDAgGGMBIGA1UdEwEB/wQIMAYBAf8CAQAw
# EwYDVR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FD
# ZXJ0aWZpY2F0aW9uQXV0aG9yaXR5LmNybDB2BggrBgEFBQcBAQRqMGgwPwYIKwYB
# BQUHMAKGM2h0dHA6Ly9jcnQudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FBZGRU
# cnVzdENBLmNydDAlBggrBgEFBQcwAYYZaHR0cDovL29jc3AudXNlcnRydXN0LmNv
# bTANBgkqhkiG9w0BAQwFAAOCAgEAbVSBpTNdFuG1U4GRdd8DejILLSWEEbKw2yp9
# KgX1vDsn9FqguUlZkClsYcu1UNviffmfAO9Aw63T4uRW+VhBz/FC5RB9/7B0H4/G
# XAn5M17qoBwmWFzztBEP1dXD4rzVWHi/SHbhRGdtj7BDEA+N5Pk4Yr8TAcWFo0zF
# zLJTMJWk1vSWVgi4zVx/AZa+clJqO0I3fBZ4OZOTlJux3LJtQW1nzclvkD1/RXLB
# GyPWwlWEZuSzxWYG9vPWS16toytCiiGS/qhvWiVwYoFzY16gu9jc10rTPa+DBjgS
# HSSHLeT8AtY+dwS8BDa153fLnC6NIxi5o8JHHfBd1qFzVwVomqfJN2Udvuq82EKD
# QwWli6YJ/9GhlKZOqj0J9QVst9JkWtgqIsJLnfE5XkzeSD2bNJaaCV+O/fexUpHO
# P4n2HKG1qXUfcb9bQ11lPVCBbqvw0NP8srMftpmWJvQ8eYtcZMzN7iea5aDADHKH
# wW5NWtMe6vBE5jJvHOsXTpTDeGUgOw9Bqh/poUGd/rG4oGUqNODeqPk85sEwu8Cg
# Yyz8XBYAqNDEf+oRnR4GxqZtMl20OAkrSQeq/eww2vGnL8+3/frQo4TZJ577AWZ3
# uVYQ4SBuxq6x+ba6yDVdM3aO8XwgDCp3rrWiAoa6Ke60WgCxjKvj+QrJVF3UuWp0
# nr1Irpgwggb1MIIE3aADAgECAhA5TCXhfKBtJ6hl4jvZHSLUMA0GCSqGSIb3DQEB
# DAUAMH0xCzAJBgNVBAYTAkdCMRswGQYDVQQIExJHcmVhdGVyIE1hbmNoZXN0ZXIx
# EDAOBgNVBAcTB1NhbGZvcmQxGDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDElMCMG
# A1UEAxMcU2VjdGlnbyBSU0EgVGltZSBTdGFtcGluZyBDQTAeFw0yMzA1MDMwMDAw
# MDBaFw0zNDA4MDIyMzU5NTlaMGoxCzAJBgNVBAYTAkdCMRMwEQYDVQQIEwpNYW5j
# aGVzdGVyMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0ZWQxLDAqBgNVBAMMI1NlY3Rp
# Z28gUlNBIFRpbWUgU3RhbXBpbmcgU2lnbmVyICM0MIICIjANBgkqhkiG9w0BAQEF
# AAOCAg8AMIICCgKCAgEApJMoUkvPJ4d2pCkcmTjA5w7U0RzsaMsBZOSKzXewcWWC
# vJ/8i7u7lZj7JRGOWogJZhEUWLK6Ilvm9jLxXS3AeqIO4OBWZO2h5YEgciBkQWzH
# wwj6831d7yGawn7XLMO6EZge/NMgCEKzX79/iFgyqzCz2Ix6lkoZE1ys/Oer6RwW
# LrCwOJVKz4VQq2cDJaG7OOkPb6lampEoEzW5H/M94STIa7GZ6A3vu03lPYxUA5HQ
# /C3PVTM4egkcB9Ei4GOGp7790oNzEhSbmkwJRr00vOFLUHty4Fv9GbsfPGoZe267
# LUQqvjxMzKyKBJPGV4agczYrgZf6G5t+iIfYUnmJ/m53N9e7UJ/6GCVPE/JefKmx
# IFopq6NCh3fg9EwCSN1YpVOmo6DtGZZlFSnF7TMwJeaWg4Ga9mBmkFgHgM1Cdaz7
# tJHQxd0BQGq2qBDu9o16t551r9OlSxihDJ9XsF4lR5F0zXUS0Zxv5F4Nm+x1Ju7+
# 0/WSL1KF6NpEUSqizADKh2ZDoxsA76K1lp1irScL8htKycOUQjeIIISoh67DuiNy
# e/hU7/hrJ7CF9adDhdgrOXTbWncC0aT69c2cPcwfrlHQe2zYHS0RQlNxdMLlNaot
# UhLZJc/w09CRQxLXMn2YbON3Qcj/HyRU726txj5Ve/Fchzpk8WBLBU/vuS/sCRMC
# AwEAAaOCAYIwggF+MB8GA1UdIwQYMBaAFBqh+GEZIA/DQXdFKI7RNV8GEgRVMB0G
# A1UdDgQWBBQDDzHIkSqTvWPz0V1NpDQP0pUBGDAOBgNVHQ8BAf8EBAMCBsAwDAYD
# VR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDBKBgNVHSAEQzBBMDUG
# DCsGAQQBsjEBAgEDCDAlMCMGCCsGAQUFBwIBFhdodHRwczovL3NlY3RpZ28uY29t
# L0NQUzAIBgZngQwBBAIwRAYDVR0fBD0wOzA5oDegNYYzaHR0cDovL2NybC5zZWN0
# aWdvLmNvbS9TZWN0aWdvUlNBVGltZVN0YW1waW5nQ0EuY3JsMHQGCCsGAQUFBwEB
# BGgwZjA/BggrBgEFBQcwAoYzaHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdv
# UlNBVGltZVN0YW1waW5nQ0EuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5z
# ZWN0aWdvLmNvbTANBgkqhkiG9w0BAQwFAAOCAgEATJtlWPrgec/vFcMybd4zket3
# WOLrvctKPHXefpRtwyLHBJXfZWlhEwz2DJ71iSBewYfHAyTKx6XwJt/4+DFlDeDr
# bVFXpoyEUghGHCrC3vLaikXzvvf2LsR+7fjtaL96VkjpYeWaOXe8vrqRZIh1/12F
# FjQn0inL/+0t2v++kwzsbaINzMPxbr0hkRojAFKtl9RieCqEeajXPawhj3DDJHk6
# l/ENo6NbU9irALpY+zWAT18ocWwZXsKDcpCu4MbY8pn76rSSZXwHfDVEHa1YGGti
# +95sxAqpbNMhRnDcL411TCPCQdB6ljvDS93NkiZ0dlw3oJoknk5fTtOPD+UTT1lE
# ZUtDZM9I+GdnuU2/zA2xOjDQoT1IrXpl5Ozf4AHwsypKOazBpPmpfTXQMkCgsRkq
# GCGyyH0FcRpLJzaq4Jgcg3Xnx35LhEPNQ/uQl3YqEqxAwXBbmQpA+oBtlGF7yG65
# yGdnJFxQjQEg3gf3AdT4LhHNnYPl+MolHEQ9J+WwhkcqCxuEdn17aE+Nt/cTtO2g
# Le5zD9kQup2ZLHzXdR+PEMSU5n4k5ZVKiIwn1oVmHfmuZHaR6Ej+yFUK7SnDH944
# psAU+zI9+KmDYjbIw74Ahxyr+kpCHIkD3PVcfHDZXXhO7p9eIOYJanwrCKNI9RX8
# BE/fzSEceuX1jhrUuUAxggVfMIIFWwIBATBvMFgxEjAQBgoJkiaJk/IsZAEZFgJk
# ZTEWMBQGCgmSJomT8ixkARkWBmFjYWRvbjEWMBQGCgmSJomT8ixkARkWBmFjYWRv
# bjESMBAGA1UEAxMJYWNhZG9uIEFHAhMYAABDb7CVfl4k3wnZAAIAAENvMAkGBSsO
# AwIaBQCgeDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEM
# BgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMCMGCSqG
# SIb3DQEJBDEWBBQ6nfsSxc98XwppIMDrKN0oQR3UOTANBgkqhkiG9w0BAQEFAASC
# AQDPdEKgQVoTL4rNdP3uFHcSIjde72vSVESOc+RCg8RQxEK0mfmWWww4ZKs0qc4d
# Mv98dkwZhHME96UK62TNQz6u6aGhMS302c1BybKmIbiHr9RMH7812A8bZ3m1PMIc
# kD/kkF7aNd2UEg4DPegqA26ZzzEjQjdJUXkflxjzoYmTHyaGK0hT2H46Gg+6ZWGm
# 8vAYv+iHvs1Xz3zWf9SVW6XEzv/XSeQa+/G8pVrSEuzILLr2f0hTddPVpEasAjzW
# ES/tmBVFCWleW0+2cfASL3LOI7PX1hv7Cr+dntRTaMs0a9v8zTigu/nufsJK5RwG
# Liu+mGCwtTldrXFkrt60fxYnoYIDSzCCA0cGCSqGSIb3DQEJBjGCAzgwggM0AgEB
# MIGRMH0xCzAJBgNVBAYTAkdCMRswGQYDVQQIExJHcmVhdGVyIE1hbmNoZXN0ZXIx
# EDAOBgNVBAcTB1NhbGZvcmQxGDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDElMCMG
# A1UEAxMcU2VjdGlnbyBSU0EgVGltZSBTdGFtcGluZyBDQQIQOUwl4XygbSeoZeI7
# 2R0i1DANBglghkgBZQMEAgIFAKB5MBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEw
# HAYJKoZIhvcNAQkFMQ8XDTIzMDczMTEyNTIxNlowPwYJKoZIhvcNAQkEMTIEMDb/
# JquO2S9bdsTFrtl5lL0JCkOtLIURNUJpnNWJ9BFeswruFdayaJOH2kZDsbrKmDAN
# BgkqhkiG9w0BAQEFAASCAgCHtOYldja3hlIZzmu7a+THd0Q/hYhvBGH2kYzExWKt
# N+PDGbo281Wot/rF3Hsw95U3QGDHtPSrbA8Xe0HFQkJHWFhNbcrM6Da6v+S/JI53
# sPNtax/loiUPxVdEOH7BOYKFPFTHBKMi2lci2RmH1k0i8PFDg0lqpWpL/1emT6za
# kuk+4qJmlmbmTno8NMgz4rKNc5Z5GWDytV6hOV2Oi5sVv6wAP8ztxdxIgCRvpGu7
# SjG2yQWiPe8rxWGv6F0meRC96oVH0QuHG82q5/Mn/wu8JlarrHRMe3QvkwptWGUd
# NuBxzPZbTKhze7Hd5a1M7Uze6SIGDL50gzdnQqYXlBAuj+/5HX3pScyycXm7oVTN
# BVbRnide41GScfjPdxeyizO7hzDcni8VvkAGIOFlw7SO8FLKCoUngKkluGKdo2mU
# Lx+I1PL9hGWjaOgFRLNDPDD1VkevlfwJnec1KJCwWXUSP6chhqrDF/U69P7W7qkf
# YsP8o220Rb4KzwznBk/CenHThCYJkMUtIRoCNtuyL3VnfQlJ8dA9/RYrGhNV4Aao
# scf4D+wMRw98Y7txPGUFJ0tJ0OicOygyFUwtBmU3SwAl36DUazFCsevaTNejz1kr
# gX8Feu9tCpWfJTSRWK1SAuP93nP9tyiZSTt+tUVf72h5TMKDvLUv3enhAWNLV6DL
# WQ==
# SIG # End signature block
