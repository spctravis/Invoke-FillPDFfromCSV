
$csvpath = "Path\Security Onion Training List.csv"
$personel_report = 'Path\Personel list.csv'
Add-Type -AssemblyName System.Windows.Forms


Start-Sleep -Seconds 2
$csv = Import-Csv -Path $csvpath
$csv_personel = Import-Csv $personel_report
$csv |ForEach-Object {
    $x = $_
    $personeldata = $csv_personel | where name -Match $x.'Last Name'

    Write-Host $personeldata.name
    #Open Blank form
  #  [System.Windows.Forms.SendKeys]::SendWait("+O")   
   # Start-Sleep -Seconds 1
   # [System.Windows.Forms.SendKeys]::SendWait("sf182 - CISSP - blank")   
   # Start-Sleep -Seconds 1
   # [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "{TAB}" + "{ENTER}")   
   # Start-Sleep -Seconds 30 
    #A
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 2
    #B
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "n")
    Start-Sleep -Seconds 1
    #1
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait($personeldata.NAME)
    Start-Sleep -Seconds 1
    #2
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + $($personeldata.SSAN -replace '-',''))
    Start-Sleep -Seconds 1
    #3
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + $($dob = $personeldata.DOB.split("/"); $dob[2]+"-"+$dob[0]+"-"+$dob[1]))
    Start-Sleep -Seconds 1
    #4
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #5
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #6
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #7
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}"+ "3882 S.Bogus, Idaho 8370")
    Start-Sleep -Seconds 1
    #8
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "208-555-3662")
    Start-Sleep -Seconds 1
    #9
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + $personeldata.EMAIL)
    Start-Sleep -Seconds 1
    #10
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + $personeldata.'DUTY TITLE')
    Start-Sleep -Seconds 1
    #11
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "n")
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #12
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #13
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "s" + "1" + "1" + "1" + "1" + "1")
    Start-Sleep -Seconds 1
    #14
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #15
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #16
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + $personeldata.RANK)
    Start-Sleep -Seconds 1
    #17
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #Section B
    #1a
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "Security Onion Solutions - 717 Bishops Cir, Evans, GA 30809")
    Start-Sleep -Seconds 1
    #1b
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "Boise, ID")
    Start-Sleep -Seconds 1
    #1c
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "N/A")
    Start-Sleep -Seconds 1
    #1d
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "sales@securityonionsolutions.com")
    Start-Sleep -Seconds 1
    #1e
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "https://www.securityonionsolutions.com")
    Start-Sleep -Seconds 1
    #1f
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "Peter Di Giorgio")
    Start-Sleep -Seconds 1
    #2a
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "Security Onion 2 Bundle")
    Start-Sleep -Seconds 1
    #2b
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #3
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "2022-01-03")
    Start-Sleep -Seconds 1
    #4
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "2022-03-25")
    Start-Sleep -Seconds 1
    #5
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "48")
    Start-Sleep -Seconds 1
    #6
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "0")
    Start-Sleep -Seconds 1
    #7
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "s")
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait("0" + "{ENTER}")
    Start-Sleep -Seconds 1
    #8
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "s")
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait("0" + "{ENTER}")
    Start-Sleep -Seconds 1
    #9
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "s") 
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait("1" + "1" + "1" + "1" + "{ENTER}")
    Start-Sleep -Seconds 1
    #10
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "s")
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait("5" + "{ENTER}")
    Start-Sleep -Seconds 1
    #11
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "s")
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait( "0" + "0" + "0" + "{ENTER}")
    Start-Sleep -Seconds 1
    #12
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #13
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "5" + "{ENTER}")
    Start-Sleep -Seconds 1
    #14
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "y")
    Start-Sleep -Seconds 1
    #15
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "n" + "n")
    Start-Sleep -Seconds 1
    #16
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #17
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "3")
    Start-Sleep -Seconds 1
    #18
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "G")
    Start-Sleep -Seconds 1
    #19
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #20
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #21
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    ###Section C
    #1a
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "627.30")
    Start-Sleep -Seconds 1
    #Approp Fund
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #1b
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #1c
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "627.30")
    Start-Sleep -Seconds 1
    #2a
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #Approp Fund
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #2b
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #2c
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #3
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #4
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #5
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #6
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" )
    Start-Sleep -Seconds 1
    #Section D
    #1a
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "Kyle Something")
    Start-Sleep -Seconds 1
    #b
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" +"85484848484")
    Start-Sleep -Seconds 1
    #1c
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" +"Mail@mail.com")
    Start-Sleep -Seconds 1
    #1d
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #1e
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "2022-01-19")
    Start-Sleep -Seconds 1
    #2a
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #2b
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1  
    #2c
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #2d
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1   
    #2e
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #3a
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1   
    #3b
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #3b
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1   
    #3c
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #3d
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1   
    #3e
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #SEction E
    #1a
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1 
    #1b
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #1c
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1   
    #1e
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #Section F
    #1a
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1   
    #1b
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #1c
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1
    #1d
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 1   
    #1e
    Start-Sleep -Seconds 1     
    [System.Windows.Forms.SendKeys]::SendWait("^+S")
    Start-Sleep -Seconds 1
    #save
    [System.Windows.Forms.SendKeys]::SendWait("SF182 - SecOnion " + $personeldata.Name)
    Start-Sleep -Milliseconds 50
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "{TAB}" + "{TAB}" + "{ENTER}")
    Start-Sleep -Seconds 10
    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50

    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50


    [System.Windows.Forms.SendKeys]::SendWait("+{TAB}")
    Start-Sleep -Milliseconds 50
} #end foreach
write-host "done"
