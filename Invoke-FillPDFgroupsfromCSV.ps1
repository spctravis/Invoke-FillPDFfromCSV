$csvpath = '\\uhhz-FS-301\VDI\Documents\r2d2\r2d2.csv'

Add-Type -AssemblyName System.Windows.Forms

$count = 1

$group = 5

 

 

Start-Sleep -Seconds 2

$csv = Import-Csv -Path $csvpath

for ($i = 0; $i -lt $csv.Count; $i +=5) {

    for ($j = $i; $j -lt ($i +5) -and $j -lt $csv.Count; $j++) {

        $x = $csv[$j]

        #Open Blank form

        Start-Sleep -Seconds 1

        #B

        [System.Windows.Forms.SendKeys]::SendWait($x.LastName + ", " + $x.FirstName)

        Start-Sleep -Seconds 1

        #1

        [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + $x.DODID)

        Start-Sleep -Seconds 1

        [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + $x.EMAIL)

        Start-Sleep -Seconds 1

        [System.Windows.Forms.SendKeys]::SendWait("{TAB}")

   }

        Start-Sleep -Seconds 1

        [System.Windows.Forms.SendKeys]::SendWait("^+{S}" )

        Start-Sleep -Seconds 2

        [System.Windows.Forms.SendKeys]::SendWait("R2D2AccountRequest_$count")

        $count ++

                Start-Sleep -Seconds 1

        [System.Windows.Forms.SendKeys]::SendWait("{TAB}")

                Start-Sleep -Seconds 1

        [System.Windows.Forms.SendKeys]::SendWait("{TAB}")

                Start-Sleep -Seconds 1

        [System.Windows.Forms.SendKeys]::SendWait("{TAB}" + "{ENTER}")

                        Start-Sleep -Seconds 4

        [System.Windows.Forms.SendKeys]::SendWait("{TAB}")

        <#

        for ($k = 1; $k -le 15; $k++){

        [System.Windows.Forms.SendKeys]::SendWait("{SHIFT}({TAB})")

        Start-Sleep -Seconds .5

        }

        #>

        }

 

write-host "done"
