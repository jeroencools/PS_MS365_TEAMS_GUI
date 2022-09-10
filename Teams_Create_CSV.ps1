
<# 
Werkende met:
2.0.2.118            AzureAD
2.5.2-preview        MicrosoftTeams        
#>


#######################################################################################################################################
                                                       
            # Importeren van het .csv bestand in dezelfde directory als je script
$csv = Import-Csv "$PSScriptRoot\teams.csv"

            # Dit stuk moet wel wat "opkuiswerk" ondergaan - geen nette code
$titucsv = $csv | select -ExpandProperty titu
$llncsv = $csv | select -ExpandProperty lln
$vakkencsv = $csv | select -ExpandProperty vakken

$totaltitu = $titucsv -split ";"
$totallln = $llncsv -split ";"
$totalvakken = $vakkencsv -split ";"
$total = $totaltitu + $totallln + $totalvakken
$complete = $total.Count
$global:tellen = 0

        # Verbinden met MS Teams + AzureAD

$credentials = Get-Credential
Connect-MicrosoftTeams -Credential $credentials  
Connect-AzureAD -Credential $credentials  
Start-Sleep -Seconds 4


########################### FUNCTIES ############################################################################################################
        # Toevoegen van de leerkrachten = owners
function toevoegen_leerkrachten
{
    param($leerkrachten,$groupid,$rol)
    $usersplit = $leerkrachten -split ";"
    for($j =0; $j -le ($usersplit.count - 1) ; $j++)
        {
            $global:tellen++
            $pcomplete = ($global:tellen / $complete) * 100
            Write-Progress -Activity "Activiteit $global:tellen van de $complete" -Status $usersplit[$j] -PercentComplete    $pcomplete

            Add-TeamUser -GroupId $GroupId -User $usersplit[$j] -Role $rol
        }

}

        # Toevoegen van leerlingen = members + private channels aanmaken als "0. Voornaam Achternaam" per leerling
function toevoegen_leerlingen
{
    param($leerlingen,$groupid,$rol)
    $usersplit = $leerlingen -split ";"
    for($j =0; $j -le ($usersplit.count - 1) ; $j++)
        {
            $global:tellen++
            $pcomplete = ($global:tellen / $complete) * 100
            Write-Progress -Activity "Activiteit $global:tellen van de $complete" -Status $usersplit[$j] -PercentComplete    $pcomplete

            Add-TeamUser -GroupId $GroupId -User $usersplit[$j] -Role $rol
            $naamkanaal = Get-AzureADUser -ObjectId $usersplit[$j] | Select Displayname
            $metnul = '0.' + $naamkanaal.DisplayName         
            New-TeamChannel -GroupId $GroupId -DisplayName $metnul -MembershipType Private
            Start-Sleep -Seconds 0.5    
        }

}
    
        # Toevoegen van leerkrachten aan private channels --> Leerlingen moeten achteraf nog gebeuren, want het team is nog niet actief.
function toevoegen_leden_kanalen
{
    param($leerlingen,$groupid,$leerkrachten)
    $usersplit = $leerlingen -split ";"
    $usersplitleerkrachten = $leerkrachten -split ";"
    for($j =0; $j -le ($usersplit.count - 1) ; $j++)
        {
            $naamkanaal = Get-AzureADUser -ObjectId $usersplit[$j] | Select Displayname
            $metnul = '0.' + $naamkanaal.DisplayName         
            for($i =1; $i -le ($usersplitleerkrachten.count - 1) ; $i++)
                {
                    Add-TeamChannelUser -GroupId $GroupId -DisplayName $metnul -User $usersplitleerkrachten[$i]
                    Start-Sleep -Seconds 0.5 
                    Add-TeamChannelUser -GroupId $GroupId -DisplayName $metnul -User $usersplitleerkrachten[$i] -Role Owner
                }                
        }

}


    
        # Toevoegen van de vakken = kanalen
function toevoegen_vakken
{
    param($vakken,$groupid)
    $vaksplit = $vakken -split ";"
    for($j =0; $j -le ($vaksplit.count - 1) ; $j++)
        {
            $global:tellen++
            $pcomplete = ($global:tellen / $complete) * 100
            Write-Progress -Activity "Activiteit $global:tellen van de $complete" -Status $vaksplit[$j] -PercentComplete    $pcomplete
   
            New-TeamChannel -GroupId $GroupId -DisplayName $vaksplit[$j]
        }

}

        # Uitvoeren van alle acties op basis van .csv
function totaaluitvoeren 
{
        # Een foreach loop die de .csv doorloopt op basis van de namen van de kolommen - bijvoorbeeld "naam", "titu"...
    foreach($i in $csv)
        {
            $teamnaam = $i.naam
            $owners = $i.titu
            $leerlingen = $i.lln
            $vak = $i.vakken
            
            Write-Host "-------------------------------------------------------------------------------------------------"
            
            Write-Host "De volgende klas wordt aangemaakt" $teamnaam
            $group = New-Team -MailNickname $teamnaam -displayname $teamnaam -Template "EDU_Class"
            Set-Team -GroupId $group.GroupId -AllowAddRemoveApps $false -AllowCreateUpdateChannels $false -AllowCreateUpdateRemoveTabs $false -AllowDeleteChannels $false -AllowUserDeleteMessages $false -AllowUserEditMessages $false
 
            write-Host "De volgende leerkrachten worden toegevoegd:" $owners
            toevoegen_leerkrachten -leerkrachten $owners -groupid $group.GroupId -rol "Owner"

            Write-Host "De volgende leerlingen worden toegevoegd:" $leerlingen
            toevoegen_leerlingen -leerlingen $leerlingen -groupid $group.GroupId -rol "Member"

            Write-Host "De volgende vakken worden toegevoegd:" $vak
            toevoegen_vakken -vakken $vak -groupid $group.GroupId

            Write-Host "Leerkrachten worden aan de private channels van leerlingen toegevoegd. Vergeet niet leerlingen zelf toe te voegen - Teams is nog niet actief"
            toevoegen_leden_kanalen -leerlingen $leerlingen -groupid $group.GroupId -leerkrachten $owners
        }
            Write-Host "-------------------------------------------------------------------------------------------------"
            Write-Host "klaar - afsluiten maar!"
}

########################### GUI ############################################################################################################

        $guiform = New-Object system.Windows.Forms.Form
        $guiform.ClientSize         = '400,300'
        $guiform.text               = "Maak teams op basis van .csv"
        $guiform.BackColor          = "#ADCDBE"
 
        $knoplaad = New-Object System.Windows.Forms.Button
        $knoplaad.Location = New-Object System.Drawing.Size(15,30)
        $knoplaad.Size = New-Object System.Drawing.Size(150,60)
        $knoplaad.BackColor             = "#ffffff"
        $knoplaad.Text = "Informatie van .csv laden"
        $knoplaad.Add_Click({
        Import-Csv "$PSScriptRoot\teams.csv" -Header naam, titu, lln, vakken | Out-GridView –Title Get-CsvData
        })
        $knoplaad.Cursor = [System.Windows.Forms.Cursors]::Hand
        $guiform.Controls.Add($knoplaad)

        $knopuitvoeren = New-Object System.Windows.Forms.Button
        $knopuitvoeren.Location = New-Object System.Drawing.Size(15,150)
        $knopuitvoeren.Size = New-Object System.Drawing.Size(150,60)
        $knopuitvoeren.BackColor             = "#ffffff"
        $knopuitvoeren.Text = "Uitvoeren met .csv informatie"
        $knopuitvoeren.Add_Click({totaaluitvoeren})
        $knopuitvoeren.Cursor = [System.Windows.Forms.Cursors]::Hand
        $guiform.Controls.Add($knopuitvoeren)
        
        [void]$guiform.ShowDialog()




