$credentials = Get-Credential
Connect-MicrosoftTeams -Credential $credentials  
Connect-AzureAD -Credential $credentials 
Start-Sleep -Seconds 4





########################### GUI ############################################################################################################

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"

        Title="MainWindow" Height="600" Width="1024">
    <Grid>
        <Grid.Background>
            <SolidColorBrush Color="#FF00366B"/>
        </Grid.Background>
        <Rectangle HorizontalAlignment="Center" Height="90" Margin="0,24,0,0" VerticalAlignment="Top" Width="1024" Fill="WhiteSmoke"/>

        <Label Content="GUI to create Teams by using a .csv-file" HorizontalAlignment="Center" Margin="0,50,0,0" VerticalAlignment="Top" Width="1024" FontSize="24" FontFamily="Arial" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" FontWeight="Bold">
            <Label.Foreground>
                <SolidColorBrush Color="#FF00366B"/>
            </Label.Foreground>
        </Label>
        <CheckBox Content="Create a private channel for each student and add all teachers to those channels?" Name="createprivatechannel" HorizontalAlignment="Left" Margin="102,342,0,0" VerticalAlignment="Top" FontFamily="Arial" FontSize="16" Height="28" Foreground="White" FontWeight="Bold" Width="699"/>
        <TextBox HorizontalAlignment="Left" Name="prefix" Margin="102,392,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="122" Height="31" FontSize="16" FontWeight="Bold"/>
        <Label Content="Default name for these private channels is 'first name last name'. Do you  &#xD;&#xA;want to add a prefix? For example: '0.' - then write the prefix '0.' in this textbox." HorizontalAlignment="Left" Margin="256,379,0,0" VerticalAlignment="Top" Foreground="White" FontSize="16" FontWeight="Bold" RenderTransformOrigin="0.069,-0.814" Width="654"/>
        <Button Content="Check content of .csv-file" Name="checkbutton" HorizontalAlignment="Left" Margin="102,212,0,0" VerticalAlignment="Top" Background="White" FontSize="16" FontWeight="Bold" Width="240" Height="40"/>
        <Button Content="Create teams" Name="createbutton" HorizontalAlignment="Left" Margin="102,470,0,0" VerticalAlignment="Top" Background="White" FontSize="16" FontWeight="Bold" Width="240" Height="40"/>
        <CheckBox Content="Create public channels for all subjects?" Name="createsubjectchannel" HorizontalAlignment="Left" Margin="102,286,0,0" VerticalAlignment="Top" FontFamily="Arial" FontSize="16" Height="28" Foreground="White" FontWeight="Bold" Width="699"/>
        <Image HorizontalAlignment="Left" Height="92" Margin="102,27,0,0" VerticalAlignment="Top" Width="84" Source="/debrem-logo.jpg" RenderTransformOrigin="-0.424,0.546"/>

    </Grid>
</Window>


"@
#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader"; exit}
# Store Form Objects In PowerShell
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)}



########################### FUNCTIES ############################################################################################################


function laden
{
Import-Csv "$PSScriptRoot\teams.csv" -Header name, teachers, students, subjects | Out-GridView –Title Get-CsvData
}

$checkbutton.Add_Click({laden})


function toevoegen_leerkrachten
{
    param($leerkrachten,$groupid,$rol)
    $usersplit = $leerkrachten -split ";"
    for($j =0; $j -le ($usersplit.count - 1) ; $j++)
        {
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

            Add-TeamUser -GroupId $GroupId -User $usersplit[$j] -Role $rol
            $naamkanaal = Get-AzureADUser -ObjectId $usersplit[$j] | Select Displayname
            $metprefix = $prefix.Text + $naamkanaal.DisplayName         
            if ($createprivatechannel.IsChecked -eq $true)
            {
                New-TeamChannel -GroupId $GroupId -DisplayName $metprefix -MembershipType Private
            }
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
            $metprefix = '0.' + $naamkanaal.DisplayName         
            for($i =1; $i -le ($usersplitleerkrachten.count - 1) ; $i++)
                {
                    Add-TeamChannelUser -GroupId $GroupId -DisplayName $metprefix -User $usersplitleerkrachten[$i]
                    Start-Sleep -Seconds 0.5 
                    Add-TeamChannelUser -GroupId $GroupId -DisplayName $metprefix -User $usersplitleerkrachten[$i] -Role Owner
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
             if ($createsubjectchannel.IsChecked -eq $true)
            {
                New-TeamChannel -GroupId $GroupId -DisplayName $vaksplit[$j]
            }
        }

}


        # Uitvoeren van alle acties op basis van .csv
function totaaluitvoeren 
{
$csv = Import-Csv "$PSScriptRoot\teams.csv"

    foreach($i in $csv)
        {
            $teamnaam = $i.name
            $owners = $i.teachers
            $leerlingen = $i.students
            $vak = $i.subjects
            
            Write-Host "-------------------------------------------------------------------------------------------------"
            
            Write-Host "De volgende klas wordt aangemaakt" $teamnaam
            $group = New-Team -MailNickname $teamnaam -displayname $teamnaam -Template "EDU_Class"
            Set-Team -GroupId $group.GroupId -AllowAddRemoveApps $false -AllowCreateUpdateChannels $false -AllowCreateUpdateRemoveTabs $false -AllowDeleteChannels $false -AllowUserDeleteMessages $false -AllowUserEditMessages $false
 
            write-Host "De volgende leerkrachten worden toegevoegd:" $owners
            toevoegen_leerkrachten -leerkrachten $owners -groupid $group.GroupId -rol "Owner"

            Write-Host "De volgende leerlingen worden toegevoegd:" $leerlingen
            toevoegen_leerlingen -leerlingen $leerlingen -groupid $group.GroupId -rol "Member"

            if ($createsubjectchannel.IsChecked -eq $true)
            {
                write-Host "De volgende vakken worden toegevoegd:" $vak
                toevoegen_vakken -vakken $vak -groupid $group.GroupId
            }
     

            if ($createprivatechannel.IsChecked -eq $true)
            {
                Write-Host "Leerkrachten worden aan de private channels van leerlingen toegevoegd. Vergeet niet leerlingen zelf toe te voegen - Teams is nog niet actief"
                toevoegen_leden_kanalen -leerlingen $leerlingen -groupid $group.GroupId -leerkrachten $owners
            }

             Write-Host "Klaar! Je kan het opnieuw uitvoeren met een andere .csv of wel het programma afsluiten."

        }

}



$createbutton.Add_Click({totaaluitvoeren})




########################### CLOSE ############################################################################################################


#Show Form
$Form.ShowDialog() | out-null

