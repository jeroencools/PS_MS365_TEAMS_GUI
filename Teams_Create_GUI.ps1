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
$reader = (New-Object System.Xml.XmlNodeReader $xaml) 
try { $Form = [Windows.Markup.XamlReader]::Load( $reader ) }
catch { Write-Host "Unable to load Windows.Markup.XamlReader"; exit }
# Store Form Objects In PowerShell
$xaml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name) }

########################### FUNCTIONS ############################################################################################################

function load {
    Import-Csv "$PSScriptRoot\teams.csv" -Header name, teachers, students, subjects | Out-GridView –Title Get-CsvData
}

$checkbutton.Add_Click({ load })

function add_teachers {
    param($teachers, $groupid, $role)
    $usersplit = $teachers -split ";"
    for ($j = 0; $j -le ($usersplit.count - 1) ; $j++) {
        Add-TeamUser -GroupId $GroupId -User $usersplit[$j] -Role $role
    }
}

# add students and optionally create private channels with a chosen prefix
function add_students {
    param($students, $groupid, $role)
    $usersplit = $students -split ";"
    for ($j = 0; $j -le ($usersplit.count - 1) ; $j++) {

        Add-TeamUser -GroupId $GroupId -User $usersplit[$j] -Role $role
        $channelname = Get-AzureADUser -ObjectId $usersplit[$j] | Select Displayname
        $withprefix = $prefix.Text + $channelname.DisplayName         
        if ($createprivatechannel.IsChecked -eq $true) {
            New-TeamChannel -GroupId $GroupId -DisplayName $withprefix -MembershipType Private
        }
        Start-Sleep -Seconds 0.5    
    }
}
    
# add teachers to private channels - students need to be added manually after team has been activated
function add_to_channels {
    param($students, $groupid, $teachers)
    $usersplit = $students -split ";"
    $usersplitteachers = $teachers -split ";"
    for ($j = 0; $j -le ($usersplit.count - 1) ; $j++) {
        $channelname = Get-AzureADUser -ObjectId $usersplit[$j] | Select Displayname
        $withprefix = $prefix.Text + $channelname.DisplayName        
        for ($i = 1; $i -le ($usersplitteachers.count - 1) ; $i++) {
            Add-TeamChannelUser -GroupId $GroupId -DisplayName $withprefix -User $usersplitteachers[$i]
            Start-Sleep -Seconds 0.5 
            Add-TeamChannelUser -GroupId $GroupId -DisplayName $withprefix -User $usersplitteachers[$i] -Role Owner
        }                
    }
}

    
# add subject channels
function add_subjects {
    param($subjects, $groupid)
    $subjectsplit = $subjects -split ";"
    for ($j = 0; $j -le ($subjectsplit.count - 1) ; $j++) {
        if ($createsubjectchannel.IsChecked -eq $true) {
            New-TeamChannel -GroupId $GroupId -DisplayName $subjectsplit[$j]
        }
    }
}

# combination of all functions
function total {
    $csv = Import-Csv "$PSScriptRoot\teams.csv"

    foreach ($i in $csv) {
        $nameteam = $i.name
        $owners = $i.teachers
        $students = $i.students
        $vak = $i.subjects
            
        Write-Host "-------------------------------------------------------------------------------------------------"
            
        Write-Host "The following team will be created:" $nameteam
        $group = New-Team -MailNickname $nameteam -displayname $nameteam -Template "EDU_Class"
        Set-Team -GroupId $group.GroupId -AllowAddRemoveApps $false -AllowCreateUpdateChannels $false -AllowCreateUpdateRemoveTabs $false -AllowDeleteChannels $false -AllowUserDeleteMessages $false -AllowUserEditMessages $false
 
        write-Host "The following teachers are added as owners:" $owners
        add_teachers -teachers $owners -groupid $group.GroupId -rol "Owner"

        Write-Host "The following students are added as members:" $students
        add_students -students $students -groupid $group.GroupId -rol "Member"

        if ($createsubjectchannel.IsChecked -eq $true) {
            write-Host "The following subjects are added as public channels:" $vak
            add_subjects -subjects $vak -groupid $group.GroupId
        }
     

        if ($createprivatechannel.IsChecked -eq $true) {
            Write-Host "Teachers are added to the private student channels. Don't forget to add each student to the corresponding channel manually after the team has been activated."
            add_to_channels -students $students -groupid $group.GroupId -teachers $owners
        }

        
    }
    
    Write-Host "-------------------------------------------------------------------------------------------------"
    Write-Host "Done! Choose antoher .csv-file or close the tool."
}

$createbutton.Add_Click({ total })

########################### CLOSE ############################################################################################################

$Form.ShowDialog() | out-null

