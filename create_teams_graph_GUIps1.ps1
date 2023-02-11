$timestamp = Get-Date | ForEach-Object { $_ -replace ":", "." }
Start-Transcript -Path "$PSScriptRoot\output_$timestamp.txt"

########################## CONNECTION ####################################

if (-not (Get-Module Microsoft.Graph  -ListAvailable)) {
    Write-Host "Installing module Microsoft.Graph."
    Install-Module Microsoft.Graph -Force
}
else {
    Write-Host "Required module Microsoft.Graph already installed."
}

$scopes = @(
    "TeamsApp.ReadWrite.All",
    "TeamsAppInstallation.ReadWriteForTeam",
    "TeamsAppInstallation.ReadWriteSelfForTeam",
    "TeamSettings.ReadWrite.All",
    "TeamsTab.ReadWrite.All",
    "TeamMember.ReadWrite.All",
    "Group.ReadWrite.All",
    "GroupMember.ReadWrite.All",
    "Directory.ReadWrite.All" ,
    "Channel.Create" , 
    "Directory.Read.All"
)
 
Connect-MgGraph -Scopes $scopes
Select-MgProfile -Name beta
Import-Module Microsoft.Graph.Teams

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

########################### SETTINGS ############################################################################################################

########################### TEAMS

$funSettings = @{ 
    "allowGiphy"            = "false"; 
    "giphyContentRating"    = "strict"; 
    "allowStickersAndMemes" = "false"; 
    "allowCustomMemes"      = "false"; 
}
 
$memberSettings = @{ 
    "allowCreateUpdateChannels"         = "false"; 
    "allowCreatePrivateChannels"        = "false"; 
    "allowDeleteChannels"               = "false"; 
    "allowAddRemoveApps"                = "false"; 
    "allowCreateUpdateRemoveTabs"       = "false"; 
    "allowCreateUpdateRemoveConnectors" = "false"; 
}
 
$guestSettings = @{ 
    "allowCreateUpdateChannels" = "false"; 
    "allowDeleteChannels"       = "false"; 
}
 
$messagingSettings = @{ 
    "allowUserEditMessages"    = "false"; 
    "allowUserDeleteMessages"  = "false" ;
    "allowOwnerDeleteMessages" = "false"; 
    "allowTeamMentions"        = "false"; 
    "allowChannelMentions"     = "false"; 
}

########################### FUNCTIONS ############################################################################################################

function load {
    Import-Csv "$PSScriptRoot\teams.csv" -Header name, teachers, students, subjects | Out-GridView –Title Get-CsvData
}

$checkbutton.Add_Click({ load })

function add_teachers {
    param($teachers, $groupid, $role)
    $usersplit = $teachers -split ";"
    for ($j = 0; $j -le ($usersplit.count - 1) ; $j++) {
        $teamOwner = Get-MgUser -UserId $usersplit[$j]
        New-MgTeamMember -TeamId $groupid -Roles $role `
            -AdditionalProperties @{ 
            "@odata.type"     = "#microsoft.graph.aadUserConversationMember" 
            "user@odata.bind" = "https://graph.microsoft.com/beta/users/" + $teamOwner.Id
        }
    }
}

function add_students {
    param($students, $groupid, $role)
    $usersplit = $students -split ";"
    for ($j = 0; $j -le ($usersplit.count - 1) ; $j++) {
        $teammember = Get-MgUser -UserId $usersplit[$j]
        New-MgTeamMember -TeamId $groupid -Roles $role `
            -AdditionalProperties @{ 
            "@odata.type"     = "#microsoft.graph.aadUserConversationMember" 
            "user@odata.bind" = "https://graph.microsoft.com/beta/users/" + $teammember.Id
        }
    }
}

function add_subjects {
    param($subjects, $groupid)
    $subjectsplit = $subjects -split ";"
    for ($j = 0; $j -le ($subjectsplit.count - 1) ; $j++) {
        New-MgTeamChannel -TeamId $groupid `
            -AdditionalProperties  @{
            "@odata.type"       = "#Microsoft.Graph.channel"
            MembershipType      = "standard"
            IsFavoriteByDefault = $true
            DisplayName         = $subjectsplit[$j]
        }  
    }
}

function add_private_channels {
    param($students, $groupid, $teachers)
    $studentsplit = $students -split ";"
    $teachersplit = $teachers -split ";"
    for ($j = 0; $j -le ($studentsplit.count - 1) ; $j++) {
        $channelname = Get-MgUser -UserId $studentsplit[$j]
        $withprefix = $prefix.Text + $channelname.DisplayName  
        New-MgTeamChannel -TeamId $groupid `
            -AdditionalProperties  @{
            "@odata.type"  = "#Microsoft.Graph.channel"
            MembershipType = "private"
            DisplayName    = $withprefix
            Members        = @(
                @{
                    "@odata.type"     = "#microsoft.graph.aadUserConversationMember"
                    "User@odata.bind" = "https://graph.microsoft.com/beta/users/" + $channelname.Id
                    Roles             = @(
                        "owner"
                    )
                }
            )
        }
        for ($i = 0; $i -le ($teachersplit.count - 1) ; $i++) {
            $channeladd = Get-MgTeamChannel -TeamId $groupid -Filter "StartsWith(DisplayName, '$withprefix')"
            $channeladdteacher = Get-MgUser -UserId $teachersplit[$i]
            $params = @{
                "@odata.type"     = "#microsoft.graph.aadUserConversationMember"
                Roles             = @(
                    "owner"
                )
                "User@odata.bind" = "https://graph.microsoft.com/beta/users/" + $channeladdteacher.Id
            }  
            New-MgTeamChannelMember -TeamId $groupid -ChannelId $channeladd.Id -BodyParameter $params    
        }
    }

}

function total {
    $csv = Import-Csv "$PSScriptRoot\teams.csv"

    foreach ($i in $csv) {
        $nameteam = $i.name
        $owners = $i.teachers
        $students = $i.students
        $subj = $i.subjects

        $paramsteam = @{
            "Template@odata.bind" = "https://graph.microsoft.com/v1.0/teamsTemplates('educationClass')"
            DisplayName           = $nameteam
            FunSettings           = $funSettings
            GuestSettings         = $guestSettings
            MessagingSettings     = $messagingSettings
            MemberSettings        = $memberSettings
        }

        Write-Host "-------------------------------------------------------------------------------------------------"
        Write-Host "The following team will be created:" $nameteam
        New-MgTeam -BodyParameter $paramsteam
        Write-Host "Starting a timeout for 120 seconds to give the system time to complete the creation of each team."
        start-sleep -Seconds 120
        
        $lookupId = Get-MgGroup -Filter "StartsWith(DisplayName, '$nameteam')"

        write-Host "The following teachers are added as owners:" $owners
        add_teachers -teachers $owners -groupid $lookupId.Id -rol "Owner"
        
        Write-Host "The following students are added as members:" $students
        add_students -students $students -groupid $lookupId.Id -rol "Member"

        if ($createsubjectchannel.IsChecked -eq $true) {
            write-Host "The following subjects are added as public channels:" $subj
            add_subjects -subjects $subj -groupid $lookupId.Id
        }

        if ($createprivatechannel.IsChecked -eq $true) {
            Write-Host "Creating private channels for each student, adding each student to his/ her own channel and adding all the teachers to all the private channels."
            add_private_channels -students $students -groupid $lookupId.Id -teachers $owners
        }
    }
    Stop-Transcript
    Write-Host "-------------------------------------------------------------------------------------------------"
    Write-Host "Done! Choose antoher .csv-file or close the tool."
    
}
$createbutton.Add_Click({ total })

########################### CLOSE ############################################################################################################

$Form.ShowDialog() | out-null