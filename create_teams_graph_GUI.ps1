$timestamp = Get-Date –format ‘MM_dd_yyyy-HH_MM_ss'
Start-Transcript -Path "$PSScriptRoot\output_$timestamp.txt"

########################## CONNECTION ####################################

if (-not (Get-InstalledModule Microsoft.Graph)) {
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
        Title="MainWindow" Height="650" Width="1100">
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
        <TabControl Margin="29,142,29,29" Background="#FF001531">
            <TabItem Header="Create teams">
                <Grid Background="#FF001531">
                    <Button Content="Create teams" Name="createbutton" HorizontalAlignment="Left" Margin="48,280,0,0" VerticalAlignment="Top" Background="White" FontSize="16" FontWeight="Bold" Width="240" Height="40"/>
                    <CheckBox Content="Change the picture of each team?" Name="updateteamsphoto"  FontFamily="Arial" FontSize="16" Foreground="White" FontWeight="Bold" Margin="48,119,0,0" VerticalAlignment="Top" HorizontalAlignment="Left"/>
                    <Button Content="Create folders in script directory" Name="folderbutton" HorizontalAlignment="Left" Margin="350,105,0,0" VerticalAlignment="Top" Background="White" FontSize="16" FontWeight="Bold" Width="440" Height="40"/>
                    <Button Content="Check content of .csv-file" Name="checkbutton" HorizontalAlignment="Left" Margin="48,48,0,0" VerticalAlignment="Top" Background="White" FontSize="16" FontWeight="Bold" Width="240" Height="40"/>
                    <TextBox HorizontalAlignment="Left" Name="welcome" Margin="48,192,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="450" Height="31" FontSize="16" FontWeight="Bold" Grid.Column="1"/>
                    <Label Content="What do you want as a welcome text in each team? &#xA;Use the variable `$name as dynamic content for the team name. " HorizontalAlignment="Left" Margin="520,181,0,0" VerticalAlignment="Top" Foreground="White" FontSize="16" FontWeight="Bold" RenderTransformOrigin="0.069,-0.814" Width="674"/>
                </Grid>
            </TabItem>
            <TabItem Header="Channels and settings">
                <Grid Background="#FF001531">
                    <CheckBox Content="Create public channels for all subjects?" Name="createsubjectchannel" HorizontalAlignment="Left" Margin="48,48,0,0" VerticalAlignment="Top" FontFamily="Arial" FontSize="16" Height="28" Foreground="White" FontWeight="Bold" Width="699" Grid.ColumnSpan="2"/>
                    <CheckBox Content="Create a private channel for each student and add all teachers to those channels?" Name="createprivatechannel" HorizontalAlignment="Left" Margin="48,120,0,0" VerticalAlignment="Top" FontFamily="Arial" FontSize="16" Height="28" Foreground="White" FontWeight="Bold" Width="699" Grid.ColumnSpan="2"/>
                    <Label Content="Default name for these channels is 'first name last name'. Do you want to add a prefix? &#xA;For example: '0. First name Last name' - then write the prefix '0.' in this textbox." HorizontalAlignment="Left" Margin="210,181,0,0" VerticalAlignment="Top" Foreground="White" FontSize="16" FontWeight="Bold" RenderTransformOrigin="0.069,-0.814" Width="674"/>
                    <TextBox HorizontalAlignment="Left" Name="prefix" Margin="48,192,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="122" Height="31" FontSize="16" FontWeight="Bold" Grid.Column="1"/>
                </Grid>
            </TabItem>
            <TabItem Header="Additional settings">
                <Grid Background="#FF001531">
                    <CheckBox Content="Allow stickers and memes?" Name="allowstickers_check"  FontFamily="Arial" FontSize="16" Height="28" Foreground="White" FontWeight="Bold" Width="698" VerticalAlignment="Top" HorizontalAlignment="Left" Margin="48,120,0,0" />
                    <CheckBox Content="Allow giphy?" Name="allowgiphy_check"  FontFamily="Arial" FontSize="16" Foreground="White" FontWeight="Bold" Margin="48,90,0,0" VerticalAlignment="Top" HorizontalAlignment="Left"/>
                    <Label Content="FUNSETTINGS:" HorizontalAlignment="Left" Margin="28,48,0,0" VerticalAlignment="Top" Foreground="White" FontSize="20" FontWeight="Bold"  Width="202"/>
                    <ComboBox Name="giphyrating_check" HorizontalAlignment="Left" Margin="208,89,0,0" VerticalAlignment="Top" Width="120" SelectedIndex="0">
                        <ComboBoxItem Content="strict"/>
                        <ComboBoxItem Content="moderate"/>
                    </ComboBox>
                    <CheckBox Content="Allow custom memes?" Name="allowcustommemes_check"  FontFamily="Arial" FontSize="16" Height="28" Foreground="White" FontWeight="Bold" Width="698" VerticalAlignment="Top" HorizontalAlignment="Left" Margin="48,150,0,0" />
                    <Label Content="MESSAGE SETTINGS:" HorizontalAlignment="Left" Margin="28,180,0,0" VerticalAlignment="Top" Foreground="White" FontSize="20" FontWeight="Bold"  Width="220"/>
                    <CheckBox Content="Allow user to edit messages?" Name="editmessages_check"  FontFamily="Arial" FontSize="16" Foreground="White" FontWeight="Bold" Margin="48,222,0,0" VerticalAlignment="Top" HorizontalAlignment="Left"/>
                    <CheckBox Content="Allow user to delete messages?" Name="deletemessages_check"  FontFamily="Arial" FontSize="16" Foreground="White" FontWeight="Bold" Margin="48,252,0,0" VerticalAlignment="Top" HorizontalAlignment="Left"/>
                    <CheckBox Content="Allow owner to delete messages?" Name="ownerdeletemessages_check"  FontFamily="Arial" FontSize="16" Foreground="White" FontWeight="Bold" Margin="48,282,0,0" VerticalAlignment="Top" HorizontalAlignment="Left"/>
                    <CheckBox Content="Allow team mentions?" Name="teammentions_check"  FontFamily="Arial" FontSize="16" Foreground="White" FontWeight="Bold" Margin="48,312,0,0" VerticalAlignment="Top" HorizontalAlignment="Left"/>
                    <CheckBox Content="Allow channel mentions?" Name="channelmentions_check"  FontFamily="Arial" FontSize="16" Foreground="White" FontWeight="Bold" Margin="48,342,0,0" VerticalAlignment="Top" HorizontalAlignment="Left"/>
                    <Label Content="MEMBERSETTINGS:" HorizontalAlignment="Left" Margin="480,48,0,0" VerticalAlignment="Top" Foreground="White" FontSize="20" FontWeight="Bold"  Width="202"/>
                    <CheckBox Content="Allow members to create and update channels?" Name="createupdatechannels_check"  FontFamily="Arial" FontSize="16" Foreground="White" FontWeight="Bold" Margin="500,90,0,0" VerticalAlignment="Top" HorizontalAlignment="Left"/>
                    <CheckBox Content="Allow members to create private channels?" Name="createprivatechannels_check"  FontFamily="Arial" FontSize="16" Height="28" Foreground="White" FontWeight="Bold" Width="698" VerticalAlignment="Top" HorizontalAlignment="Left" Margin="500,120,0,0" />
                    <CheckBox Content="Allow members to delete channels?" Name="deletechannels_check"  FontFamily="Arial" FontSize="16" Height="28" Foreground="White" FontWeight="Bold" Width="698" VerticalAlignment="Top" HorizontalAlignment="Left" Margin="500,150,0,0" />
                    <CheckBox Content="Allow members to add and remove apps?" Name="removeapps_check"  FontFamily="Arial" FontSize="16" Height="28" Foreground="White" FontWeight="Bold" Width="698" VerticalAlignment="Top" HorizontalAlignment="Left" Margin="500,180,0,0" />
                    <CheckBox Content="Allow members to create, update and remove tabs?" Name="createremovetabs_check"  FontFamily="Arial" FontSize="16" Height="28" Foreground="White" FontWeight="Bold" Width="698" VerticalAlignment="Top" HorizontalAlignment="Left" Margin="500,210,0,0" />
                    <CheckBox Content="Allow members to create, update and remove connectors?" Name="createremoveconnectors_check"  FontFamily="Arial" FontSize="16" Height="28" Foreground="White" FontWeight="Bold" Width="698" VerticalAlignment="Top" HorizontalAlignment="Left" Margin="500,240,0,0" />
                    <Label Content="GUEST SETINGS:" HorizontalAlignment="Left" Margin="480,270,0,0" VerticalAlignment="Top" Foreground="White" FontSize="20" FontWeight="Bold"  Width="202"/>
                    <CheckBox Content="Allow guests to create and update channels?" Name="guestcreateupdatechannels_check"  FontFamily="Arial" FontSize="16" Height="28" Foreground="White" FontWeight="Bold" Width="698" VerticalAlignment="Top" HorizontalAlignment="Left" Margin="500,312,0,0" />
                    <CheckBox Content="Allow guests to delete channels?" Name="guestdeletechannels_check"  FontFamily="Arial" FontSize="16" Height="28" Foreground="White" FontWeight="Bold" Width="698" VerticalAlignment="Top" HorizontalAlignment="Left" Margin="500,342,0,0" />
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@


$reader = (New-Object System.Xml.XmlNodeReader $xaml) 
try { $Form = [Windows.Markup.XamlReader]::Load( $reader ) }
catch { Write-Host "Unable to load Windows.Markup.XamlReader"; exit }
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
            IsFavoriteByDefault = $false
            DisplayName         = $subjectsplit[$j]
        }  
    }
}

function update_teams_photo {
    param($groupid, $name)

    Set-MgTeamPhotoContent -TeamId $groupid -InFile "$PSScriptRoot\images\$name\photo.png"
    
}

function update_welcome_text {
    param($groupid, $name)
    
    $PrimaryChannel = Get-MgTeamPrimaryChannel -TeamId $groupid
    New-MgTeamChannelMessage -TeamId $groupid `
                         -ChannelId $PrimaryChannel.Id `
                         -Body @{
                             Content = $ExecutionContext.InvokeCommand.ExpandString($welcome.Text)
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

function foldercreate
{
    $csv = Import-Csv "$PSScriptRoot\teams.csv" 
    foreach ($i in $csv) 
    {
    $foldername = $i.name

    if (Test-Path -Path "$PSScriptRoot\images\$foldername")
        {
        Write-Host "The folders already exist."
        }
    else
        {
   
        mkdir "$PSScriptRoot\images\$foldername"
        }
    }
    
}

function total {

$funSettings = @{ 
    "allowGiphy"            = $allowgiphy_check.IsChecked; 
    "giphyContentRating"    = $giphyrating_check.Text; 
    "allowStickersAndMemes" = $allowstickers_check.Ischecked; 
    "allowCustomMemes"      = $allowcustommemes_check.Ischecked; 
}
 
$memberSettings = @{ 
    "allowCreateUpdateChannels"         = $createupdatechannels_check.IsChecked; 
    "allowCreatePrivateChannels"        = $createprivatechannels_check.IsChecked; 
    "allowDeleteChannels"               = $deletechannels_check.IsChecked; 
    "allowAddRemoveApps"                = $removeapps_check.IsChecked; 
    "allowCreateUpdateRemoveTabs"       = $createremovetabs_check.IsChecked; 
    "allowCreateUpdateRemoveConnectors" = $createremoveconnectors_check.IsChecked; 
}
 
$guestSettings = @{ 
    "allowCreateUpdateChannels" = $guestcreateupdatechannels_check.IsChecked; 
    "allowDeleteChannels"       = $guestdeletechannels_check.IsChecked; 
}
 
$messagingSettings = @{ 
    "allowUserEditMessages"    = $editmessages_check.IsChecked; 
    "allowUserDeleteMessages"  = $deletemessages_check.IsChecked;
    "allowOwnerDeleteMessages" = $ownerdeletemessages_check.IsChecked; 
    "allowTeamMentions"        = $teammentions_check.IsChecked; 
    "allowChannelMentions"     = $channelmentions_check.IsChecked; 
}

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
        
        if ($updateteamsphoto.IsChecked -eq $true) {
            Write-Host "Updating the picture of each team."
            update_teams_photo -groupid $lookupId.Id -name $nameteam

        }

        if ($welcome.Text -ne $null) {
            Write-Host "Updating the welcome text for each team"
            update_welcome_text -groupid $lookupId.Id -name $nameteam

        }

    }
    Stop-Transcript
    Write-Host "-------------------------------------------------------------------------------------------------"
    Write-Host "Done! Choose another .csv-file or close the tool."
    
}

$createbutton.Add_Click({ total })

$folderbutton.Add_Click({ foldercreate })

########################### CLOSE ############################################################################################################

$Form.ShowDialog() | out-null
