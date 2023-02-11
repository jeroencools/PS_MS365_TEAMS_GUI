Working with: 

* 1.19.0               Microsoft.Graph        

![GUI](https://user-images.githubusercontent.com/113233490/208303660-ab4e8536-d2ed-4551-b997-c44eea87714e.PNG)


## Teams_Create_GUI.ps1
You can use this to create "edu teams" based on the information from a .csv-file (teams.csv) in the same directory.

In the csv-file you can provide a team name, owners (teachers), members (students) and channels (subjects). 
    
### The following actions are always executed:
* Create teams with a custom name

* Add students and teachers

* Configure the following settings:

    funsettings =
           
            "allowGiphy" = "false"; 
            
            "giphyContentRating" = "strict"; 
            
            "allowStickersAndMemes" = "false"; 
            
            "allowCustomMemes" = "false"; 
            
    memberSettings =
      
        "allowCreateUpdateChannels" = "false"; 
        
        "allowCreatePrivateChannels" = "false"; 
        
        "allowDeleteChannels" = "false"; 
        
        "allowAddRemoveApps" = "false"; 
        
        "allowCreateUpdateRemoveTabs" = "false"; 
        
        "allowCreateUpdateRemoveConnectors" = "false"; 
        
    $guestSettings = 
           
           "allowCreateUpdateChannels" = "false"; 
           
           "allowDeleteChannels" = "false"; 
    messagingSettings
            
            "allowUserEditMessages" = "false"; 
            
            "allowUserDeleteMessages" = "false" ;
            
            "allowOwnerDeleteMessages" = "false"; 
            
            "allowTeamMentions" = "false"; 
            
            "allowChannelMentions" = "false"; 
  
### The following actions are optional - you can choose these in the GUI

* create a public channel for each subject

* create a private channel for each student, add the student to his/her channel and add all the teachers to these channels

* choose a prefix for these private channels (default = "First name Last name", but to make sure they are always on top in the list of channels you can add "0." 
for example "0. First name Last name")
                
*(The creation of these private channels is something that our schools have chosen so you always have an online space for each student that is shared with the teachers. By doing so the teachers can check homework, add comments to files, organise the folders of the students ...)*

   
## teams.csv
  4 headers: name,teachers,students,subjects
  - name = the name of the team you want to create
  - teachers = email addresses of the teachers you want to add - separated by ";"
  - students = email addresses of the students you want to add - separated by ";"
  - vakken = subjects you want to add as channels - separated by ";"
    
