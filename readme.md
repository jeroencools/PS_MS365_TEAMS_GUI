## to do:
* Simplify the process for adding teams pictures: selecting images, different names ...
* Check for folders with no images
* Add a function to sync it with other (administrative) software

## Working with: 

* 1.19.0               Microsoft.Graph        

![1](https://user-images.githubusercontent.com/113233490/225577157-a825a3e2-4219-4265-9239-a536301dfd9b.png)
![2](https://user-images.githubusercontent.com/113233490/225577188-6ba7c444-3751-42c8-a7e8-019dcc766fe1.png)
![3](https://user-images.githubusercontent.com/113233490/225577200-3007fe93-0b67-47aa-9150-9d4b9c8af470.png)



## Teams_Create_GUI.ps1
You can use this to create "edu teams" based on the information from a .csv-file (teams.csv) in the same directory.

In the csv-file you can provide a team name, owners (teachers), members (students) and channels (subjects). 

By executing this file a transcript "output_*timeanddate*.txt" will be created in the same folder. After the script is done, you can use this to check for errors.
    
### The following actions are always executed:
* Create teams with a custom name

* Add students and teachers

### The following actions are optional - you can choose these in the GUI

* Configure the following settings:

    funsettings =
           
            "allowGiphy"; 
            
            "giphyContentRating"; 
            
            "allowStickersAndMemes"; 
            
            "allowCustomMemes"; 
            
    memberSettings =
      
        "allowCreateUpdateChannels"; 
        
        "allowCreatePrivateChannels"; 
        
        "allowDeleteChannels"; 
        
        "allowAddRemoveApps"; 
        
        "allowCreateUpdateRemoveTabs"; 
        
        "allowCreateUpdateRemoveConnectors"; 
        
    $guestSettings = 
           
           "allowCreateUpdateChannels"; 
           
           "allowDeleteChannels"; 
    messagingSettings
            
            "allowUserEditMessages"; 
            
            "allowUserDeleteMessages";
            
            "allowOwnerDeleteMessages"; 
            
            "allowTeamMentions"; 
            
            "allowChannelMentions"; 
  


* create a public channel for each subject

* create a private channel for each student, add the student to his/her channel and add all the teachers to these channels

* choose a prefix for these private channels (default = "First name Last name", but to make sure they are always on top in the list of channels you can add "0." 
for example "0. First name Last name")
                
*(The creation of these private channels is something that our schools have chosen so you always have an online space for each student that is shared with the teachers. By doing so the teachers can check homework, add comments to files, organise the folders of the students ...)*

* Choose a "welcoming text" for each team. If left empty, no text will be posted. If you add text, each team will have a new post in the "general" channel, posted by the user who executes the script. In the textboxt you can use the variable $name to insert dynamic content in the text (in this example, the name of the team). For example: "Welcome to $name"

* Change the picture of each team. If you enable this feature, you will have to click the button "Create folders in script directory". This will create a folder "images" in the directory of the script containing separate folders for each team. In each folder you can provide an image and the script will set this as the teams picture of the team with the same name. These can be different images for each team. The images must have the name "photo.png". See the following example:

    * Teams_Create_GUI.ps1
    * images
        * team 1
            
            photo.png
        * team 2
            
            photo.png
        * team 3
           
           photo.png
        * team 4
            
            photo.png
   
## teams.csv
  4 headers: name,teachers,students,subjects
  - name = the name of the team you want to create
  - teachers = email addresses of the teachers you want to add - separated by ";"
  - students = email addresses of the students you want to add - separated by ";"
  - subjects = subjects you want to add as channels - separated by ";"
    
