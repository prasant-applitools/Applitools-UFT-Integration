# Applitools-UFT-Integration
UFT Integration using Applitools Browser extension

The important files for integration are :

  a. **applitools_eyes.vbs:**
      Contains the wrapper methods around the Applitools Browser extension object.
      Location: [Project Base Folder/eyes/](https://github.com/prasant-applitools/Applitools-UFT-Integration/blob/main/UFG_POC/eyes/applitools_eyes.vbs)
      
  b. **Applitools_Environment_Variables.xml**
      Contains the Applitools related metadata like, cloud url, Batch details etc.
      Location: [Project Base Folder/eyes/](https://github.com/prasant-applitools/Applitools-UFT-Integration/blob/main/UFG_POC/eyes/Applitools_Environment_Variables.xml)
  
  **Usage Instruction:**
  
  a. Create Page Object:
      ```Set current_page_reference = Browser("name:=Browser Name").Page("title:=Page Title")```

  b. Initialize the eyes Object:
      ' Create Eyes Object
      ```Set eyes = eyes_factory(current_page_reference)```
  
  c. Open Eyes connection:
      ' Open Eyes
      ```eyes.open_eye()```
  
  d. Call the Eyes check commands on eyes object:
      ' Execute check full screen on current page
      ```Set result = eyes.check_full_screen("Landing Page","Strict")```
  
  e. Eyes connection will automatically close when the eyes object is destroyed.
