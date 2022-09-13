 '  Current Page reference
Set current_page_reference = Browser("name:=Insurance and Asset Management worldwide").Page("title:=Insurance and Asset Management worldwide")
current_page_reference.Sync

' Remove the floating icon for cookies settings
current_page_reference.RunScript("document.getElementById('ot-sdk-btn-floating').style = 'display:none';")

' Create Eyes Object
Set eyes = eyes_factory(current_page_reference)

' Open Eyes
eyes.open_eye()

' Execute check full screen on current page
Set result = eyes.check_full_screen("Landing Page","Strict")

' Navigate to next page
current_page_reference.Link("css:=a#about-us").Click
current_page_reference.Sync
current_page_reference.RunScript("document.getElementById('ot-sdk-btn-floating').style = 'display:none';")

' Reusable function call to visit screens and perform full screen checks:
' Check About Us Page and navigate to Economic research Page
check_and_navigate_to_screen  eyes, "About us ", "css:=a#economic_research"

' Check Economic research Page and navigate to Investor Relations Page
check_and_navigate_to_screen eyes, "Economic Research ", "css:=a#investor_relations"

' Check Investor Relations Page and navigate to Press Page
check_and_navigate_to_screen eyes, "Investor Relations ", "css:=a#press"

' Check Press Page and navigate to Sustainability Page
check_and_navigate_to_screen eyes, "Corporate Communication", "css:=a#sustainability"

' Reference to Sustainability Page
Set eyes.set_page_reference = Browser("Corporate Communication").Page("Collaborating for a sustainabl")
 
 Browser("Corporate Communication").Page("Collaborating for a sustainabl").RunScript("document.getElementById('ot-sdk-btn-floating').style = 'display:none';")
 
 ' Execute check full screen on Sustainability Page
 Set result = eyes.check_full_screen("sustainability Page","Strict")
