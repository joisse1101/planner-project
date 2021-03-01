# The Planner
A simple powershell project to generate an A4 planner in powerpoint.

This project is/was an exercise in learning a new language and understanding docs and also a result of being tired of having to throw away papers in my planner every year.

### How to use
1. Download all files to the same folder
2. Edit the colour scheme, fonts and shapes in slide master of planner.pptx
    - Do NOT change position of title placeholders in the slides
    - Do NOT delete the templates_stickers slide
3. Change [int]$year [int]$mth [int]$day variables in main.ps1 to desired first day.
4. In powershell, navigate to the folder where downloaded files are.
5. Run .\main.ps1
    - If powershell has not been configured to run scripts, first configure powershell to run scripts.
6. Save generated .ppt as .pdf, then export to note-taking app of choice.