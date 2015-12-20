#### Project Description

Photo slideshow implemented in PowerShell -> Windows Forms

#### Features:
* Easy to modify for your own preferences - it's just PowerShell
* Photo folder defined in variable at top of script
* Slideshow can be started on demand via TrayIcon menu...
* and also via user defined timeout (see variable at top of script) - honors running video (via powercfg /requests, which thereby requires **elevated permissions**... created shortcut that "runas administrator") 
* Good randomization - One soon learns that a key feature of a photo screensaver is good randomization.  Fortunately PowerShell has a readily available _random_ commandlet that seems to do quite well.
* Does a fade-in and slide of the image for ambience
* Skips _.hidden_ folders
* Several hotkeys defined - open current image folder, copy current image to _My Photos_, rotate current image (and save), back up to previously shown photo and pause.
