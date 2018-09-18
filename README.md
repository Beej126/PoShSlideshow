# Project Description

Photo slideshow implemented in PowerShell -> Windows Forms

Simply target a (nested) folder of images. Local or LAN UNC path supported.
&nbsp;<br/>
&nbsp;<br/>
# Features:
![](https://user-images.githubusercontent.com/6301228/45711128-6e233380-bb3d-11e8-9d5f-adb141b7522a.png)

* **task tray icon** to start slideshow on demand...
* otherwise kicks off after user defined **idle timeout** (honors running video)
* **good randomization** - one soon realizes pleasantly random photos are the key want of a photo slideshow ... fortunately PowerShell has a readily available _random_ commandlet that seems to do quite well
  * persists "lastShown" for each subfolder and avoids re-showing within XX days (currently 1 month)
* image **fade-in and slide** for ambience
* several **hotkeys** functional:
	* <kbd>o</kbd>pen current image folder
	* <kbd>c</kbd>opy current image to _My Photos_
	* <kbd>f</kbd>avorites - add folder to favorites. show more frequently.
	* <kbd>r</kbd>otate current image (and save) - *generally honors EXIF rotation metadata where present, this option allows for manual correction where EXIF is missing*
	* <kbd>u</kbd>pdate folder cache
	* <kbd>d</kbd>debug - show last few random files selected
	* reverse to previously shown photo (<kbd>left cursor</kbd>)
	* pause/play (<kbd>space</kbd>)
	* hotkey legend pops up on any other keypress
* screen click functions:
	* double click in center hides slideshow
	* single click in center pauses slideshow
	* click arrows on far left and right for prev/next image
* skips _.hidden_ folders
* plays videos via VLC
* open to modification - it's just PowerShell :) no compiling tools required

# Install - basically just launch the ps1... here's some tips:
1. only the ps1 and ico files are needed, download them to a folder
2. ensure VLC.exe is in your path
3. (see screenshot below) **create a shortcut** to the ps1 and tweak the target to include ```powershell``` before the ps1 filename... 
2. example full shorcut command line: ```C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -WindowStyle Hidden \\beejquad\Dev\_PersonalProjects\PoShSlideshow\PoShSlideshow.ps1 -photoPath \\beejquad\photos -idleTimeout 2```
1. select Run: <kbd>Minimized</kbd> to make script launch more polished
2. add ```-WindowStyle Hidden``` after powershell.exe on target command line for further polish
1. then hit the <kbd>Advanced</kbd> button and select <kbd>Run as administrator</kbd> - *this is only required for the ``` powercfg /requests``` used to identify running video and avoid starting slideshow after user input idle timeout (wouldn't mind hearing a slicker approach???)*
1. script parameters:
	* add ```-photoPath {path\to\your\images}``` to the end of the shortcut path - UNC shared folder fair game, **write permissions required to persist folder cache flat file**
	* add ```-idleTimeout 2``` to the end of the shortcut path - units are in minutes
1. Copy this shortcut to ```shell:startup``` in Windows FileExplorer to automatically launch this script when you login to your desktop

![](https://user-images.githubusercontent.com/6301228/45711239-c5c19f00-bb3d-11e8-967c-a929a3fe5e35.png)

# Wishlist

* [done] <s>show videos as well - thinking VLC convenient</s>
* add new param to csv list folders always shown (i.e. not subject to "lastShown" exclusion logic)
1. [done] <s>show videos as well - thinking VLC convenient</s>
1. Right mouse to show commands menu same as keyboard
1. Implement a Hide button akin to the forward back buttons
1. Email current photo - on screen keyboard? fire gmail to get contacts
1. [blog request](http://www.beejblog.com/2015/12/powershell-photo-slideshow.html#comment-424): Automatically update folder cache upon new items... to be clear, current approach automatically recognizes new files in existing folders since it only caches the list of folders from which it randomly grabs the next image. Thoughts - Seems pretty straightforward to throw in [PowerShell FileWatcher](http://stackoverflow.com/a/29067433/813599) configured to call the existing `updateFolderCache` function.
