#### Project Description

Photo slideshow implemented in PowerShell -> Windows Forms

Simply target a loca/LAN folder of images.

#### Features:
![](http://2.bp.blogspot.com/-XWNHk4bUjmw/VnZCbx0p_RI/AAAAAAAAR6Y/hjDvvY8mkqE/s1600/Screen%2BShot%2B2015-12-19%2Bat%2B9.36.43%2BPM.png)

* **task tray icon** to start slideshow on demand...
* otherwise kicks off after user defined **idle timeout** (honors running video)
* **good randomization** - one soon realizes pleasantly random photos are the key want of a photo slideshow ... fortunately PowerShell has a readily available _random_ commandlet that seems to do quite well
* image **fade-in and slide** for ambience
* several **hotkeys** functional:
	* <kbd>o</kbd>pen current image folder
	* <kbd>c</kbd>opy current image to _My Photos_
	* <kbd>r</kbd>otate current image (and save) - *specifcally coded to honor EXIF rotation metadata where present but sometimes this gets lost*
	* reverse to previously shown photo (<kbd>left cursor</kbd>)
	* pause/play (<kbd>space</kbd>)
	* hotkey legend pops up on any other keypress
* skips _.hidden_ folders
* open to modification - it's just PowerShell :) no compiling tools required

#### Install - basically just launch the ps1... here's some tips:
1. only the ps1 and ico files are needed, download them to a folder
1. (see screenshot below) **create a shortcut** to the ps1 and tweak the target to include ```powershell``` before the ps1 filename... 
1. select Run: <kbd>Minimized</kbd> to make script launch that much more polished
1. then hit the <kbd>Advanced</kbd> button and select <kbd>Run as administrator</kbd> - *this is only required for the ``` powercfg /requests``` used to identify running video and avoid starting slideshow after user input idle timeout (wouldn't mind hearing a slicker approach???)*
1. script parameters:
	* add ```-photoPath {path\to\your\images}``` to the end of the shortcut path - shared folder fair game
	* add ```-idleTimeout 2``` to the end of the shortcut path - units are in minutes
1. Copy this shortcut to ```shell:startup``` in Windows FileExplorer to automatically launch this script when you login to your desktop

![](http://3.bp.blogspot.com/-fON1N7pNVps/VnbndLY3ipI/AAAAAAAAR7A/p1T8oja9fso/s1600/Screen%2BShot%2B2015-12-20%2Bat%2B9.26.42%2BAM.png)

#### Wishlist

* show videos as well - thinking VLC convenient... would need to pause image display loop and figure out how to recognize when VLC was done playing... would be nice not to keep launching new VLC instance vs reusing 