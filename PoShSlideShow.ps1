param(
  [string]$photoPath = "D:\Photos\_Main_Library",
  [string]$idleTimeout = 2 #minutes
)

$Error.Clear()

& $PSScriptRoot\Win32.ps1 #bunch of win32 api exports

#save MainWindowHandle so we can hide explicitly here and keep available for showing at end for troubleshooting, IIF PowerShell errors were emitted 
#nugget: don't use -WindowStyle Hidden on the ps1 shortcut, it prevents retrieval of main MainWindowHandle here...
$process = Get-Process -Id $pid
$poShConsoleHwnd = $process.MainWindowHandle
if ($process.ProcessName -eq "powershell_ise") { $poShConsoleHwnd=0 }
function showPoShConsole {
  param([bool]$show = $true)
  
  if ($show -and [Win32]::IsWindowVisible($poShConsoleHwnd)) { $show=$false }

  [Win32]::ShowWindowAsync($poShConsoleHwnd, @([Win32]::SW_HIDE, [Win32]::SW_SHOWNORMAL)[$show]) | Out-Null
  [Win32]::SetForegroundWindow($poShConsoleHwnd) | Out-Null
}

showPoShConsole $false


<#
Credit: Initial idea and windows forms screen code from here: https://github.com/adamdriscoll/PoshInternals
#>

# bail out if this same script is already running
if ((Get-WmiObject Win32_Process -Filter "Name like 'powershell%' AND CommandLine like '%$(split-path $PSCommandPath -leaf)'").Count -gt 1) {
  exit
}

Add-Type -AssemblyName System.Windows.Forms

[System.WIndows.Forms.Application]::EnableVisualStyles()
try { [System.WIndows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { $Error.Clear() }

#fire up the task tray icon as quickly as possible to give the allusion of progress :)
$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Icon = New-Object System.Drawing.Icon "$(Split-Path -parent $PSCommandPath)\icon.ico"
#$notifyIcon.Icon = New-Object System.Drawing.Icon ".\icon.ico"
$notifyIcon.Visible = $true

function saveFolderCache {
  $folders | export-csv $folderCacheFile -Encoding Unicode #nugget: without -encoding sometimes yielded binary garbage 
}

function cleanExit {
  $notifyIcon.Visible = $false
  #save the updated folder datestamps to be reloaded next time we start up and keep our randomization fresh
  saveFolderCache
  [System.Windows.Forms.Application]::Exit()
  exit
}

#Create a custom form with double buffering... otherwise each image would visibly repaint top to bottom a couple times... assuming because of the labels used for image path etc.
#concept from here: http://www.winsoft.se/category/dotnet/powershell/
try {
Add-Type -WarningAction SilentlyContinue -ReferencedAssemblies @("System.Windows.Forms", "System.Drawing") @'
using System;
using System.Windows.Forms;
//using System.Drawing;
using System.Runtime.InteropServices;

public class DoubleBufferedForm : System.Windows.Forms.Form
{
    public DoubleBufferedForm() {
        DoubleBuffered = true;
        SetStyle(ControlStyles.SupportsTransparentBackColor, true);
    }
}

//from: http://stackoverflow.com/questions/15845508/get-idle-time-of-machine
public static class IdleHelper {

    [DllImport("user32.dll", SetLastError=false)]
    private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

    [StructLayout(LayoutKind.Sequential)]
    private struct LASTINPUTINFO {
        public uint cbSize;
        public int dwTime;
    }

    public static DateTime LastInput {
        get {
            DateTime bootTime = DateTime.UtcNow.AddMilliseconds(-Environment.TickCount);
            DateTime lastInput = bootTime.AddMilliseconds(LastInputTicks);
            return lastInput;
        }
    }

    public static TimeSpan IdleTime {
        get {
            return DateTime.UtcNow.Subtract(LastInput);
        }
    }

    public static int LastInputTicks {
        get {
            LASTINPUTINFO lii = new LASTINPUTINFO();
            lii.cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO));
            GetLastInputInfo(ref lii);
            return lii.dwTime;
        }
    }
}
'@} catch {$Error.Clear()}

function isElevated {
  return (new-object System.Security.Principal.WindowsPrincipal([System.Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
}

function showBalloon { param($message, $icon)
  $script:notifyIcon.ShowBalloonTip(5000, "Slideshow", $message + ("", " - aborting")[$icon -eq "error"], ("Info", $icon)[!!($icon)])
  if ($icon -eq "error") { start-sleep 5; cleanExit }
}

if (!(isElevated)) { showBalloon "Not running under elevated permissions" "error" }

#from here: http://automagical.rationalmind.net/2009/08/25/correct-photo-orientation-using-exif/
$script:rotationMap = @{}
$script:rotationMap["1"] = [Drawing.RotateFlipType]::RotateNoneFlipNone
$script:rotationMap["2"] = [Drawing.RotateFlipType]::RotateNoneFlipX
$script:rotationMap["3"] = [Drawing.RotateFlipType]::Rotate180FlipNone
$script:rotationMap["4"] = [Drawing.RotateFlipType]::Rotate180FlipX
$script:rotationMap["5"] = [Drawing.RotateFlipType]::Rotate90FlipX
$script:rotationMap["6"] = [Drawing.RotateFlipType]::Rotate90FlipNone
$script:rotationMap["7"] = [Drawing.RotateFlipType]::Rotate270FlipX
$script:rotationMap["8"] = [Drawing.RotateFlipType]::Rotate270FlipNone

$script:filesShown = New-Object collections.arraylist
$script:rewindIndex = ($script:filesShown.Count - 1)

function isVideo($fileInfoObj) {
  return $fileInfoObj.Extension -ne ".jpg" -and $fileInfoObj.Extension -ne ".png"
}

function showImage() {
  $script:img = $null
  if ($script:pictureBox.Image -ne $null) { $script:pictureBox.Image.Dispose() }
  $script:pictureBox.Image = $null
  $script:pictureBox.Hide()

  #video?
  if (isVideo($script:randomFile)) {
    $script:frmFade.Opacity = 0.001
    start-process -WindowStyle Hidden -wait -filepath "C:\Program Files (x86)\VideoLAN\VLC\vlc.exe" -ArgumentList "--fullscreen --video-on-top --play-and-exit `"$($randomFile.FullName)`""
    $script:frmFade.Opacity = 1
    $script:frmFade.BringToFront()
    #doesn't give focus to enabled keyboard events - $script:frmFade.Activate()
    [Win32]::SetFocus($script:frmFade.Handle)
    $script:timerAnimate.Start()
  }

  #static image
  else {

    try {
      $script:img = [system.drawing.image]::FromFile($script:randomFile.FullName)
    }
    catch {
      $Error.Clear()
      return
    }
    
    #EXIF Spec: https://web.archive.org/web/20131207065832/http://exif.org/Exif2-2.PDF
    $EXIForientation = 274

    if ($script:img.PropertyIdList -contains $EXIForientation) {
      try { 
        $script:img.RotateFlip($script:rotationMap[[string]$script:img.GetPropertyItem($EXIForientation).Value[0]])
      }
      catch {$Error.Clear()}
    }

    $script:pictureBox.Left = 0
    $script:pictureBox.Top = 0

    $imagePathLabel.Text = $script:randomFile.FullName.Replace("$photoPath\", "")

    $script:pictureBox.Image = $script:img
    $script:pictureBox.Show()
    $script:animationMode = "fadeIn"
  }

}

function addLabel([string] $text) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $text
    $label.BackColor = "Transparent"
    $label.AutoSize = $true
    $label.ForeColor = [System.Drawing.Color]::White
    $label.Font = New-Object System.Drawing.Font -ArgumentList "Segoe UI", 30
    $label.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $label.Parent = $script:frmImage
    return $label
}

$screen = [System.Windows.Forms.Screen]::PrimaryScreen

# another window to cover over the image and fade in/out via opactity
$script:frmFade = New-Object DoubleBufferedForm
$script:frmFade.Name = "frmFade"
$script:frmFade.Text = "PoShSlideShow"
$script:frmFade.Bounds = $screen.Bounds
$script:frmFade.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
$script:frmFade.BackColor = "Black"
$script:frmFade.ShowInTaskbar = $false
$script:frmFade.TopMost = $true

$rightLabel = New-Object System.Windows.Forms.Label
$rightLabel.Font = New-Object System.Drawing.Font("Times New Roman",100,[System.Drawing.FontStyle]::Regular)
$rightLabel.Text = ">"
$rightLabel.ForeColor = "white"
$rightLabel.Parent = $frmFade
$rightLabel.Dock = "left"
$rightLabel.add_Click({forward})

$leftLabel = New-Object System.Windows.Forms.Label
$leftLabel.Font = New-Object System.Drawing.Font("Times New Roman",100,[System.Drawing.FontStyle]::Regular)
$leftLabel.Text = "<"
$leftLabel.ForeColor = "white"
$leftLabel.Parent = $frmFade
$leftLabel.Dock = "left"
$leftLabel.add_Click({reverse})

$btnPause = New-Object System.Windows.Forms.Panel
$btnPause.Parent = $frmFade
$btnPause.Dock = "fill"
$btnPause.add_Click({TogglePause})
$btnPause.add_DoubleClick({ToggleDisplay})

# the main form that holds the photo image content (and labels like the folder name)
$script:frmImage = New-Object DoubleBufferedForm  #System.Windows.Forms.Form
$script:frmImage.Name = "frmImage"
$script:frmImage.Text = "frmImage"
$script:frmImage.Bounds = $screen.Bounds
$script:frmImage.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::None
$script:frmImage.BackColor = "Black" #[System.Drawing.Color]::FromArgb(17, 114, 169)
$script:frmImage.ShowInTaskbar = $false
$script:frmImage.Owner = $script:frmFade

$imagePathLabel = addLabel
$script:debugLabel = addLabel
$script:debugLabel.Visible = $false

$script:pictureBox = New-Object System.Windows.Forms.PictureBox
$script:pictureBox.SizeMode = "Zoom"
#$script:pictureBox.BackgroundImageLayout = "Zoom" # None, Tile, Center, Stretch, Zoom
$script:pictureBox.BackColor = "Black"
$script:pictureBox.Parent = $script:frmImage
$script:pictureBox.Bounds = $frmImage.Bounds
#$script:pictureBox.Width += 125 #a little bigger than screen to allow for sliding animation - caused significant slowdown while sliding... assuming clipping math is more complex
#$script:pictureBox.Height += 125

$script:idleTimer = New-Object System.Windows.Forms.Timer
$script:idleTimer.Interval = 30000
$script:idleTimer.add_Tick({
  if ([IdleHelper]::IdleTime.TotalMinutes -gt $script:idleTimeout) {
      #pause timer since this is a long retrieval
    $script:idleTimer.Stop()
    #only start if there's nothing currently running on the display
    #currently checking via heavy "powercfg" CLI call... would be nice to find a lighter approach
    $powerCfg = powercfg /requests
    if (
      #this catches video streams that don't register as a Display request but do show under System as audio stream playing 
      !($powerCfg | select-string "playing" -quiet) `
      -or ($powerCfg | select-string "DISPLAY:" -context 0,1).Context.PostContext -eq "None." ) { ToggleDisplay; return }
    $script:idleTimer.Start()
  }
})


function ToggleDisplay {
  if ($script:frmFade.Visible) {
    $script:timerAnimate.Stop()
    $script:frmImage.Hide()
    $script:frmFade.Hide()

    if ($script:enableMenuItem.Text -eq "Disable Idle") { $script:idleTimer.Start() }
  }
  else {
    $script:idleTimer.Stop()

    $script:frmFade.Opacity = 1
    $script:frmFade.Show()
    #doesn't set focus to enable keyboard event handlers - $script:frmFade.Activate()
    [Win32]::SetFocus($script:frmFade.Handle)
    $script:frmImage.Show()
    $script:frmFade.BringToFront()

    $script:animationMode = "showImage"
    $script:timerAnimate.Start()
  }

  $script:playMenu.Text = ("Play", "Pause")[$script:playMenu.Text -eq "Play"]; 
}

function pauseAnimation {
  $script:timerAnimate.Stop()
  $script:frmFade.Opacity = 0.05
}

function reverse {
  if ($script:rewindIndex -gt 0) {
    pauseAnimation
    $script:randomFile = (gi $script:filesShown[(--$script:rewindIndex)])
    showImage
  }
  else { showBalloon "at the beginning of image sequence" }
}

function forward {
  if ($script:rewindIndex -lt ($script:filesShown.Count - 1)) {
    pauseAnimation
    $script:randomFile = (gi $script:filesShown[++$script:rewindIndex])
    showImage
  } 
  else {
    $script:animationMode = "showimage"
    $script:timerAnimate.Start()
  }

  #$script:debugLabel.Text = $script:rewindIndex
}

function TogglePause {
  if ($script:timerAnimate.Enabled) { pauseAnimation } else {$script:timerAnimate.Start()};
}

filter IndexPrefix {
  param([string]$delimiter)
  $index++
  "$index$delimiter$($_.replace("$photoPath\", ''))"
}

function getFolders {
  return dir -Directory -Recurse -Exclude .* -Path $photoPath | select -ExpandProperty FullName
}

function createFolderStructs {
  $input | select @{Name="lastShown"; expression={[datetime]0}}, @{Name="path"; expression={$_}}
}

function updateFolderCache {
    showBalloon "Updating folder cache..."
    $script:existingFolders = $script:folders | select -ExpandProperty path
    $script:current = getFolders
    #nugget: "where not exists pattern"
    $script:newFolders = $script:current |? {$script:existingFolders -notcontains $_} 
    $script:deleted = $script:folders |? {$script:current -notcontains $_.path} | select -ExpandProperty path #only for display
    #add newly found folders to the existing collection
    $script:folders += ($script:newFolders | createFolderStructs)
    #delete those in collection not found in physical path by creating new list only including those found (since default pipeline behavior creates unmutable arrays)
    $script:folders = $script:folders |? {$script:current -contains $_.path}

    saveFolderCache
    $script:pathMenu.Text = pathMenuText

    $wshell.Popup(
    "Added:`n   $(@("{none}", ($script:newFolders -join "`n   "))[$script:newFolders.Count -gt 0])`n" +`
    "`nDeleted:`n   $(@("{none}", ($script:deleted -join "`n   "))[$script:deleted.Count -gt 0])`n" +`
    "`nTotal: " + $script:folders.Count
    , 0 <#timeout seconds#>, "Folder Cache Updates:", 4096 <#TopMost#> + 64 <#Information icon#>) 
}

$commands = {

    $keyEventArgs = $_

    $keycode = [string]($keyEventArgs.KeyCode)
    $keycode = $keycode.Substring(0, [math]::Min(6, $keycode.Length))

    if ($keycode -eq "Volume" -or $keycode -like "*win*") {return}

    $wasPaused = -not $script:timerAnimate.Enabled

    pauseAnimation #just makes sense to pause on nearly every keystroke

    #debug: $keyEventArgs | Format-List | Out-Host

	#nugget: these event handlers run in their own separate scope and thus need to reference all other global script variables as $script:var

    switch ($keycode) {
        "Escape" { ToggleDisplay }
        "O" { Invoke-Item $script:randomFile.DirectoryName; ToggleDisplay }

        "C" {
            $destination = $env:HOMEDRIVE+$env:Homepath+"\Pictures\"+$script:randomFile.FullName.Replace("$photoPath\", "").Replace("\", "_")
            copy-item $script:randomFile.FullName $destination
            showBalloon "File copied successfully:`n$destination"
            $script:timerAnimate.Start()
        }

        "M" { Invoke-Item ($env:HOMEDRIVE+$env:Homepath+"\Pictures\"); ToggleDisplay }

        "R" { $script:pictureBox.Image = $null; $script:img.RotateFlip([Drawing.RotateFlipType]::Rotate90FlipNone); $script:pictureBox.Image = $script:img; $script:img.Save($script:randomFile.FullName); }

        "Left" { reverse } 
        "Right" { forward }
        "Space" { if ($wasPaused) {$script:timerAnimate.Start()} }

        "D" {
          $wshell.Popup("last few:`n`n$(($script:filesShown | select -last 20 | IndexPrefix ". ") -join "`n")", 0, "Debug:", 4096 <#TopMost#> + 64 <#Information icon#>)
          $script:timerAnimate.Start()
        }

        "U" { updateFolderCache }

        "F" { $script:currentFolder.lastShown = "1/1/1900"; showBalloon "$($script:currentFolder.path) added to favorites"; saveFolderCache }

        default {
            $wshell.Popup("Keycode: $keycode`n`n" +`
              "ESC - Exit`nO - Open folder`n" +`
              "C - [C]opy to 'My Pictures'`n" +`
              "M - Open [M]y Pictures folder`n" +`
              "R - Rotate`nLeft Cursor - Previous image`n" +`
              "Right Cursor - Next image`n" +`
              "Space - Pause`n" +`
              "D - Debug - list last few showings`n" +`
              "U - update folder cache with additions & deletions - preserves existing timestamps`n" +`
              "F - mark folder as favorite"
              , 3 <#timeout seconds#>, "Usage:", 4096 <#TopMost#> + 64 <#Information icon#>) 
            
            $script:timerAnimate.Start()
		}
    }
}

$script:frmFade.add_KeyDown($commands)

$wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop

$folderCacheFile = "$photoPath\AllFolders_$(("shared", "local")[$photoPath -like "*:*"]).txt"
if (!(test-path $folderCacheFile)) {
  showBalloon "Creating folder cache: $folderCacheFile"
  try {
    #nugget: -ExpandProperty prevents ellipsis on long strings
    $newfolders = getFolders | createFolderStructs
    $folders | Export-Csv $folderCacheFile
  }
  catch {
    showBalloon "unable to create: $folderCacheFile`n`n$Error" "error"
  }
}
else {
  $folders = import-csv $folderCacheFile | select @{Name="lastShown"; expression={[datetime]$_.lastShown}},path
}

#fade and slide timer
$script:timerAnimate = New-Object System.Windows.Forms.Timer
$script:timerAnimate.Interval = 1
$script:timerAnimate.add_Tick({
    
    #$script:debugLabel.Text = $script:animationMode
    #$script:debugLabel.Text = "left: $($script:pictureBox.Left), top: $($script:pictureBox.Top)"

    switch ($script:animationMode) {
        "fadeOut" {
            #bring down the lights by increasing the overlay form's opacity from 0 to 1
            if ($script:frmFade.Opacity -lt 1) { $script:frmFade.Opacity += 0.05 }
            else { $script:animationMode = "showImage" }
        }
        "showImage" {
            $script:timerAnimate.Stop()

            <# _   _               _       _   _            _                __ 
              | | | | ___ _ __ ___( )___  | |_| |__   ___  | |__   ___  ___ / _|
              | |_| |/ _ \ '__/ _ \// __| | __| '_ \ / _ \ | '_ \ / _ \/ _ \ |_ 
              |  _  |  __/ | |  __/ \__ \ | |_| | | |  __/ | |_) |  __/  __/  _|
              |_| |_|\___|_|  \___| |___/  \__|_| |_|\___| |_.__/ \___|\___|_|    
            ######################################################################>

            do {

                #get next random folder... that we haven't seen for XX days... 
                $lookBackDays = -12
                do {
                  $lookBackDays += 2 # progressively relaxing freshness if none found
                  $freshies = $folders | where { $_.lastShown -lt [DateTime]::UtcNow.AddDays($lookBackDays) }
                }
                until ($freshies.Count -gt 0 -or $lookBackDays -eq 0)

                $script:currentFolder = @($folders, $freshies)[$freshies.Count -gt 0] | random #should never happen that max lastShown is somehow greater than "now" but just in case, bail out and pull from the whole list
                if ($script:currentFolder.lastShown -eq "1/1/1900") {} #skip update if this is a favorite, favorite flagged by this special datestamp
                #elseif ($script:currentFolder.path -like "*unfiled*") { $script:currentFolder.lastShown = "1/1/1901" -as [datetime] } #any folder with "unfiled" in the name, never gets datestamped, always "fresh" in the rotation
                else { $script:currentFolder.lastShown = [DateTime]::UtcNow }

                #get next randome file
                #nugget: http://stackoverflow.com/questions/790796/confused-with-include-parameter-of-the-get-childitem-cmdlet
                #there's an interesting interplay between -filter, -include and -recurse...
                #-filter only applies to a single extension, so we need to use -include (less performant since it scans the whole folder but we don't expect that many files per folder)
                #the wildcard on end of path allows the -include to apply to the contents... otherwise the -include operates on the path vs the contents
                #-recurse would basically work without the wildcard tacked on to the path... but then it digs into subfolders undesirably for this use case
                $showFiles = gci -Path "$($script:currentFolder.path)\*" -File -Include @("*.jpg", "*.mp4", "*.mov", "*.avi")
            }
            until ($showFiles.count -gt 0) #account for folders without any viewable files

            $script:randomFile = $showFiles | random

            $script:filesShown.Add($script:randomFile.FullName)
            $script:rewindIndex = ($script:filesShown.Count - 1)
            #$script:debugLabel.Text = $script:rewindIndex

            showImage

            $script:timerAnimate.Start()
        }
        "fadeIn" {
            #fade in the new image
            if ($script:frmFade.Opacity -gt 0.05) { $script:frmFade.Opacity -= 0.05 }
            else {
                $script:slideDirection = random -input (1,-1),(1,1),(-1,1),(-1,-1)
                $script:slideCount = 0
                $script:animationMode = "slide"
            }
        }
        "slide" {
            #slide it for a little ambience
            $script:slideCount += 1
            if ($script:slideCount -lt 250) {
                if ($script:slideCount % 2 -eq 0) {
                    $script:pictureBox.Left += $script:slideDirection[0]
                    $script:pictureBox.Top += $script:slideDirection[1]
                }
            }
            else {
                $script:animationMode = "fadeOut"
            }
            #$script:debugLabel.Text = "Location X,Y, fade: $($script:frmImage.Left), $($script:frmImage.Top), $fade"
        }
    }
})

#open context menu on left mouse for convenience (right mouse works by default)
$notifyIcon.add_MouseDown( { 
  if ($script:contextMenu.Visible) { $script:contextMenu.Hide(); return }
  if ($_.Button -ne [System.Windows.Forms.MouseButtons]::Left) {return}

  #from: http://stackoverflow.com/questions/21076156/how-would-one-attach-a-contextmenustrip-to-a-notifyicon
  #nugget: ContextMenu.Show() yields a known popup positioning bug... this trick leverages notifyIcons private method that properly handles positioning
  [System.Windows.Forms.NotifyIcon].GetMethod("ShowContextMenu", [System.Reflection.BindingFlags] "NonPublic, Instance").Invoke($script:notifyIcon, $null)
})

function pathMenuText {
  return "Path: $script:photoPath ({0} folders)" -f $script:folders.Count
}

$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$contextMenu.ShowImageMargin = $false
$notifyIcon.ContextMenuStrip = $contextMenu
$script:pathMenu = $contextMenu.Items.Add( $(pathMenuText), $null, { updateFolderCache } ) 
$contextMenu.Items.Add( "E&xit", $null, { cleanExit } ) | Out-Null

$enableMenuItem = $contextMenu.Items.Add( "Disable Idle", $null, {
  $script:enableMenuItem.Text = ("Enable Idle", "Disable Idle")[$script:enableMenuItem.Text -eq "Enable Idle"]
  if(!$script:frmFade.Visible -and $script:enableMenuItem.Text -eq "Disable Idle") { $script:idleTimer.Start() }
})

$playMenu = $contextMenu.Items.Add( "Play", $null, { ToggleDisplay } )

$script:idleTimer.Start()
[System.Windows.Forms.Application]::Run()
if ($Error -and $poShConsoleHwnd -ne 0) { showPoShConsole; pause }
