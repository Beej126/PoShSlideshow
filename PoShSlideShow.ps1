param(
  [string]$photoPath = "D:\Photos\_Main_Library",
  #$photoPath = "\\BeejQuad\Photos"
  [string]$timeout = 2 #minutes
)

<#
Credit: Initial idea and windows forms screen code from here: https://github.com/adamdriscoll/PoshInternals
#>

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

public class Win32
{
  [DllImport("user32.dll")]
  [return: MarshalAs(UnmanagedType.Bool)]
  public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

  [DllImport("user32")]
  public static extern int SetForegroundWindow(int hwnd);

  [DllImport("user32.dll", CharSet = CharSet.Unicode)]
  public static extern IntPtr FindWindow(String sClassName, String sWindowCaption);

  [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
  private static extern IntPtr FindWindowByCaptionInternal(IntPtr ZeroOnly, string sWindowCaption);
  public static IntPtr FindWindowByCaption(String sWindowCaption)
  {
    return FindWindowByCaptionInternal(IntPtr.Zero, sWindowCaption);
  }
}

//from: http://stackoverflow.com/questions/15845508/get-idle-time-of-machine
public static class UserInput {

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

if ((Get-WmiObject Win32_Process -Filter "Name like 'powershell%' AND CommandLine like '%$(split-path $PSCommandPath -leaf)'").Count -gt 1) {
    #$existingWindow = [win32]::FindWindowByCaption("PoShSlideShow")
    #[Win32]::SetForegroundWindow($existingWindow)
    #[Win32]::ShowWindow($existingWindow, 9 <#SW_SHOW#>)
    exit
}

[Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

[System.WIndows.Forms.Application]::EnableVisualStyles()
try { [System.WIndows.Forms.Application]::SetCompatibleTextRenderingDefault($false) } catch { $Error.Clear() }

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

function showImage() {
    $script:img = $nul

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

    if ($script:pictureBox.Image -ne $null) { 
        $script:pictureBox.Image.Dispose()
    }

    $script:pictureBox.Left = 0
    $script:pictureBox.Top = 0

    $imagePathLabel.Text = $script:randomFile.FullName.Replace("$photoPath\", "")

    $script:pictureBox.Image = $script:img
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

# dock=fill to easily set automatic bounds then remove docking so that pictureBox can be "slid"
#$script:pictureBox.Dock = [System.Windows.Forms.DockStyle]::Fill
#$bounds = $script:pictureBox.Bounds
#$script:pictureBox.Dock = [System.Windows.Forms.DockStyle]::None
$script:pictureBox.Bounds = $screen.Bounds
#$script:pictureBox.Width += 125
#$script:pictureBox.Height += 125

$script:idleTimer = New-Object System.Windows.Forms.Timer
$script:idleTimer.Interval = 30000
$script:idleTimer.add_Tick({

  #$script:notifyIcon.ShowBalloonTip(1, "slideshow", "IdleTime.TotalMinutes: $([UserInput]::IdleTime.TotalMinutes)", [System.Windows.Forms.ToolTipIcon]::Info) 
  
  if ([UserInput]::IdleTime.TotalMinutes -gt $script:timeout) {
  
    #pause timer since this is a long retrieval
    $script:idleTimer.Stop()
    #only start if there's nothing currently running on the display
    #currently checking via heavy "powercfg" CLI call... would be nice to find a lighter approach
    if ((powercfg /requests | select-string "DISPLAY:" -context 0,1).Context.PostContext -eq "None.") { ToggleActive; return }
    $script:idleTimer.Start()
  }
})

<#
$script:frmFade.add_Load({
    #[System.Windows.Forms.Cursor]::Hide()
    $script:frmImage.BringToFront()
    $script:frmFade.BringToFront()
    $script:frmFade.Hide()

    ToggleActive
})
#>

function ToggleActive {
  if ($script:frmFade.Visible) {
    $script:timerAnimate.Stop()
    $script:frmImage.Hide()
    $script:frmFade.Hide()

    $script:idleTimer.Start()
  }
  else {
    $script:idleTimer.Stop()

    $script:frmFade.Opacity = 1
    $script:frmFade.Show()
    $script:frmFade.Activate()
    $script:frmImage.Show()

    $script:animationMode = "showImage"
    $script:timerAnimate.Start()
  }

  $script:enableMenuItem.Text = ("Enable", "Disable")[$script:frmFade.Visible]
}

$commands = {

    $keyEventArgs = $_

    $keycode = [string]($keyEventArgs.KeyCode)
    $keycode = $keycode.Substring(0, [math]::Min(6, $keycode.Length))

    if ($keycode -eq "Volume") {return}

    $script:frmFade.Opacity = 0
    $script:timerAnimate.Stop()

    #debug: $keyEventArgs | Format-List | Out-Host

	#nugget: these windows control event handlers are in a different scope and thus need to reference all other global script variables as $script:var

    switch ($keycode) {
        "Escape" { ToggleActive }
        "O" { Invoke-Item $script:randomFile.DirectoryName; [System.Windows.Forms.Application]::Exit() }

        "C" {
            copy-item $script:randomFile.FullName ($env:HOMEDRIVE+$env:Homepath+"\Pictures\"+($script:randomFile -replace "$photoPath\\", "" -replace "\\", "_"))
            $wshell.Popup("File copied successfully", 1, "Screensaver", 4096 <#TopMost#> + 64 <#Information icon#>)
            $script:timerAnimate.Start()
        }

        "M" { Invoke-Item ($env:HOMEDRIVE+$env:Homepath+"\Pictures\") }

        "R" { $script:pictureBox.Image = $null; $script:img.RotateFlip([Drawing.RotateFlipType]::Rotate90FlipNone); $script:pictureBox.Image = $script:img; $script:img.Save($script:randomFile.FullName); }

        "Left" { if ($script:rewindIndex -gt 0) { $script:randomFile = (gi $script:filesShown[(--$script:rewindIndex)]); showImage}; } #$script:debugLabel.Text = $script:rewindIndex 
        "Right" { if ($script:rewindIndex -lt ($script:filesShown.Count - 1)) { $script:randomFile = (gi $script:filesShown[++$script:rewindIndex]); showImage; } else {$script:timerAnimate.Start();} } #$script:debugLabel.Text = $script:rewindIndex
        "Space" { if ($script:timerAnimate.Enabled) { $script:timerAnimate.Stop() } else {$script:timerAnimate.Start()}; }

        default {
            $wshell.Popup("Keycode: $keycode`n`nESC - Exit`nO - Open folder`rC - [C]opy to 'My Pictures'`rM - Open [M]y Pictures folder`rR - Rotate`rLeft Cursor - Previous image`rRight Cursor - Next image`rSpace - Pause", 3 <#timeout#>, "Usage:", 4096 <#TopMost#> + 64 <#Information icon#>) 
            $script:timerAnimate.Start()
		}
    }
}

$script:frmFade.add_KeyDown($commands)
$script:frmImage.add_KeyDown($commands)

$wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop


if (!(test-path "$photoPath\AllFolders.txt")) {
    dir -Recurse -Directory $photoPath -Exclude .* | select -ExpandProperty FullName | Out-File "$photoPath\AllFolders.txt" #nugget: -ExpandProperty prevents ellipsis on long strings 
}

$folders = gc "$photoPath\AllFolders.txt" -Filter .hidden

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
            do {$script:randomFile = gci ($folders | random) -file -Recurse -Include *.jpg | random } until ($script:randomFile -ne $null)
            $script:filesShown.Add($script:randomFile.FullName)
            $script:rewindIndex = ($script:filesShown.Count - 1)
            #$script:debugLabel.Text = $script:rewindIndex

            showImage
            $script:animationMode = "fadeIn"
            $script:timerAnimate.Start()
        }
        "fadeIn" {
            #fade in the new image
            if ($script:frmFade.Opacity -gt 0) { $script:frmFade.Opacity -= 0.05 }
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

$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Icon = New-Object System.Drawing.Icon "$(Split-Path -parent $PSCommandPath)\icon.ico"
#$notifyIcon.Icon = New-Object System.Drawing.Icon ".\icon.ico"
$notifyIcon.Visible = $true

$notifyIcon.add_MouseDown( { 
  #from: http://stackoverflow.com/questions/21076156/how-would-one-attach-a-contextmenustrip-to-a-notifyicon
  #nugget: ContextMenu.Show() yields a known popup positioning bug... this trick leverages notifyIcons private method that properly handles positioning
  if ($_.Button -ne [System.Windows.Forms.MouseButtons]::Left) {return}
  [System.Windows.Forms.NotifyIcon].GetMethod("ShowContextMenu", [System.Reflection.BindingFlags] "NonPublic, Instance").Invoke($script:notifyIcon, $null)
})

$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$contextMenu.ShowImageMargin = $false
$contextMenu.Show
$notifyIcon.ContextMenuStrip = $contextMenu
$contextMenu.Parent = $notifyIcon
$contextMenu.Items.Add( "Path: $photoPath", $null, $null ) | Out-Null
$contextMenu.Items.Add( "E&xit", $null, { $notifyIcon.Visible = $false; [System.Windows.Forms.Application]::Exit() } ) | Out-Null
$enableMenuItem = $contextMenu.Items.Add( "Enable", $null, { ToggleActive } )

$script:idleTimer.Start()
[System.Windows.Forms.Application]::Run()
if ($Error -ne $null) { pause }
