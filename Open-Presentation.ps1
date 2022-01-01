
<#
.Synopsis
   Opens a Microsoft PowerPoint presentation in Kiosk Mode
.DESCRIPTION
   Long description
.EXAMPLE
   Open-Presentation -Path ./filename.pptx 
#>
function Open-Presentation {
    [CmdletBinding()]
    #    [Alias(OpenPPT)]
    [OutputType([System.Void])]
    Param
    (
        # Param1 help description
        [Parameter(Mandatory = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [ValidateScript({
                if (-Not ($_ | Test-Path) ) {
                    throw "File or folder does not exist"
                }
                if (-Not ($_ | Test-Path -PathType Leaf) ) {
                    throw "The Path argument must be a file. Folder paths are not allowed."
                }
                if ($_ -notmatch "(\.pptx|\.ppt)") {
                    throw "The file specified in the path argument must be either of type .ppt or .pptx"
                }
                return $true 
            })]
        [System.IO.FileInfo]$Path
    )

    Begin {
        # Load Powerpoint Interop Assembly
        [Reflection.Assembly]::LoadWithPartialname("Microsoft.Office.Interop.Powerpoint") > $null
        [Reflection.Assembly]::LoadWithPartialname("Office") > $null

        $ppAdvanceOnTime = 2                      # Advance using preset timers instead of clicks.
        $ppShowTypeKiosk = 3                      # Run in "Kiosk" mode (fullscreen)
        $ppAdvanceTime = 5                        # Amount of time in seconds that each slide will be shown.
        $ppSlideShowPointerType = 4               # Hide the mouse cursor
        $ppSlideShowUseSlideTimings = 2           # specify the mode of advance of the slide show

        $msoFalse = [Microsoft.Office.Core.MsoTristate]::msoFalse
        $msoTrue = [Microsoft.Office.Core.MsoTristate]::msoTrue
    }
    Process {
        # start Powerpoint
        $Application = [Microsoft.Office.Interop.PowerPoint.ApplicationClass]::new() 			# Powerpoint application	
        $Application.Visible = $msoTrue                                                         # Set Visible 
        $Presentation = $Application.Presentations.Open($Path.FullName, $msoTrue) 				# The current powerpoint open

        # Apply powerpoint settings
        $Presentation.SlideShowSettings.AdvanceMode = $ppAdvanceOnTime
        $Presentation.SlideShowSettings.ShowType = $ppShowTypeKiosk

        $Application.ActivePresentation.SlideShowSettings.StartingSlide = 1
        $Application.ActivePresentation.SlideShowSettings.EndingSlide = $Application.ActivePresentation.Slides.Count
        $Application.ActivePresentation.SlideShowSettings.AdvanceMode = $ppSlideShowUseSlideTimings
        $Application.ActivePresentation.SlideShowSettings.LoopUntilStopped = $msoTrue
        $Application.ActivePresentation.SlideShowSettings.ShowType = $ppShowTypeKiosk

        <#
        # another way to write it, 

        $commonParams = @{
        StartingSlide      = '1'
        EndingSlide        = '$Application.ActivePresentation.Slides.Count'
        AdvanceMode        = '$ppSlideShowUseSlideTimings'
        LoopUntilStopped   = '$msoTrue'
        ShowType           = '$ppShowTypeKiosk'
        }

        foreach($key in $commonParams.keys)
            {
                $message = '$Application.ActivePresentation.SlideShowSettings.{0} = {1}' -f $key, $commonParams[$key]
                $message
            }
        
        #>

        # Apply settings to each slide in ForEach loop
        ForEach ($s In $Application.ActivePresentation.Slides) {
            $s.SlideShowTransition | ForEach-Object {
                $_.AdvanceOnTime = $msoTrue
                $_.AdvanceTime = $ppAdvanceTime
                return
            }
        }
        $Presentation.SlideShowSettings.Run()
        $Presentation.SlideShowSettings.Run().view.PointerType = $ppSlideShowPointerType        # Attempt to hide mouse cursor 
    }
    End {
        # Write-Host (Resolve-Path $Path)
        # Write-Debug "End Block"
    }
}




######################################

function with {

 param(

  [Parameter(Mandatory = $true, Position = 0, ValueFromPipeLine = $true)]

  [Object]

  $Object,

  [Parameter(Mandatory = $true, Position = 1)]

  [String]

  $Block

 )

 begin {

  $code = $Block -replace '(?m)^\s*(?=\.)', '$Object'

 }

 process {

  [ScriptBlock]::Create($code).Invoke()

 }

}

 
 # ...or use the function

with ($Application.ActivePresentation.SlideShowSettings) @'
        .StartingSlide      = 1
        .EndingSlide        = $Application.ActivePresentation.Slides.Count
        .AdvanceMode        = $ppSlideShowUseSlideTimings
        .LoopUntilStopped   = $msoTrue
        .ShowType           = $ppShowTypeKiosk
'@




#######################################


$commonParams = @{
        StartingSlide      = '1'
        EndingSlide        = '$Application.ActivePresentation.Slides.Count'
        AdvanceMode        = '$ppSlideShowUseSlideTimings'
        LoopUntilStopped   = '$msoTrue'
        ShowType           = '$ppShowTypeKiosk'
}

foreach($key in $commonParams.keys)
{
    $message = '$Application.ActivePresentation.SlideShowSettings.{0} = {1}' -f $key, $commonParams[$key]
    Write-Output $message
}