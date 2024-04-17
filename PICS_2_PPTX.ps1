<###################################################################################
    ################################################################################
    ////////////////////////////////////////////////////////////////////////////////
    //////   ____  __  ___  ____    ____    ____  ____  ____  _  _   ///////////////
    //////  (  _ \(  )/ __)/ ___)  (___ \  (  _ \(  _ \(_  _)( \/ )  ///////////////
    //////   ) __/ )(( (__ \___ \   / __/   ) __/ ) __/  )(   )  (   ///////////////
    //////  (__)  (__)\___)(____/  (____)  (__)  (__)   (__) (_/\_)  ///////////////
    //////                                                           ///////////////
    ////////////////////////////////////////////////////////////////////////////////
    ################################################################################


    .Synopsis
    Creates an PowerPointPresentation like a Diashow.
    .Description
    This Script creates a blank presentation and insert your pictures.



####################################################################################>



## <<< Define/insert the Image-PATH >>> ##

Write-Host -f green "Hello! Please insert your Image-Directory:"
$dirMaterial = Read-Host


## <<< Array && ENV-Variables >>> ##

$pathPics = Get-ChildItem $dirMaterial
$picArray = @()
$picArray += @($pathPics.Name)
$rtLength = @($picArray).Count
$loopVal = 0


## <<< Create a blank PPTX >>> ##

Write-Host -ForegroundColor Green 'Creating PPTX...'
add-type -assembly microsoft.office.interop.powerpoint
$Application = New-Object -ComObject powerpoint.application
$application.visible = [Microsoft.Office.Core.MsoTriState]::msoTrue
$slideType = "microsoft.office.interop.powerpoint.ppSlideLayout" -as [type]
$emptylayout = $slideType::ppLayoutBlank
$presentation = $application.Presentations.add()


## <<< Creating the correct amount of blank slides >>> ##

while ($loopVal -lt $rtLength)
{
  $slide = $presentation.slides.add(1,$emptylayout)
  $loopVal ++
  Start-Sleep -Seconds 0.5 #Sleep-Timer depends on your pc performance - alternatively you can comment this line
}

$blankPath = $dirMaterial + '/PIC2PPTX.pptx'
$presentation.SaveAs($blankPath)
$presentation.Close()

Write-Host -f green "BEEP BE BU BEEEP! Now I insert your images..."


## <<< Insert amount of images in blank slides >>> ##

$Powerpoint = New-Object -ComObject Powerpoint.Application
$Presentation1 = $Powerpoint.Presentations.Open($blankPath,0)

$loopVal = 0

while ($loopVal -lt $rtLength )
{
  $Slide = $Presentation1.Slides($loopVal + 1)
  $yourImageName = $dirMaterial + '\' + $picArray[$loopVal]
  $Slide.Shapes.AddPicture($yourImageName, 0, 1, 0, 0)
  $yourImageName
  $loopVal++
  Start-Sleep -Seconds 0.5 #Sleep-Timer depends on your pc performance - alternatively you can comment this line
}

$picArray = $null


$Presentation1.SaveAs($blankPath)
$Powerpoint.Quit()

Write-Host -ForegroundColor Green 'Done! Happy presenting :)'
