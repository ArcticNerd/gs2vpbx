# David Joy
# V 1.1
# April 4th 2023

Function Get-FileName($initialDirectory)
{  
 [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) |
 Out-Null
 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = “All files (*.*)| *.*”
 $OpenFileDialog.Title = "Open CSV file for GS Ext Export"
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
} #end function Get-FileName

$SelectedFile = Get-FileName -initialDirectory “c:fso”

if ($SelectedFile -eq "") {
# File location for the exported grandstream extensions (And get item for future code)
#$locExportedGrandstreamExtensions = "C:\Temp\CodeStuff\GS_PBX_blnx_test_export_sip_extensions.csv"
#Write-Host "Cancelled"
Exit
} else {
$locExportedGrandstreamExtensions = $SelectedFile
}
$locExportedGrandstreamGetItem = Get-Item $locExportedGrandstreamExtensions

$locationInputFile = $locExportedGrandstreamGetItem.DirectoryName
$defaultFileName = $locExportedGrandstreamGetItem.BaseName + ".VPBXINPUTFILE.csv"

function Save-File([string] $initialDirectory){
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
 $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
 $SaveFileDialog.initialDirectory = $locationInputFile
 $SaveFileDialog.filter = "CSV files (*.csv)| *.csv"
 $SaveFileDialog.Title = "Output file name"
 $SaveFileDialog.FileName = $defaultFileName
 $SaveFileDialog.ShowDialog() |  Out-Null
 return $SaveFileDialog.filename
}

# *** Entry Point to Script ***
# File Output:
$SaveFile=Save-File $locationInputFile

if ($SaveFile -ne "") 
{
   $locOutputCSV = $SaveFile
} else {
   $locOutputCSV = $locExportedGrandstreamGetItem.Directory.FullName + "\" + $locExportedGrandstreamGetItem.BaseName + ".VPBXINPUTFILE.csv"
}

# Input the grandstream extensions to the $inputCSV variable, sorting it by the extensions (lowest to highest)
$inputCSV = Import-csv -Path $locExportedGrandstreamExtensions | Sort-Object -Property {[int]$_.Extension}

## CSV Template

$colVar1 = "mode"
$colVar2 = "extension"
$colVar3 = "ext_name"
$colVar4 = "language"
$colVar5 = "class_of_service"
$colVar6 = "technology"
$colVar7 = "Profile_Name"
$colVar8 = "device_user"
$colVar9 = "device_password"
$colVar10 = "device_description"
$colVar11 = "devices_emergency_cid_name"
$colVar12 = "devices_emergency_cid_number"
$colVar13 = "channel"
$colVar14 = "virtual_number"
$colVar15 = "ring_device"
$colVar16 = "codecs"
$colVar17 = "max_contacts"
$colVar18 = "features_password"
$colVar19 = "email"
$colVar20 = "did_number"
$colVar21 = "cid_number"
$colVar22 = "call-limit"
$colVar23 = "call_waiting"
$colVar24 = "vm_enabled"
$colVar25 = "vm_password"
$colVar26 = "saycid"
$colVar27 = "sayduration"
$colVar28 = "envelope"
$colVar29 = "attach"
$colVar30 = "delete"
$colVar31 = "ask_password"
$colVar32 = "skip_instructions"
$colVar33 = "outgoing_rec"
$colVar34 = "incoming_rec"
$colVar35 = "external_cid_name"
$colVar36 = "external_cid_number"
$colVar37 = "emergency_cid_name"
$colVar38 = "emergency_cid_number"
$colVar39 = "dial_profile"
$colVar40 = "accountcode"
$colVar41 = "followme_numbers"
$colVar42 = "initial_ringtime"
$colVar43 = "fw_ringtime"
$colVar44 = "ring_strategy"
$colVar45 = "followme-enabled"
$colVar46 = "recname"
$colVar47 = "enable_callee_prompt"
$colVar48 = "internal_numbers_confirmation"
$colVar49 = "dynamic_queues"
$colVar50 = "static_queues"
$colVar51 = "mobile_number"
$colVar52 = "home_number"
$colVar53 = "organization"
$colVar54 = "job_title"
$colVar55 = "send_welcome_email"
$colVar56 = "generate_qr"

## CSV Responses

$resYes = "yes"
$resNo = "no"
$resDefault = ""

$resMode = "add"
$resLang = "en"
$resClass = "all"
$resTechnology = "pjsip"
$resTechProfile = $resDefault
$resDevicePass = $resDefault
$resDeviceEmergName = $resDefault
$resDeviceEmergCID = $resDefault
$resFXSChannel = $resDefault
$resVirtualNumber = $resDefault
$resRingDevice = $resDefault
$resFeaturePass = $resDefault
$resEmail = $resDefault
$resDID = $resDefault
$resCID = $resDefault
$resCallLimit = "0"
$resCallWaiting = "1"
$resEnableVM = $resYes
$resVMPass = $resDefault
$resSayCID = $resYes
$resDayDur = $resYes
$resEnvelope = $resYes
$resAttach = $resYes
$resDeleteVM = $resNo
$resAskPass = $resYes
$resSkip = $resNo
$resRecOut = $resNo
$resRecIn = $resNo
$resExternalCIDName = $resDefault
$resExternalCIDNum = $resDefault
$resEmergCID = $resDefault
$resEmergNum = $resDefault
$resDialProfile = "Default"
$resAccountCode = $resDefault
$resCodecs = "g722|g723|ulaw|alaw"
$resMaxContact = "1"
$resFollowMeNum = $resDefault
$resInitialRingTime = "0"
$resFWRingTime = "0"
$resRingStrat = "one_by_one"
$resFollowMeActive = $resNo
$resRecName = $resNo
$resCalleePrompt = $resNo
$resIntNumConf = $resNo
$resDynamicQueues = $resDefault
$resStaticQueues = $resDefault
$resMobileNum = $resDefault
$resHomeNum = $resDefault
$resOrg = $resDefault
$resJobtitle = $resDefault
$resSendWelcomeemail = $resDefault
$resGenerateQR = $resDefault

## DO NOT EDIT BELOW THIS LINE

# Create blank hashtables
$mergedData = @()
$workingData = @{}

# Input the grandstream extensions to the $inputCSV variable, sorting it by the extensions (lowest to highest)
$inputCSV = Import-csv -Path $locExportedGrandstreamExtensions | Sort-Object -Property {[int]$_.Extension}

foreach ($row in $inputCSV) {

$resExt = $row.Extension
$resName = $row."First Name" + " " + $row."Last Name"

# If Statement if Email exists
if ($row.'Email Address' -eq "") {
$resEmail = $resDefault
} else {
$resEmail = $row.'Email Address'
}

# If Statement if Voicemial Pass Exists
if ($row.'Voicemail Password'.Length -eq 8) {
$resVMPass = $resDefault
} else {
$resVMPass = $row.'Voicemail Password'
}

$workingData = New-Object PSObject -Property @{
    $colVar1 = $resMode
    $colVar2 = $resExt
    $colVar3 = $resName
    $colVar4 = $resLang
    $colVar5 = $resClass
    $colVar6 = $resTechnology
    $colVar7 = $resTechProfile
    $colVar8 = $resExt
    $colVar9 = $resDevicePass
    $colVar10 = $resName
    $colVar11 = $resDeviceEmergName
    $colVar12 = $resDeviceEmergCID
    $colVar13 = $resFXSChannel
    $colVar14 = $resVirtualNumber
    $colVar15 = $resRingDevice
    $colVar16 = $resCodecs
    $colVar17 = $resMaxContact
    $colVar18 = $resFeaturePass
    $colVar19 = $resEmail
    $colVar20 = $resDID
    $colVar21 = $resCID
    $colVar22 = $resCallLimit
    $colVar23 = $resCallWaiting
    $colVar24 = $resEnableVM
    $colVar25 = $resVMPass
    $colVar26 = $resSayCID
    $colVar27 = $resDayDur
    $colVar28 = $resEnvelope
    $colVar29 = $resAttach
    $colVar30 = $resDeleteVM
    $colVar31 = $resAskPass
    $colVar32 = $resSkip
    $colVar33 = $resRecOut
    $colVar34 = $resRecIn
    $colVar35 = $resExternalCIDName
    $colVar36 = $resExternalCIDNum
    $colVar37 = $resEmergCID
    $colVar38 = $resEmergNum
    $colVar39 = $resDialProfile
    $colVar40 = $resAccountCode
    $colVar41 = $resFollowMeNum
    $colVar42 = $resInitialRingTime
    $colVar43 = $resFWRingTime
    $colVar44 = $resRingStrat
    $colVar45 = $resFollowMeActive
    $colVar46 = $resRecName
    $colVar47 = $resCalleePrompt
    $colVar48 = $resIntNumConf
    $colVar49 = $resDynamicQueues
    $colVar50 = $resStaticQueues
    $colVar51 = $resMobileNum
    $colVar52 = $resHomeNum
    $colVar53 = $resOrg
    $colVar54 = $resJobtitle
    $colVar55 = $resSendWelcomeemail
    $colVar56 = $resGenerateQR}

$mergedData += $workingData

}

$mergedData | Select-Object $colVar1,$colVar2,$colVar3,$colVar4,$colVar5,$colVar6,$colVar7,$colVar8,$colVar9,$colVar10,$colVar11,$colVar12,$colVar13,$colVar14,$colVar15,$colVar16,$colVar17,$colVar18,$colVar19,$colVar20,$colVar21,$colVar22,$colVar23,$colVar24,$colVar25,$colVar26,$colVar27,$colVar28,$colVar29,$colVar30,$colVar31,$colVar32,$colVar33,$colVar34,$colVar35,$colVar36,$colVar37,$colVar38,$colVar39,$colVar40,$colVar41,$colVar42,$colVar43,$colVar44,$colVar45,$colVar46,$colVar47,$colVar48,$colVar49,$colVar50,$colVar51,$colVar52,$colVar53,$colVar54,$colVar55,$colVar56 | Export-Csv -Path $locOutputCSV -NoTypeInformation

Exit