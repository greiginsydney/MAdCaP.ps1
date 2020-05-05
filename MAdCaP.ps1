<#  
<#  
.SYNOPSIS  
	This script provides a GUI administrative interface for Lync's Analog Devices & Common Area Phones. It takes no command-line parameters.

.DESCRIPTION  
	MAdCaP excuses the administrator from having to remember the syntax for the creation of Analog Devices and Common Area Phones.
	
	It also captures the installation's existing parameters (SIP domains, Gateway FQDNs, Dial Plans and Policies), enabling them to be easily selected.

.NOTES  
    Version      	   	: 2.1
	Date			    : 9th June 2018 
	Lync Version		: 2010, 2013 & SfB2015
    Author    			: Greig Sheridan
	Header stolen from  : Pat Richard's amazing "Get-CsConnections.ps1"

	WISH-LIST / TODO:
					- ??
	
	Revision History	
					v2.1 - 9th June 2018
						Rearranged calls to "handler_ValidateGo" to fix where the Go button wasn't lighting/going out
						Corrected errors in the New Object DN and OU popup help text
	
					v2.0 - 29th April 2018
						Incorporated my version of Pat's "Get-UpdateInfo". Credit: https://ucunleashed.com/3168
						Added test for AD module to prevent re-loading unnecessarily  
						Suppressed lots of "loading" noise from verbose output with "-verbose:`$false"
						Stripped the "Tag:" name from the start of the relevant policies (kinda redundant, and was getting in the way of below)
						Updated the Existing Objects tab: the Policies update in real-time to show the selected object's values
						Replaced the "Browse" button on the Existing Objects tab with a new "Filter" button and form
						Improved efficiency:
							- The "Refresh" button on the Existing items tab (function "Update-DeviceList") now reads ADs & CAPs into separate global arrays
							- "Update-Display" now just reads the item directly from the relevant array rather than re-querying
							- "Update-DeviceList" no longer fires if the user Cancels or makes no change on the OU / Browse form
							- "Grant-Policy" now checks the existing and new policy values & skips the commands that would make no change
						Peppered "write-progress" throughout the loading process to help debugging
						Added handling for "-debug" switch for in-depth debug display
						Corrected tab order on the Existing items tab
						Added "-ShowExisting" switch so you can launch with that tab selected
	
					v1.9 - 29th September 2017
						Fixed bug where selecting the "Existing Objects" tab on launch (without doing anything else) didn't pick up the default.
						
					v1.8 - 22nd September 2017
						Added the OU picker to the "Existing Objects" tab
						Added a "Select & Make Default" button to the OU picker & associated code to save/read config
						Saved the chosen RegistrarPool & SIPDomain to the new config file
						Corrected minor whoops: the OK message after creating a new device referenced the "Existing Item" tab but it's correctly the "Existing Object" tab
						Revised  OU TextBox & DN TextBox to remove the code that moves the cursor to the end of the line with every character
						Neatened OU TextBox & DN TextBox to remove duplicate calls to handler_ValidateGo
					
					v1.7 - 27th June 2017
						MAJOR UPDATE / REFRESH:
						Added the "OU" picker stolen from Anthony Caragol's brilliant Lync_Common_Area_Phone_Tool_v1_2.ps1
							(https://gallery.technet.microsoft.com/Lync-2013-Common-Area-57bc4ff1)
							Added the "sticky" enhancement so it opens to the previously selected OU. Tip: Copy/Paste from an Existing object on the other tab!
						Rearranged the tabs: Policies are now only visible for Existing objects
							(timing problems caused real headaches trying to set a Policy or PIN when you initially create the object)
						Disabled the PIN text box on the "Existing Object" tab if you've ONLY selected an Analog Device
						Enabled "DN" for AnalogDevices (previously not an option - was that old Lync 2010 behaviour or a coding error?)
						Script now lets you specify a DN and an OU - an invalid pairing - but will disregard/ignore the OU & use the DN
						Changed script to (re)populate the AD & CA phone lists each time you select the "Existing Item" tab - where it previously did this on script load
						When an object is created on the New tab and you select the Existing tab the script now pre-selects that object (if it exists when we query for it)
							- if it doesn't and you press Refresh, it will auto-select the object just created
						Added a popup MsgBox to indicate success/fail after creating an object on the New tab
						Added "-warningaction silentlycontinue" liberally to suppress the yellow that sprays in the underlying P$ window if you've been deleting policies, etc 
							that are still assigned to users or devices
							
					v1.6 - 1st March 2016
						Added Location Policy to the policies you can set. (Thank you @JohnACook)
						
					v1.5 - 28th Dec 2014
						Signed the script with my code-signing certificate (thanks DigiCert!)
						Changed "$NewDisplayNumberTextBox.Add_TextChanged" to accept a dash as valid
								
					v1.4 - 1st Nov 2013
						Added quotes around Gateway, Line URI & PIN before sending them to Lync
						
					v1.3 - 11th June 2013 
						Corrected bug where I was incorrectly sending "DN=" instead of "CN=" to create a CommonAreaPhone referencing an existing object

					v1.2 - 26th April 2013
						Added quotes around registrar in "$GoButton.Add_Click" as FQDNs with '-' were being rejected and raising errors.
						
					v1.1 - 5th March 2013
						Revised "$NewLineUriTextBox.Add_TextChanged" to support ";ext="
						
					v1.0 - 29th December 2012
						Initial release
	
.LINK  
    https://greiginsydney.com/madcap-ps1-a-gui-for-lync-analog-devices-common-area-phones

.PARAMETER SkipUpdateCheck
		Boolean. Skips the automatic check for an Update. Courtesy of Pat: http://www.ucunleashed.com/3168			
		
.PARAMETER ShowExisting
		Boolean. Launch with the Existing Items tab selected
	
#>

[CmdletBinding(SupportsShouldProcess = $False)]
Param(
	[switch] $SkipUpdateCheck,
	[switch] $ShowExisting
)

$ScriptVersion = "2.1"
$Error.Clear()          #Clear PowerShell's error variable
$Global:Debug = $psboundparameters.debug.ispresent

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
if (!(Get-Module ActiveDirectory))
{
	try
	{
		Import-Module ActiveDirectory -erroraction Stop| out-null	 # For Windows 2008 Support
	}
	catch 
	{
		Write-Warning 'Script is exiting. Failed to load AD Module. (Add "-debug" switch for more info)'
		if ($Global:Debug)
		{				
			$Global:error | fl * -f #This dumps to screen as white for the time being. I haven't been able to get it to dump in red
		}
		exit
	}
}
$global:Icon = [system.drawing.icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe")
$global:LastTouchedObject = ""
$global:AddCAPselectedOU = ""
$Global:DefaultOU = ""
$Global:DefaultPool = ""
$Global:DefaultSIPDomain = ""
$scriptpath = $MyInvocation.MyCommand.Path
$Configfile = ([System.IO.Path]::ChangeExtension($scriptpath, "xml"))
$Global:ADs = @()
$Global:CAPs = @()

#Values used in the Filter form:
$Global:FilterOu = ""
$Global:FilterSipText = ""
$Global:FilterLineUriText = ""
$Global:FilterDisplayNameText = ""
$Global:FilterRegistrarPool = ""

write-progress -id 1 -Activity "Loading" -Status "Querying SIP domains and Registrars" 
$Global:RegistrarList = Invoke-Expression "get-csservice -registrar -verbose:`$false"
$Global:SipDomains = Invoke-Expression "Get-CsSipDomain -verbose:`$false"
write-progress -id 1 -Activity "Loading" -Status "Querying SIP domains and Registrars" -Complete

# ============================================================================
# START FUNCTIONS ============================================================
# ============================================================================ 

Function Add-QuickOU-Node($Nodes, $Path, $DefaultSelectedOU)
{
	$OUArray=$Path.Split(",")
	[array]::Reverse($OuArray)
	$SelectPath=""

	$OuArray | %{
		if ($SelectPath.length -eq 0) {$SelectPath=$_} else {$SelectPath = $_ + "," + $SelectPath}
		$FindIt = $Nodes.Find($_, $False)
		if ($FindIt.Count -eq 1)
		{
			$Nodes = $FindIt[0].Nodes
		}
		else
		{
			$Node = New-Object Windows.Forms.TreeNode($_)
			$Node.Name = $_
			$Node.Tag = $SelectPath
			[void]$Nodes.Add($Node)
			$FindIt = $Nodes.Find($_, $False)
			$Nodes = $FindIt[0].Nodes
		}
		if ($FindIt.tag -eq $DefaultSelectedOU)
		{
			$Global:NodeToShow = $FindIt[0]
		}
 	}
}

Function Show-QuickOu-Form($DefaultSelectedOU)
{
	$SelectOUForm = New-Object Windows.Forms.Form
	$SelectOUForm.Size = New-Object System.Drawing.Size(515,580) 
	$SelectOUForm.StartPosition = "CenterScreen" # Manual, WindowsDefaultLocation, WindowsDefaultBounds, CenterParent
	$SelectOUForm.FormBorderStyle = "FixedDialog" # FixedSingle, Fixed3D, FixedDialog, FixedToolWindow
	$SelectOuForm.MaximizeBox = $false
	$SelectOuForm.Text = "Please Select an Organizational Unit"
	$SelectOuForm.Icon = $Global:Icon

	$OUTreeView = New-Object Windows.Forms.TreeView
	$OUTreeView.PathSeparator = ","
	$OUTreeView.Size = New-Object System.Drawing.Size(500,500) 
	$OUTreeView.SelectAll
	$SelectOUForm.Controls.Add($OUTreeView)

	$IPProperties = [System.Net.NetworkInformation.IPGlobalProperties]::GetIPGlobalProperties()
	$strDNSDomain = $IPProperties.DomainName.toLower()
	$strDomainDN = $strDNSDomain.toString().split('.'); foreach ($strVal in $strDomainDN) {$strTemp += "dc=$strVal,"}; $strDomainDN = $strTemp.TrimEnd(",").toLower()
	$AllOUs= Get-ADObject -Filter 'ObjectClass -eq "organizationalUnit"' -SearchScope SubTree -SearchBase $strDomainDN
	#$AllOUs= Get-ADObject -Filter 'ObjectClass -eq "person"' -SearchScope SubTree -SearchBase $strDomainDN
	ForEach ($OU in $AllOUs)
	{
		$MyOU=$OU.DistinguishedName
		Add-QuickOU-Node $OUTreeView.Nodes $MyOU $DefaultSelectedOU
	}
	$OUTreeView.SelectedNode = $Global:NodeToShow

	$SelectOUButton = New-Object System.Windows.Forms.Button
	$SelectOUButton.Location = New-Object System.Drawing.Size(10,500)
	$SelectOUButton.Size = New-Object System.Drawing.Size(150,35)
	$SelectOUButton.Text = "Select"
	$SelectOUButton.Add_Click({
		$Global:AddCAPselectedOU = $OUTreeView.SelectedNode.tag
		$SelectOUForm.Close()
	})
	
	$SelectOUButton.Anchor = 'Bottom, Left'
	$SelectOUForm.Controls.Add($SelectOUButton)
	
	$MakeDefaultButton = New-Object System.Windows.Forms.Button
	$MakeDefaultButton.Location = New-Object System.Drawing.Size(170,500)
	$MakeDefaultButton.Size = New-Object System.Drawing.Size(150,35)
	$MakeDefaultButton.Text = "Select && Make Default"
	$MakeDefaultButton.Add_Click({
		$Global:AddCAPselectedOU = $OUTreeView.SelectedNode.tag
		$Global:DefaultOU = $OUTreeView.SelectedNode.tag
		$SelectOUForm.Close()
	})
	
	$MakeDefaultButton.Anchor = 'Bottom, Left'
	$SelectOUForm.Controls.Add($MakeDefaultButton)
	
	$CancelOUButton = New-Object System.Windows.Forms.Button
	$CancelOUButton.Location = New-Object System.Drawing.Size(335,500)
	$CancelOUButton.Size = New-Object System.Drawing.Size(150,35)
	$CancelOUButton.Text = "Cancel"
	$CancelOUButton.Add_Click({
	$SelectOUForm.Close()})
	$CancelOUButton.Anchor = 'Bottom, Left'
	$SelectOUForm.Controls.Add($CancelOUButton)
	$SelectOUForm.ShowDialog()
	$SelectOUForm.Dispose()
}

Function ReadSettings ()
{
	if (Test-Path -Path "$($Configfile)")
	{
		try
		{
			$xml = [xml](get-Content -path "$($Configfile)")
			$myDefaultOU = $xml.configuration.DefaultOU
			$myDefaultPool = $xml.configuration.DefaultPool
			$myDefaultSIPDomain = $xml.configuration.DefaultSIPDomain
		}
		catch
		{
			$myDefaultOU = ""
			$myDefaultPool = ""
			$myDefaultSIPDomain = ""
		}
	}
	else
	{
		# No file? Apply some defaults:
		$myDefaultOU = ""
		$myDefaultPool = ""
		$myDefaultSIPDomain = ""
	}
	return $myDefaultOU,$myDefaultPool,$myDefaultSIPDomain
}

Function WriteSettings ()
{
	param ([string]$myConfigfile, [string]$myDefaultOU, [string]$myDefaultPool, [string]$myDefaultSIPDomain)
	
	$SavedOU,$SavedPool,$SavedSIPDomain = ReadSettings
	
	if 	(($SavedOU.CompareTo($myDefaultOU) -eq 0) `
	-and ($SavedPool.CompareTo($myDefaultPool) -eq 0) `
	-and ($SavedSIPDomain.CompareTo($myDefaultSIPDomain) -eq 0))
	{
		#No need to touch the file, there are no changes
	}
	else
	{
		[xml]$Doc = New-Object System.Xml.XmlDocument
		$Dec = $Doc.CreateXmlDeclaration("1.0","UTF-8",$null)
		$Doc.AppendChild($Dec) | out-null
		$Root = $Doc.CreateNode("element","configuration",$null)
		$Element = $Doc.CreateElement("DefaultOU")
		$Element.InnerText = $myDefaultOU
		$Root.AppendChild($Element) | out-null
		$Element = $Doc.CreateElement("DefaultPool")
		$Element.InnerText = $myDefaultPool
		$Root.AppendChild($Element) | out-null
		$Element = $Doc.CreateElement("DefaultSIPDomain")
		$Element.InnerText = $myDefaultSIPDomain
		$Root.AppendChild($Element) | out-null
		$Doc.AppendChild($Root) | out-null
		try
		{
			$Doc.save(("$($myConfigfile)"))
		}
		catch
		{
		}
	}
}


function Filter-PolicyNames ([string]$Identity) 
{
	switch -wildcard ($Identity)
	{
		'Global' 	{ return }
		'Site*' 	{ return }
		'Service*' 	{ return }
		'Tag:*' 	{ return $Identity.SubString(4)}
	}
}	

function Filter-Objects ([object[]]$AllObjects)
{
	if ($Global:FilterSipText -ne "")
	{
		$AllObjects = $AllObjects | where-object {$_.SipAddress -like "$Global:FilterSipText"}
	}
	if ($Global:FilterLineUriText -ne "")
	{
		$AllObjects = $AllObjects | where-object {$_.LineUri -like "$Global:FilterLineUriText"}
	}
	if ($Global:FilterDisplayNameText -ne "")
	{
		$AllObjects = $AllObjects | where-object {$_.DisplayName -like "$Global:FilterDisplayNameText"}
	}
	if ($global:FilterRegistrarPool -ne "")
	{
		$allobjects = $allobjects | where-object {$_.RegistrarPool -match "$global:FilterRegistrarPool"}
	}
	return $AllObjects
}

function Test-ForPolicyChange ([object]$device, [string]$PolicyType, [string]$NewPolicyValue)
{
	switch ($PolicyType)
	{
		"DialPlan" 			{ $CurrentPolicyValue = $device.DialPlan }
		"VoicePolicy" 		{ $CurrentPolicyValue = $device.VoicePolicy }
		"ClientPolicy" 		{ $CurrentPolicyValue = $device.ClientPolicy }
		"LocationPolicy"	{ $CurrentPolicyValue = $device.LocationPolicy }
	}
	if ($NewPolicyValue -eq "") 	{ return } # No change - no policy selected (possible if the "existing" policy has been deleted)
	if ($NewPolicyValue -eq $null) 	{ return } # No change - no policy selected (possible if the "existing" policy has been deleted)
	if ($NewPolicyValue -eq "<No Change>") { return } # No change
	if ($CurrentPolicyValue -eq $null)
	{
		if ($NewPolicyValue -eq "<Automatic>")
		{ 
			return # No change
		} 
	}
	else
	{
		if ($CurrentPolicyValue.ToString() -eq $NewPolicyValue ) { return } # No change
	}
	if ($NewPolicyValue -eq "<Automatic>" ) { return "`$null" } #Note the quotes & back-tick here - this is the 5-character string "$null", not a null value!
	return "Tag:$NewPolicyValue"
}


function Show-FilterForm([string] $TempFilterOUText)
{

	write-progress -id 1 -Activity "Loading" -Status "Define the Filter form" 
	$FilterForm = New-Object System.Windows.Forms.Form
	$FilterForm.Text = "Filter"
	$FilterForm.Size = New-Object System.Drawing.Size(515,580) 
	$FilterForm.StartPosition = "CenterScreen" # Manual, WindowsDefaultLocation, WindowsDefaultBounds, CenterParent
	$FilterForm.Autosize = $False
	$FilterForm.FormBorderStyle = "FixedDialog" # FixedSingle, Fixed3D, FixedDialog, FixedToolWindow
	$FilterForm.MinimizeBox = $False
	$FilterForm.MaximizeBox = $False
	$FilterForm.WindowState = "Normal" # Maximized, Minimized, Normal
	$FilterForm.SizeGripStyle = "Hide" # Auto, Hide, Show
	$FilterForm.ShowInTaskbar = $False
	$FilterForm.Icon = $Global:Icon
	write-progress -id 1 -Activity "Loading" -Status "Define the Filter form" -Complete
	
	# ============================================================================
	# OU GroupBox
	# ============================================================================	
	$OUFilterGroupBox = New-Object System.Windows.Forms.GroupBox
	$OUFilterGroupBox.Location = New-Object System.Drawing.Size(15,15)
	$OUFilterGroupBox.Size = New-Object System.Drawing.Size(465,110)
	$OUFilterGroupBox.Font = New-Object System.Drawing.Font("Arial", "10")
	$OUFilterGroupBox.Text = "OU"
	$FilterForm.Controls.Add($OUFilterGroupBox)	
	
	# ============================================================================
	# OU GroupBox - Textbox
	# ============================================================================	
	$FilterOuTextBox = New-Object System.Windows.Forms.TextBox
	$FilterOuTextBox.Location = New-Object System.Drawing.Size(15,25)
	$FilterOuTextBox.Size = New-Object System.Drawing.Size(430,20)
	$FilterOuTextBox.Multiline = $False
	$FilterOuTextBox.Font = New-Object System.Drawing.Font("Arial", "10")
	$FilterOuTextBox.ReadOnly = $True
	$FilterOuTextBox.Text = $Global:FilterOu
	$FilterOuTextBox.TabStop = $False #Otherwise being the first item on the form it's selected by default
	$OUFilterGroupBox.Controls.Add($FilterOuTextBox)

	# ============================================================================
	# OU GroupBox - Clear OU Filters button
	# ============================================================================	
	$FilterClearOuFilterButton = New-Object System.Windows.Forms.Button
	$FilterClearOuFilterButton.Location = New-Object System.Drawing.Size(15,60)
	$FilterClearOuFilterButton.Size = New-Object System.Drawing.Size(150,35)
	$FilterClearOuFilterButton.Text = "Clear OU Filter"
	$FilterClearOuFilterButton.Add_Click({
		$FilterOuTextBox.Text = ""
	})
	$FilterClearOuFilterButton.Anchor = 'Bottom, Left'
	$OUFilterGroupBox.Controls.Add($FilterClearOuFilterButton)
	
	# ============================================================================
	# OU GroupBox - Browse button
	# ============================================================================	
	$FilterOuBrowseButton = New-Object System.Windows.Forms.Button
	$FilterOuBrowseButton.Name = "Browse"
	$FilterOuBrowseButton.Text = "Browse"
	$FilterOuBrowseButton.TabIndex = 7
	$FilterOuBrowseButton.Location = New-Object System.Drawing.Size(365,60)
	$FilterOuBrowseButton.Size = New-Object System.Drawing.Size(80,35)
	$FilterOuBrowseButton.Add_Click({
		Show-QuickOu-Form ($FilterOuTextBox.Text)
		$FilterOuTextBox.Text=$Global:AddCAPselectedOU
		})
	$OUFilterGroupBox.Controls.Add($FilterOuBrowseButton)
	
	# ============================================================================
	# Filter Filters GroupBox
	# ============================================================================	
	$FilterFiltersGroupBox = New-Object System.Windows.Forms.GroupBox
	$FilterFiltersGroupBox.Location = New-Object System.Drawing.Size(15,155)
	$FilterFiltersGroupBox.Size = New-Object System.Drawing.Size(465,280)
	$FilterFiltersGroupBox.Font = New-Object System.Drawing.Font("Arial", "10")
	$FilterFiltersGroupBox.Text = "Filters"
	$FilterForm.Controls.Add($FilterFiltersGroupBox)	
	
	# ============================================================================
	# Filter Filters GroupBox - Description text
	# ============================================================================	
	$FilterSipUriTitleBox = New-Object System.Windows.Forms.Label
	$FilterSipUriTitleBox.Location = New-Object System.Drawing.Size(115,32)
	$FilterSipUriTitleBox.Size = New-Object System.Drawing.Size(285,20)
	$FilterSipUriTitleBox.Font = New-Object System.Drawing.Font("Arial", "10")
	$FilterSipUriTitleBox.Text = 'Use PowerShell "-like" syntax, e.g. *+441*'
	$FilterFiltersGroupBox.Controls.Add($FilterSipUriTitleBox)
	
	# ============================================================================
	# Filter Filters GroupBox - SIP Address filter text & field
	# ============================================================================	
	$FilterSipUriTitleBox = New-Object System.Windows.Forms.Label
	$FilterSipUriTitleBox.Location = New-Object System.Drawing.Size(15,62)
	$FilterSipUriTitleBox.Size = New-Object System.Drawing.Size(85,20)
	$FilterSipUriTitleBox.Font = New-Object System.Drawing.Font("Arial", "10")
	$FilterSipUriTitleBox.Text = "SIP Address"
	$FilterFiltersGroupBox.Controls.Add($FilterSipUriTitleBox)

	$FilterSipUriTextBox = New-Object System.Windows.Forms.TextBox
	$FilterSipUriTextBox.Location = New-Object System.Drawing.Size(115,60)
	$FilterSipUriTextBox.Size = New-Object System.Drawing.Size(320,20)
	$FilterSipUriTextBox.Multiline = $False
	$FilterSipUriTextBox.Font = New-Object System.Drawing.Font("Arial", "10")
	$FilterSipUriTextBox.ReadOnly = $False
	$FilterSipUriTextBox.Text = ""
	$FilterFiltersGroupBox.Controls.Add($FilterSipUriTextBox)
	
	# ============================================================================
	# Filter Filters GroupBox - Line URI text & field
	# ============================================================================	
	$FilterLineUriTitleBox = New-Object System.Windows.Forms.Label
	$FilterLineUriTitleBox.Location = New-Object System.Drawing.Size(15,92)
	$FilterLineUriTitleBox.Size = New-Object System.Drawing.Size(60,20)
	$FilterLineUriTitleBox.Font = New-Object System.Drawing.Font("Arial", "10")
	$FilterLineUriTitleBox.Text = "Line URI"
	$FilterFiltersGroupBox.Controls.Add($FilterLineUriTitleBox)

	$FilterLineUriTextBox = New-Object System.Windows.Forms.TextBox
	$FilterLineUriTextBox.Location = New-Object System.Drawing.Size(115,90)
	$FilterLineUriTextBox.Size = New-Object System.Drawing.Size(320,20)
	$FilterLineUriTextBox.Multiline = $False
	$FilterLineUriTextBox.Font = New-Object System.Drawing.Font("Arial", "10")
	$FilterLineUriTextBox.ReadOnly = $False
	$FilterLineUriTextBox.Text = ""
	$FilterFiltersGroupBox.Controls.Add($FilterLineUriTextBox)
	
	# ============================================================================
	# Filter Filters GroupBox - Display Name text & field
	# ============================================================================	
	$FilterDisplayNameTitleBox = New-Object System.Windows.Forms.Label
	$FilterDisplayNameTitleBox.Location = New-Object System.Drawing.Size(15,122)
	$FilterDisplayNameTitleBox.Size = New-Object System.Drawing.Size(100,20)
	$FilterDisplayNameTitleBox.Font = New-Object System.Drawing.Font("Arial", "10")
	$FilterDisplayNameTitleBox.Text = "Display Name"
	$FilterFiltersGroupBox.Controls.Add($FilterDisplayNameTitleBox)

	$FilterDisplayNameTextBox = New-Object System.Windows.Forms.TextBox
	$FilterDisplayNameTextBox.Location = New-Object System.Drawing.Size(115,120)
	$FilterDisplayNameTextBox.Size = New-Object System.Drawing.Size(320,20)
	$FilterDisplayNameTextBox.Multiline = $False
	$FilterDisplayNameTextBox.Font = New-Object System.Drawing.Font("Arial", "10")
	$FilterDisplayNameTextBox.ReadOnly = $False
	$FilterDisplayNameTextBox.Text = ""
	$FilterFiltersGroupBox.Controls.Add($FilterDisplayNameTextBox)
	
	# ============================================================================
	# Filter Filters GroupBox - Registrar Pool text & *combobox*
	# ============================================================================	
	$FilterRegistrarPoolTitleBox = New-Object System.Windows.Forms.Label
	$FilterRegistrarPoolTitleBox.Location = New-Object System.Drawing.Size(15,172)
	$FilterRegistrarPoolTitleBox.Size = New-Object System.Drawing.Size(100,20)
	$FilterRegistrarPoolTitleBox.Font = New-Object System.Drawing.Font("Arial", "10")
	$FilterRegistrarPoolTitleBox.Text = "Registrar Pool"
	$FilterFiltersGroupBox.Controls.Add($FilterRegistrarPoolTitleBox)

	$FilterRegistrarPoolCombobox = New-Object System.Windows.Forms.Combobox
	$FilterRegistrarPoolCombobox.Location = New-Object System.Drawing.Size(115,170)
	$FilterRegistrarPoolCombobox.Size = New-Object System.Drawing.Size(320,20)
	$FilterRegistrarPoolCombobox.Font = New-Object System.Drawing.Font("Arial", "9")
	$FilterRegistrarPoolCombobox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
	[void] $FilterRegistrarPoolCombobox.Items.Add("<Don't care>")
	foreach ($item in $Global:RegistrarList)
	{   
		[void] $FilterRegistrarPoolCombobox.Items.Add($item.poolFQDN)
	}
	$FilterFiltersGroupBox.Controls.Add($FilterRegistrarPoolCombobox)

	# ============================================================================
	# Filter Filters GroupBox - Clear Filters button
	# ============================================================================	
	$FilterClearFiltersButton = New-Object System.Windows.Forms.Button
	$FilterClearFiltersButton.Location = New-Object System.Drawing.Size(15,230)
	$FilterClearFiltersButton.Size = New-Object System.Drawing.Size(150,35)
	$FilterClearFiltersButton.Text = "Clear Filters"
	$FilterClearFiltersButton.Add_Click({
		$FilterSipUriTextBox.Text = ""
		$FilterLineUriTextBox.Text = ""
		$FilterDisplayNameTextBox.Text = ""
		$FilterRegistrarPoolCombobox.SelectedIndex = $FilterRegistrarPoolCombobox.findstring("<Don't care>")
	})
	$FilterClearFiltersButton.Anchor = 'Bottom, Left'
	$FilterFiltersGroupBox.Controls.Add($FilterClearFiltersButton)
	
	
	# ============================================================================
	# Apply button
	# ============================================================================	
	$FilterApplyButton = New-Object System.Windows.Forms.Button
	$FilterApplyButton.Location = New-Object System.Drawing.Size(15,500)
	$FilterApplyButton.Size = New-Object System.Drawing.Size(150,35)
	$FilterApplyButton.Text = "Apply"
	$FilterApplyButton.Add_Click({
		$Global:FilterSipText = $FilterSipUriTextBox.Text
		$Global:FilterLineUriText = $FilterLineUriTextBox.Text
		$Global:FilterDisplayNameText = $FilterDisplayNameTextBox.Text
		if ($FilterRegistrarPoolCombobox.SelectedItem -eq "<Don't care>")
		{
			$Global:FilterRegistrarPool = ""
		}
		else
		{
			$Global:FilterRegistrarPool = $FilterRegistrarPoolCombobox.SelectedItem
		}
		$FilterForm.Close()
		if ($FilterOuTextBox.Text -ne $TempFilterOUText)
		{
			#A new OU was selected - update the global variable
			$Global:FilterOu = $FilterOuTextBox.Text
			#Update-DeviceList	#Reads all objects from Lync/SfB
		}
	})
	$FilterApplyButton.Anchor = 'Bottom, Left'
	$FilterForm.Controls.Add($FilterApplyButton)
	
	# ============================================================================
	# Cancel button
	# ============================================================================		
	$CancelFilterButton = New-Object System.Windows.Forms.Button
	$CancelFilterButton.Location = New-Object System.Drawing.Size(330,500)
	$CancelFilterButton.Size = New-Object System.Drawing.Size(150,35)
	$CancelFilterButton.Text = "Cancel"
	$CancelFilterButton.Add_Click({
		#Abandon any changes made here
		$Global:FilterSipText = ""
		$Global:FilterLineUriText = ""
		$Global:FilterRegistrarPool = ""
		$FilterForm.Close()
		})
	$CancelFilterButton.Anchor = 'Bottom, Left'
	$FilterForm.Controls.Add($CancelFilterButton)
	
	#Initialise fields with last chosen filter values:
	$FilterSipUriTextBox.Text = $Global:FilterSipText
	$FilterLineUriTextBox.Text = $Global:FilterLineUriText
	$FilterDisplayNameTextBox.Text = $Global:FilterDisplayNameText
	$FilterRegistrarPoolCombobox.SelectedIndex = $FilterRegistrarPoolCombobox.findstring("$Global:FilterRegistrarPool")
	$FilterForm.ShowDialog()
	$FilterForm.Dispose()
}


function Get-UpdateInfo
{
  <#
      .SYNOPSIS
      Queries an online XML source for version information to determine if a new version of the script is available.
	  *** This version customised by Greig Sheridan. @greiginsydney https://greiginsydney.com ***

      .DESCRIPTION
      Queries an online XML source for version information to determine if a new version of the script is available.

      .NOTES
      Version               : 1.2 - See changelog at https://ucunleashed.com/3168 for fixes & changes introduced with each version
      Wish list             : Better error trapping
      Rights Required       : N/A
      Sched Task Required   : No
      Lync/Skype4B Version  : N/A
      Author/Copyright      : © Pat Richard, Office Servers and Services (Skype for Business) MVP - All Rights Reserved
      Email/Blog/Twitter    : pat@innervation.com  https://ucunleashed.com  @patrichard
      Donations             : https://www.paypal.me/PatRichard
      Dedicated Post        : https://ucunleashed.com/3168
      Disclaimer            : You running this script/function means you will not blame the author(s) if this breaks your stuff. This script/function 
                            is provided AS IS without warranty of any kind. Author(s) disclaim all implied warranties including, without limitation, 
                            any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use 
                            or performance of the sample scripts and documentation remains with you. In no event shall author(s) be held liable for 
                            any damages whatsoever (including, without limitation, damages for loss of business profits, business interruption, loss 
                            of business information, or other pecuniary loss) arising out of the use of or inability to use the script or 
                            documentation. Neither this script/function, nor any part of it other than those parts that are explicitly copied from 
                            others, may be republished without author(s) express written permission. Author(s) retain the right to alter this 
                            disclaimer at any time. For the most up to date version of the disclaimer, see https://ucunleashed.com/code-disclaimer.
      Acknowledgements      : Reading XML files 
                            http://stackoverflow.com/questions/18509358/how-to-read-xml-in-powershell
                            http://stackoverflow.com/questions/20433932/determine-xml-node-exists
      Assumptions           : ExecutionPolicy of AllSigned (recommended), RemoteSigned, or Unrestricted (not recommended)
      Limitations           : 
      Known issues          : 

      .EXAMPLE
      Get-UpdateInfo -Title "Compare-PkiCertificates.ps1"

      Description
      -----------
      Runs function to check for updates to script called <Varies>.

      .INPUTS
      None. You cannot pipe objects to this script.
  #>
	[CmdletBinding(SupportsShouldProcess = $true)]
	param (
	[string] $title
	)
	try
	{
		[bool] $HasInternetAccess = ([Activator]::CreateInstance([Type]::GetTypeFromCLSID([Guid]'{DCB00C01-570F-4A9B-8D69-199FDBA5723B}')).IsConnectedToInternet)
		if ($HasInternetAccess)
		{
			write-verbose "Performing update check"
			# ------------------ TLS 1.2 fixup from https://github.com/chocolatey/choco/wiki/Installation#installing-with-restricted-tls
			$securityProtocolSettingsOriginal = [System.Net.ServicePointManager]::SecurityProtocol
			try {
			  # Set TLS 1.2 (3072). Use integers because the enumeration values for TLS 1.2 won't exist in .NET 4.0, even though they are 
			  # addressable if .NET 4.5+ is installed (.NET 4.5 is an in-place upgrade).
			  [System.Net.ServicePointManager]::SecurityProtocol = 3072
			} catch {
			  Write-verbose 'Unable to set PowerShell to use TLS 1.2 due to old .NET Framework installed.'
			}
			# ------------------ end TLS 1.2 fixup
			[xml] $xml = (New-Object -TypeName System.Net.WebClient).DownloadString('https://greiginsydney.com/wp-content/version.xml')
			[System.Net.ServicePointManager]::SecurityProtocol = $securityProtocolSettingsOriginal #Reinstate original SecurityProtocol settings
			$article  = select-XML -xml $xml -xpath "//article[@title='$($title)']"
			[string] $Ga = $article.node.version.trim()
			if ($article.node.changeLog)
			{
				[string] $changelog = "This version includes: " + $article.node.changeLog.trim() + "`n`n"
			}
			if ($Ga -gt $ScriptVersion)
			{
				$wshell = New-Object -ComObject Wscript.Shell -ErrorAction Stop
				$updatePrompt = $wshell.Popup("Version $($ga) is available.`n`n$($changelog)Would you like to download it?",0,"New version available",68)
				if ($updatePrompt -eq 6)
				{
					Start-Process -FilePath $article.node.downloadUrl
					Write-Warning "Script is exiting. Please run the new version of the script after you've downloaded it."
					exit
				}
				else
				{
					write-verbose "Upgrade to version $($ga) was declined"
				}
			}
			elseif ($Ga -eq $ScriptVersion)
			{
				write-verbose "Script version $($Scriptversion) is the latest released version"
			}
			else
			{
				write-verbose "Script version $($Scriptversion) is newer than the latest released version $($ga)"
			}
		}
		else
		{
		}
	
	} # end function Get-UpdateInfo
	catch
	{
		write-verbose "Caught error in Get-UpdateInfo"
		if ($Global:Debug)
		{				
			$Global:error | fl * -f #This dumps to screen as white for the time being. I haven't been able to get it to dump in red
		}
	}
}


# ============================================================================
# END FUNCTIONS ==============================================================
# ============================================================================ 

$Global:DefaultOU,$Global:DefaultPool,$Global:DefaultSIPDomain = ReadSettings	#This has to happen here otherwise the form won't populate with the defaults.
$global:AddCAPselectedOU = $Global:DefaultOU
$Global:FilterOu = $Global:DefaultOU

if ($skipupdatecheck)
{
	write-verbose "Skipping update check"
}
else
{
	write-progress -id 1 -Activity "Performing update check" -Status "Running Get-UpdateInfo" -PercentComplete (50)
	Get-UpdateInfo -title "MAdCaP.ps1"
	write-progress -id 1 -Activity "Back from performing update check" -Status "Running Get-UpdateInfo" -Completed
}

# ============================================================================
# Define the form ============================================================
# ============================================================================ 
write-progress -id 1 -Activity "Loading" -Status "Define the form" 
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "MAdCaP.ps1"
$Form.Size = New-Object System.Drawing.Size(780,680)
$Form.StartPosition = "CenterScreen" # Manual, WindowsDefaultLocation, WindowsDefaultBounds, CenterParent
$Form.Autosize = $False
$Form.FormBorderStyle = "FixedDialog" # FixedSingle, Fixed3D, FixedDialog, FixedToolWindow
$Form.MinimizeBox = $True
$Form.MaximizeBox = $False
$Form.WindowState = "Normal" # Maximized, Minimized, Normal
$Form.SizeGripStyle = "Hide" # Auto, Hide, Show
$Form.ShowInTaskbar = $True
$Form.Icon = $Global:Icon

# ============================================================================
# Define the TAB structure ===================================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Define the tab structure"
$TabControl = New-Object System.Windows.Forms.TabControl
$TabControl.Location = New-Object System.Drawing.Point(15,15)
$TabControl.Name = "tabControl"
$TabControl.Size = New-Object System.Drawing.Size(735,540)
$Form.Controls.Add($TabControl)

$TabPageNew = New-Object System.Windows.Forms.TabPage
$TabPageNew.Name = "tabPageNew"
$TabPageNew.Text = "New Object"
$TabControl.Controls.Add($TabPageNew)

$TabPageSet = New-Object System.Windows.Forms.TabPage
$TabPageSet.Name = "tabPageSet"
$TabPageSet.Text = "Existing Object"
$TabControl.Controls.Add($TabPageSet)

$TabControl.Add_SelectedIndexChanged(
{
	if ($TabControl.SelectedIndex -eq 1) 
	{
		Update-DeviceList	#Reads all objects from Lync/SfB
		Update-ADCAPLists	#Populates the two lists on the Existing Items tab
		Update-Display		#Shows the contents of one item to the pane & the RH Policy lists
	}
	handler_ValidateGo
})

# ============================================================================
# Add some descriptive / instructional text ==================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add some text"
$ExplanationTextBox = New-Object System.Windows.Forms.Label
$ExplanationTextBox.Location = New-Object System.Drawing.Size(300,25)
$ExplanationTextBox.Size = New-Object System.Drawing.Size(440,40)
$ExplanationTextBox.Font = New-Object System.Drawing.Font("Arial", "10",[System.Drawing.FontStyle]::Bold)
#$ExplanationTextBox.ForeColor = [System.Drawing.Color]::FromArgb(255,176,196,222)
$ExplanationTextBox.ForeColor = [System.Drawing.Color]::"DarkBlue"
$ExplanationTextBox.Text = "Create new objects on this tab.`nSelect the ""Existing Object"" tab to set Policies and a PIN"
$TabPageNew.Controls.Add($ExplanationTextBox)


# ============================================================================
# Add the Analog Device / Common Area Phone Radio buttons ====================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add radio buttons"
$RadiobuttonAD = New-Object System.Windows.Forms.RadioButton
$RadiobuttonAD.Location = New-Object System.Drawing.Point(25,20)
$RadiobuttonAD.Name = "AnalogDevice"
$RadiobuttonAD.Size = New-Object System.Drawing.Size(180,20)
$RadiobuttonAD.Font = New-Object System.Drawing.Font("Arial", "10",[System.Drawing.FontStyle]::Bold)
$RadiobuttonAD.TabStop = $True
$RadiobuttonAD.Checked = $True
$RadiobuttonAD.Text = "Analog Device"
$RadiobuttonAD.Add_CheckedChanged(
{
	if ($RadiobuttonAD.Checked -eq $True)
	{
	$NewAnalogFaxTitleBox.Enabled = $True
	$NewAnalogFaxComboBox.Enabled = $True
	$NewGatewayTitleBox.Enabled = $True
	$NewPstnGatewayCombobox.Enabled = $True
	$NewDescriptionTextBox.Enabled = $False
	}
	else
	{
	$NewAnalogFaxTitleBox.Enabled = $False
	$NewAnalogFaxComboBox.Enabled = $False
	$NewGatewayTitleBox.Enabled = $False
	$NewPstnGatewayCombobox.Enabled = $False
	$NewDescriptionTextBox.Enabled = $True
	}
})
$TabPageNew.Controls.Add($RadiobuttonAD)

$RadiobuttonCA = New-Object System.Windows.Forms.RadioButton
$RadiobuttonCA.Location = New-Object System.Drawing.Point(25,45)
$RadiobuttonCA.Name = "CommonAreaPhone"
$RadiobuttonCA.Size = New-Object System.Drawing.Size(180,20)
$RadiobuttonCA.Font = New-Object System.Drawing.Font("Arial", "10",[System.Drawing.FontStyle]::Bold)
$RadiobuttonCA.TabStop = $True
$RadiobuttonCA.Text = "Common Area Phone"
$TabPageNew.Controls.Add($RadiobuttonCA)

$RadiobuttonCA.add_CheckedChanged({handler_ValidateGo})


# ============================================================================
# Add the "Required" group box ===============================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the ""Required"" group box"
$NewRequiredGroupBox = New-Object System.Windows.Forms.GroupBox
$NewRequiredGroupBox.Location = New-Object System.Drawing.Size(15,75)
$NewRequiredGroupBox.Size = New-Object System.Drawing.Size(695,250)
$NewRequiredGroupBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewRequiredGroupBox.Text = "Required Parameters"
$TabPageNew.Controls.Add($NewRequiredGroupBox)


# ============================================================================
# Line URI ===================================================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add line URI"
$NewLineUriTitleBox = New-Object System.Windows.Forms.Label
$NewLineUriTitleBox.Location = New-Object System.Drawing.Size(15,32)
$NewLineUriTitleBox.Size = New-Object System.Drawing.Size(60,20)
$NewLineUriTitleBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewLineUriTitleBox.Text = "Line URI"
$NewRequiredGroupBox.Controls.Add($NewLineUriTitleBox)

$NewLineUriTelBox = New-Object System.Windows.Forms.Label
$NewLineUriTelBox.Location = New-Object System.Drawing.Size(107,32)
$NewLineUriTelBox.Size = New-Object System.Drawing.Size(28,20)
$NewLineUriTelBox.Font = New-Object System.Drawing.Font("Arial", "10",[System.Drawing.FontStyle]::Bold)
$NewLineUriTelBox.Text = "tel:"
$NewRequiredGroupBox.Controls.Add($NewLineUriTelBox)

$NewLineUriTextBox = New-Object System.Windows.Forms.TextBox
$NewLineUriTextBox.Location = New-Object System.Drawing.Size(135,30)
$NewLineUriTextBox.Size = New-Object System.Drawing.Size(270,20)
$NewLineUriTextBox.Multiline = $False
$NewLineUriTextBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewLineUriTextBox.ReadOnly = $False
$NewLineUriTextBox.Text = ""
$NewRequiredGroupBox.Controls.Add($NewLineUriTextBox)


# ============================================================================
# Validate Line URI text values ==============================================
# ============================================================================
$NewLineUriTextBox.Add_TextChanged(
{ 
	# Only accept digits 0-9 and "+" in this field
	$NewLineUriTextBox.Text = [regex]::replace($NewLineUriTextBox.Text, "([^0-9+;ext=])" , "")
	$NewLineUriTextBox.SelectionStart = $NewLineUriTextBox.Text.Length
})
$NewLineUriTextBox.Add_TextChanged({handler_ValidateGo})

# ============================================================================
# Add the ComboBox containing the Registrar Pool =============================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add registrar pool combo box"
$NewRegistrarPoolTitleBox = New-Object System.Windows.Forms.Label
$NewRegistrarPoolTitleBox.Location = New-Object System.Drawing.Size(15,62)
$NewRegistrarPoolTitleBox.Size = New-Object System.Drawing.Size(105,20)
$NewRegistrarPoolTitleBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewRegistrarPoolTitleBox.Text = "Registrar Pool"
$NewRequiredGroupBox.Controls.Add($NewRegistrarPoolTitleBox)

$NewRegistrarPoolCombobox = New-Object System.Windows.Forms.Combobox
$NewRegistrarPoolCombobox.Location = New-Object System.Drawing.Size(135,60)
$NewRegistrarPoolCombobox.Size = New-Object System.Drawing.Size(270,20)
$NewRegistrarPoolCombobox.Font = New-Object System.Drawing.Font("Arial", "9")
$NewRegistrarPoolCombobox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
foreach ($item in $Global:RegistrarList)
{   
	[void] $NewRegistrarPoolCombobox.Items.Add($item.poolFQDN)
}
# If there's a default, select it:
if ($Global:DefaultPool -ne "")
{	
	$NewRegistrarPoolCombobox.SelectedIndex = $NewRegistrarPoolCombobox.findstring("$Global:DefaultPool")
}
# If the above failed to match or there's no Default, select the first one in the list:
if ($NewRegistrarPoolCombobox.SelectedIndex -eq -1)
{
	$NewRegistrarPoolCombobox.SelectedIndex = 0      # Automatically select the first one
}
$NewRegistrarPoolCombobox.Add_SelectedIndexChanged({ $Global:DefaultPool = $NewRegistrarPoolCombobox.SelectedItem })
$NewRequiredGroupBox.Controls.Add($NewRegistrarPoolCombobox)


# ============================================================================
# Add the text for the OU ====================================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add OU text"
$NewOuTitleBox = New-Object System.Windows.Forms.Label
$NewOuTitleBox.Location = New-Object System.Drawing.Size(15,92)
$NewOuTitleBox.Size = New-Object System.Drawing.Size(60,20)
$NewOuTitleBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewOuTitleBox.Text = "OU"
$NewRequiredGroupBox.Controls.Add($NewOuTitleBox)

$NewOuTextBox = New-Object System.Windows.Forms.TextBox
$NewOuTextBox.Location = New-Object System.Drawing.Size(135,90)
$NewOuTextBox.Size = New-Object System.Drawing.Size(436,20)
$NewOuTextBox.Multiline = $False
$NewOuTextBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewOuTextBox.ReadOnly = $False
$NewOuTextBox.Text = $Global:DefaultOU
$NewRequiredGroupBox.Controls.Add($NewOuTextBox)

$ToolTip = New-Object System.Windows.Forms.ToolTip
$ToolTip.BackColor = [System.Drawing.Color]::LightGoldenrodYellow
$ToolTip.IsBalloon = $true
$ToolTip.InitialDelay = 500
$ToolTip.ReshowDelay = 500
$ToolTip.SetToolTip($NewOuTextBox, "Enter in the format OU=blah,DC=contoso,DC=local") 
$ToolTip.SetToolTip($NewOuTitleBox, "Enter in the format OU=blah,DC=contoso,DC=local") 

$NewOuBrowseButton = New-Object System.Windows.Forms.Button
$NewOuBrowseButton.Name = "Browse"
$NewOuBrowseButton.Text = "Browse"
$NewOuBrowseButton.TabIndex = 7
$NewOuBrowseButton.Size = New-Object System.Drawing.Size(60,20)
$NewOuBrowseButton.Location = New-Object System.Drawing.Size(586,90)
$NewOuBrowseButton.Add_Click({
	if ($Global:AddCAPselectedOU -eq "")
	{
		$Global:AddCAPselectedOU = $Global:DefaultOU
	}
	Show-QuickOu-Form ($NewOuTextBox.Text)
	$NewOuTextBox.Text=$Global:AddCAPselectedOU
	})
$NewRequiredGroupBox.Controls.Add($NewOuBrowseButton)


# ============================================================================
# Validate the OU ============================================================
# ============================================================================
$NewOuTextBox.Add_TextChanged(
{ 
	# # Prevent them adding quotes around it
	$NewOuTextBox.Text = [regex]::replace($NewOuTextBox.Text, '"' , "")
	#$NewOuTextBox.SelectionStart = $NewOuTextBox.Text.Length
	$Global:AddCAPselectedOU = $NewOuTextBox.Text
	handler_ValidateGo
})

# ============================================================================
# Add the text for the DN ====================================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add DN text"
$NewDnTitleBox = New-Object System.Windows.Forms.Label
$NewDnTitleBox.Location = New-Object System.Drawing.Size(15,122)
$NewDnTitleBox.Size = New-Object System.Drawing.Size(60,20)
$NewDnTitleBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewDnTitleBox.Text = "DN"
$NewRequiredGroupBox.Controls.Add($NewDnTitleBox)

$NewDnTextBox = New-Object System.Windows.Forms.TextBox
$NewDnTextBox.Location = New-Object System.Drawing.Size(135,120)
$NewDnTextBox.Size = New-Object System.Drawing.Size(436,20)
$NewDnTextBox.Multiline = $False
$NewDnTextBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewDnTextBox.ReadOnly = $False
$NewDnTextBox.Text = ""
$NewRequiredGroupBox.Controls.Add($NewDnTextBox)

$ToolTip = New-Object System.Windows.Forms.ToolTip
$ToolTip.BackColor = [System.Drawing.Color]::LightGoldenrodYellow
$ToolTip.IsBalloon = $true
$ToolTip.InitialDelay = 500
$ToolTip.ReshowDelay = 500
$ToolTip.SetToolTip($NewDnTextBox, "Enter in the format CN=ExistingMeetingRoom,OU=blah,DC=contoso,DC=local") 
$ToolTip.SetToolTip($NewDnTitleBox, "Enter in the format CN=ExistingMeetingRoom,OU=blah,DC=contoso,DC=local") 

$NewDnWarningBox = New-Object System.Windows.Forms.Label
$NewDnWarningBox.Location = New-Object System.Drawing.Size(135,150)
$NewDnWarningBox.Size = New-Object System.Drawing.Size(340,20)
$NewDnWarningBox.Font = New-Object System.Drawing.Font("Arial", "9")
$NewDnWarningBox.Text = "(OU && DN are mutually exclusive. A DN overrides an OU)"
$NewRequiredGroupBox.Controls.Add($NewDnWarningBox)


# ============================================================================
# Validate the DN ============================================================
# ============================================================================
$NewDnTextBox.Add_TextChanged(
{ 
	# Prevent them adding quotes around it
	$NewDnTextBox.Text = [regex]::replace($NewDnTextBox.Text, '"' , "")
	#$NewDnTextBox.SelectionStart = $NewDnTextBox.Text.Length
	if ($NewDnTextBox.Text -ne "")
	{
		$NewOuTextBox.Enabled = $False
	}
	else
	{
		$NewOuTextBox.Enabled = $True
	}
	handler_ValidateGo
})

# ============================================================================
# Add the Analog Fax ComboBox ================================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the analog fax combo box"
$NewAnalogFaxTitleBox = New-Object System.Windows.Forms.Label
$NewAnalogFaxTitleBox.Location = New-Object System.Drawing.Size(15,182)
$NewAnalogFaxTitleBox.Size = New-Object System.Drawing.Size(80,20)
$NewAnalogFaxTitleBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewAnalogFaxTitleBox.Text = "Analog Fax"
$NewRequiredGroupBox.Controls.Add($NewAnalogFaxTitleBox)

$NewAnalogFaxCombobox = New-Object System.Windows.Forms.Combobox
$NewAnalogFaxCombobox.Location = New-Object System.Drawing.Size(135,180)
$NewAnalogFaxCombobox.Size = New-Object System.Drawing.Size(270,20)
$NewAnalogFaxCombobox.Font = New-Object System.Drawing.Font("Arial", "9")
$NewAnalogFaxCombobox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void] $NewAnalogFaxCombobox.Items.Add("$True")
[void] $NewAnalogFaxCombobox.Items.Add("$False")
$NewAnalogFaxCombobox.SelectedIndex = 0            # Automatically select the first one
$NewRequiredGroupBox.Controls.Add($NewAnalogFaxCombobox)


# ============================================================================
# Add the text for the Gateway ===============================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the Gateway combo box"
$NewGatewayTitleBox = New-Object System.Windows.Forms.Label
$NewGatewayTitleBox.Location = New-Object System.Drawing.Size(15,212)
$NewGatewayTitleBox.Size = New-Object System.Drawing.Size(80,20)
$NewGatewayTitleBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewGatewayTitleBox.Text = "Gateway"
$NewRequiredGroupBox.Controls.Add($NewGatewayTitleBox)

$NewPstnGatewayCombobox = New-Object System.Windows.Forms.Combobox
$NewPstnGatewayCombobox.Location = New-Object System.Drawing.Size(135,210)
$NewPstnGatewayCombobox.Size = New-Object System.Drawing.Size(270,20)
$NewPstnGatewayCombobox.Font = New-Object System.Drawing.Font("Arial", "9")
$NewPstnGatewayCombobox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$items = Invoke-Expression "Get-CsService -PstnGateway -verbose:`$false"
foreach ($item in $items)
{   
	[void] $NewPstnGatewayCombobox.Items.Add($item.poolFQDN)
}
$NewPstnGatewayCombobox.SelectedIndex = 0      # Automatically select the first one
$NewRequiredGroupBox.Controls.Add($NewPstnGatewayCombobox)


# ============================================================================
# Add the "Optional" group box ===============================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the Optional parameters group box"
$NewOptionalGroupBox = New-Object System.Windows.Forms.GroupBox
$NewOptionalGroupBox.Location = New-Object System.Drawing.Size(15,335)
$NewOptionalGroupBox.Size = New-Object System.Drawing.Size(695,170)
$NewOptionalGroupBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewOptionalGroupBox.Text = "Optional Parameters"
$TabPageNew.Controls.Add($NewOptionalGroupBox)


# ============================================================================
# Add the Display Name =======================================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the Display Name"
$NewDisplayNameTitleBox = New-Object System.Windows.Forms.Label
$NewDisplayNameTitleBox.Location = New-Object System.Drawing.Size(15,27)
$NewDisplayNameTitleBox.Size = New-Object System.Drawing.Size(90,20)
$NewDisplayNameTitleBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewDisplayNameTitleBox.Text = "Display Name"
$NewOptionalGroupBox.Controls.Add($NewDisplayNameTitleBox)

$NewDisplayNameTextBox = New-Object System.Windows.Forms.TextBox
$NewDisplayNameTextBox.Location = New-Object System.Drawing.Size(130,25)
$NewDisplayNameTextBox.Size = New-Object System.Drawing.Size(270,20)
$NewDisplayNameTextBox.Multiline = $False
$NewDisplayNameTextBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewDisplayNameTextBox.ReadOnly = $False
$NewDisplayNameTextBox.Text = ""
$NewOptionalGroupBox.Controls.Add($NewDisplayNameTextBox)

$NewDisplayNameWarningBox = New-Object System.Windows.Forms.Label
$NewDisplayNameWarningBox.Location = New-Object System.Drawing.Size(15,50)
$NewDisplayNameWarningBox.Size = New-Object System.Drawing.Size(460,20)
$NewDisplayNameWarningBox.Font = New-Object System.Drawing.Font("Arial", "8")
$NewDisplayNameWarningBox.Text = "(Careful: If you nominate a DN, a new Display Name here will overwrite the existing name)"
$NewOptionalGroupBox.Controls.Add($NewDisplayNameWarningBox)

$ToolTip = New-Object System.Windows.Forms.ToolTip
$ToolTip.BackColor = [System.Drawing.Color]::LightGoldenrodYellow
$ToolTip.IsBalloon = $True
$ToolTip.InitialDelay = 500
$ToolTip.ReshowDelay = 500
$ToolTip.SetToolTip($NewDisplayNameTextBox, "Some punctuation characters disallowed. 64 character limit") 
$ToolTip.SetToolTip($NewDisplayNameTitleBox, "Some punctuation characters disallowed. 64 character limit") 


# ============================================================================
# Validate the Display Name ==================================================
# ============================================================================
$NewDisplayNameTextBox.Add_TextChanged(
{ 
	# Reference: http://support.microsoft.com/kb/909264 - a period is acceptable, but not as the first character:
	if (($NewDisplayNameTextBox.Text.Length -ge 1) -and ($NewDisplayNameTextBox.Text.SubString(0,1) -eq "."))
	{
		$NewDisplayNameTextBox.Text = $NewDisplayNameTextBox.Text.Remove(0,1)
	}
	# Reference: http://technet.microsoft.com/en-us/library/bb726984.aspx - block invalid characters:
	$NewDisplayNameTextBox.Text = [regex]::replace($NewDisplayNameTextBox.Text, '["/\\[\]:;|=,+*?<>]' , "")
	#Limit the input to 64 characters
	if ($NewDisplayNameTextBox.Text.Length -gt 64) 
	{
		$NewDisplayNameTextBox.Text = $NewDisplayNameTextBox.Text.SubString(0,64)
	}
	$NewDisplayNameTextBox.SelectionStart = $NewDisplayNameTextBox.Text.Length
})


# ============================================================================
# Add the Label & Text box containing the SIP URI ===========================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the SIP URI"
$NewSipUriTitleBox = New-Object System.Windows.Forms.Label
$NewSipUriTitleBox.Location = New-Object System.Drawing.Size(15,77)
$NewSipUriTitleBox.Size = New-Object System.Drawing.Size(85,20)
$NewSipUriTitleBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewSipUriTitleBox.Text = "SIP Address"
$NewOptionalGroupBox.Controls.Add($NewSipUriTitleBox)

$NewSipUriSipBox = New-Object System.Windows.Forms.Label
$NewSipUriSipBox.Location = New-Object System.Drawing.Size(100,77)
$NewSipUriSipBox.Size = New-Object System.Drawing.Size(30,20)
$NewSipUriSipBox.Font = New-Object System.Drawing.Font("Arial", "10",[System.Drawing.FontStyle]::Bold)
$NewSipUriSipBox.Text = "sip:"
$NewOptionalGroupBox.Controls.Add($NewSipUriSipBox)

$NewSipUriTextBox = New-Object System.Windows.Forms.TextBox
$NewSipUriTextBox.Location = New-Object System.Drawing.Size(130,75)
$NewSipUriTextBox.Size = New-Object System.Drawing.Size(270,20)
$NewSipUriTextBox.Multiline = $False
$NewSipUriTextBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewSipUriTextBox.ReadOnly = $False
$NewSipUriTextBox.Text = ""
$NewOptionalGroupBox.Controls.Add($NewSipUriTextBox)

$NewSipUriAtBox = New-Object System.Windows.Forms.Label
$NewSipUriAtBox.Location = New-Object System.Drawing.Size(402,77)
$NewSipUriAtBox.Size = New-Object System.Drawing.Size(16,20)
$NewSipUriAtBox.Font = New-Object System.Drawing.Font("Arial", "9")
$NewSipUriAtBox.Text = "@"
$NewOptionalGroupBox.Controls.Add($NewSipUriAtBox)

$SipDomainCombobox = New-Object System.Windows.Forms.Combobox
$SipDomainCombobox.Location = New-Object System.Drawing.Size(420,75)
$SipDomainCombobox.Size = New-Object System.Drawing.Size(150,20)
$SipDomainCombobox.Font = New-Object System.Drawing.Font("Arial", "9")
$SipDomainCombobox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
foreach ($item in $Global:SipDomains)
{   
	[void] $SipDomainCombobox.Items.Add($item.Identity)
}
# If there's a default, select it:
if ($Global:DefaultSIPDomain -ne "")
{	
	$SipDomainCombobox.SelectedIndex = $SipDomainCombobox.findstring("$Global:DefaultSIPDomain")
}
# If the above failed to match or there's no default, select the first one in the list:
if ($SipDomainCombobox.SelectedIndex -eq -1)
{
	$SipDomainCombobox.SelectedIndex = 0      # Automatically select the first one
}
$SipDomainCombobox.Add_SelectedIndexChanged({$Global:DefaultSIPDomain = $SipDomainCombobox.SelectedItem})
$NewOptionalGroupBox.Controls.Add($SipDomainCombobox)


# ============================================================================
# Validate SIP text values ===================================================
# ============================================================================
$NewSipUriTextBox.Add_TextChanged(
{ 
	#Reference http://www.ietf.org/rfc/rfc3261.txt
	#Reference http://www.ietf.org/rfc/rfc2396.txt
	# + tests of what the Lync Control Panel will accept
	$NewSipUriTextBox.Text = [regex]::replace($NewSipUriTextBox.Text, "[^\w\.\-+;!~*()\047]" , "") # "\047" = single quote
	$NewSipUriTextBox.SelectionStart = $NewSipUriTextBox.Text.Length
})


# ============================================================================
# Display Number =============================================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the Display Number"
$NewDisplayNumberTitleBox = New-Object System.Windows.Forms.Label
$NewDisplayNumberTitleBox.Location = New-Object System.Drawing.Size(15,112)
$NewDisplayNumberTitleBox.Size = New-Object System.Drawing.Size(105,20)
$NewDisplayNumberTitleBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewDisplayNumberTitleBox.Text = "Display Number"
$NewOptionalGroupBox.Controls.Add($NewDisplayNumberTitleBox)

$NewDisplayNumberTextBox = New-Object System.Windows.Forms.TextBox
$NewDisplayNumberTextBox.Location = New-Object System.Drawing.Size(130,110)
$NewDisplayNumberTextBox.Size = New-Object System.Drawing.Size(270,20)
$NewDisplayNumberTextBox.Multiline = $False
$NewDisplayNumberTextBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewDisplayNumberTextBox.ReadOnly = $False
$NewDisplayNumberTextBox.Text = ""
$NewOptionalGroupBox.Controls.Add($NewDisplayNumberTextBox)


# ============================================================================
# Validate the Display Number ================================================
# ============================================================================
$NewDisplayNumberTextBox.Add_TextChanged(
{ 
	# Only accept digits 0-9, "+", space, dash and round brackets in this field
	$NewDisplayNumberTextBox.Text = [regex]::replace($NewDisplayNumberTextBox.Text, "([^0-9+ ()-])" , "")
	$NewDisplayNumberTextBox.SelectionStart = $NewDisplayNumberTextBox.Text.Length
})


# ============================================================================
# Description ================================================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the Description"
$NewDescriptionTitleBox = New-Object System.Windows.Forms.Label
$NewDescriptionTitleBox.Location = New-Object System.Drawing.Size(15,142)
$NewDescriptionTitleBox.Size = New-Object System.Drawing.Size(105,20)
$NewDescriptionTitleBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewDescriptionTitleBox.Text = "Description"
$NewOptionalGroupBox.Controls.Add($NewDescriptionTitleBox)

$NewDescriptionTextBox = New-Object System.Windows.Forms.TextBox
$NewDescriptionTextBox.Location = New-Object System.Drawing.Size(130,140)
$NewDescriptionTextBox.Size = New-Object System.Drawing.Size(436,20)
$NewDescriptionTextBox.Multiline = $False
$NewDescriptionTextBox.Font = New-Object System.Drawing.Font("Arial", "10")
$NewDescriptionTextBox.Enabled = $False	# We default to Analog Device, so this box is disabled on launch
$NewDescriptionTextBox.Text = ""
$NewOptionalGroupBox.Controls.Add($NewDescriptionTextBox)


# ============================================================================
# Validate the Description ===================================================
# ============================================================================
$NewDescriptionTextBox.Add_TextChanged(
{ 
	#Stop the user adding any quotation marks.
	$NewDescriptionTextBox.Text = [regex]::replace($NewDescriptionTextBox.Text, '"' , "")
	#Limit the input to 1024 characters (unlikely??)
	if ($NewDescriptionTextBox.Text.Length -gt 1024) 
	{
		$NewDescriptionTextBox.Text = $NewDescriptionTextBox.Text.SubString(0,1024)
	}
	$NewDescriptionTextBox.SelectionStart = $NewDescriptionTextBox.Text.Length
})


# ============================================================================
# Add the listbox containing the Get-CsAnalogDevices =========================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the Analog Devices list box"
$AnalogDeviceTitleBox = New-Object System.Windows.Forms.Label
$AnalogDeviceTitleBox.Location = New-Object System.Drawing.Size(15,15)
$AnalogDeviceTitleBox.Size = New-Object System.Drawing.Size(160,20)
$AnalogDeviceTitleBox.Font = New-Object System.Drawing.Font("Arial", "10",[System.Drawing.FontStyle]::Bold)
$AnalogDeviceTitleBox.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$AnalogDeviceTitleBox.Text = "Analog Devices"
$TabPageSet.Controls.Add($AnalogDeviceTitleBox)

$AnalogDeviceCount = New-Object System.Windows.Forms.Label
$AnalogDeviceCount.Location = New-Object System.Drawing.Size(180,15)
$AnalogDeviceCount.Size = New-Object System.Drawing.Size(40,20)
$AnalogDeviceCount.Font = New-Object System.Drawing.Font("Arial", "10")
$AnalogDeviceCount.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$AnalogDeviceCount.Text = ""
$TabPageSet.Controls.Add($AnalogDeviceCount)

$AnalogDeviceListbox = New-Object System.Windows.Forms.Listbox
$AnalogDeviceListbox.Location = New-Object System.Drawing.Size(15,40)
$AnalogDeviceListbox.Size = New-Object System.Drawing.Size(460,120)
$AnalogDeviceListbox.HorizontalScrollbar = $true
$AnalogDeviceListbox.SelectionMode = "MultiExtended"
$AnalogDeviceListbox.TabIndex = 1
$TabPageSet.Controls.Add($AnalogDeviceListbox)

$AnalogDeviceListbox.Add_SelectedIndexChanged({Update-Display})


# ============================================================================
# Add the Existing Items tab's FILTER button =================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the Filter button"
$ExistFilterButton = New-Object System.Windows.Forms.Button
$ExistFilterButton.Name = "Filter"
$ExistFilterButton.Text = "Filter"
$ExistFilterButton.TabIndex = 3
$ExistFilterButton.Location = New-Object System.Drawing.Size(280,10)
$ExistFilterButton.Size = New-Object System.Drawing.Size(90,25)
$ExistFilterButton.Add_Click({
	$TempFilterOu = $Global:FilterOu
	Show-FilterForm ($Global:FilterOu)
	if ($TempFilterOu -ne $Global:FilterOu)
	{
		#The user change the filter OU. Update the listboxes:
		Update-DeviceList
	}
	Update-ADCAPLists	#This applies the filter - may be redundant but due to the structure I have no option but to re-test
	})
$TabPageSet.Controls.Add($ExistFilterButton)


# ============================================================================
# Add the REFRESH button =====================================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the Refresh button"
$RefreshButton = New-Object System.Windows.Forms.Button
$RefreshButton.Name = "Refresh"
$RefreshButton.Text = "Refresh"
$RefreshButton.TabIndex = 4
$RefreshButton.Location = New-Object System.Drawing.Size(385, 10)
$RefreshButton.Size = New-Object System.Drawing.Size(90,25)

$RefreshButton.Add_Click({
	Update-DeviceList
	Update-ADCAPLists
	})
$TabPageSet.Controls.Add($RefreshButton)


# ============================================================================
# Add the listbox containing the Get-CsCommonAreaPhones ======================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the CAPs list box"
$CommonAreaPhoneTitleBox = New-Object System.Windows.Forms.Label
$CommonAreaPhoneTitleBox.Location = New-Object System.Drawing.Size(15,155)
$CommonAreaPhoneTitleBox.Size = New-Object System.Drawing.Size(160,20)
$CommonAreaPhoneTitleBox.Font = New-Object System.Drawing.Font("Arial", "10",[System.Drawing.FontStyle]::Bold)
$CommonAreaPhoneTitleBox.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$CommonAreaPhoneTitleBox.Text = "Common Area Phones"
$TabPageSet.Controls.Add($CommonAreaPhoneTitleBox)

$CommonAreaPhoneCount = New-Object System.Windows.Forms.Label
$CommonAreaPhoneCount.Location = New-Object System.Drawing.Size(180,155)
$CommonAreaPhoneCount.Size = New-Object System.Drawing.Size(40,20)
$CommonAreaPhoneCount.Font = New-Object System.Drawing.Font("Arial", "10")
$CommonAreaPhoneCount.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$CommonAreaPhoneCount.Text = ""
$TabPageSet.Controls.Add($CommonAreaPhoneCount)

$CommonAreaPhoneListbox = New-Object System.Windows.Forms.Listbox
$CommonAreaPhoneListbox.Location = New-Object System.Drawing.Size(15,180)
$CommonAreaPhoneListbox.Size = New-Object System.Drawing.Size(460,120)
$CommonAreaPhoneListbox.HorizontalScrollbar = $true
$CommonAreaPhoneListbox.SelectionMode = "MultiExtended"
$CommonAreaPhoneListbox.TabIndex = 2
$TabPageSet.Controls.Add($CommonAreaPhoneListbox)

$CommonAreaPhoneListbox.Add_SelectedIndexChanged({Update-Display})


# ============================================================================
# Add the textbox displaying *1* item ========================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the text box displaying 1 item"
$OutputBox = New-Object System.Windows.Forms.TextBox
$OutputBox.Location = New-Object System.Drawing.Size(15,320)
$OutputBox.Size = New-Object System.Drawing.Size(460,160)
$OutputBox.Multiline = $True
$OutputBox.Font = New-Object System.Drawing.Font("Courier New", "10")
$OutputBox.Wordwrap = $True
$OutputBox.ReadOnly = $True
$OutputBox.TabStop = $False
$OutputBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$OutputBox.Text = "Select an Analog Device or Common Area Phone to display it"
$TabPageSet.Controls.Add($OutputBox)

function Update-Display
{ 
	$AD_count = 0
	$CA_count = 0
	foreach ($AD in $AnalogDeviceListbox.SelectedItems)
	{
		if ($AD -ne "") {$AD_count++}
	}
	foreach ($CA in $CommonAreaPhoneListbox.SelectedItems)
	{
		if ($CA -ne "") {$CA_count++}
	}
	if (($AD_count -eq 1) -and ($CA_count -eq 0))
	{
		$AD_name = $AnalogDeviceListbox.SelectedItem
		if ($AD_name -ne "")
		{
			# Then they've not selected the first, empty value
			$SelectedAD = $Global:ADs | where-object {$_.Identity -match $AD_name}
			$OutputBox.Text = ($SelectedAD | Format-List | Out-String).Trim()
			if ($SelectedAD.DialPlan -eq $null)
			{
				$DialPlanListbox.SelectedIndex = $DialPlanListbox.findstring("<Automatic>")
			}
			else
			{
				$DialPlanListbox.SelectedIndex = $DialPlanListbox.findstring("$($SelectedAD.DialPlan)")
			}
			if ($SelectedAD.VoicePolicy -eq $null)
			{
				$VoicePolicyListbox.SelectedIndex = $VoicePolicyListbox.findstring("<Automatic>")
			}
			else
			{
				$VoicePolicyListbox.SelectedIndex = $VoicePolicyListbox.findstring("$($SelectedAD.VoicePolicy)")
			}
			if ($SelectedAD.ClientPolicy -eq $null)
			{
				$ClientPolicyListbox.SelectedIndex = $ClientPolicyListbox.findstring("<Automatic>")
			}
			else
			{
				$ClientPolicyListbox.SelectedIndex = $ClientPolicyListbox.findstring("$($SelectedAD.ClientPolicy)")
			}
			if ($SelectedAD.LocationPolicy -eq $null)
			{
				$LocationPolicyListbox.SelectedIndex = $LocationPolicyListbox.findstring("<Automatic>")
			}
			else
			{
				$LocationPolicyListbox.SelectedIndex = $LocationPolicyListbox.findstring("$($SelectedAD.LocationPolicy)")
			}
		}
	}
	elseif (($AD_count -eq 0) -and ($CA_count -eq 1))
	{
		$CA_name = $CommonAreaPhoneListbox.SelectedItem
		if ($CA_name -ne "")
		{
			# Then they've not selected the first, empty value
			$SelectedCAP = $Global:CAPs | where-object {$_.Identity -match $CA_name}
			$OutputBox.Text = ($SelectedCAP | Format-List | Out-String).Trim()
			if ($SelectedCAP.DialPlan -eq $null)
			{
				$DialPlanListbox.SelectedIndex = $DialPlanListbox.findstring("<Automatic>")
			}
			else
			{
				$DialPlanListbox.SelectedIndex = $DialPlanListbox.findstring("$($SelectedCAP.DialPlan)")
			}
			if ($SelectedCAP.VoicePolicy -eq $null)
			{
				$VoicePolicyListbox.SelectedIndex = $VoicePolicyListbox.findstring("<Automatic>")
			}
			else
			{
				$VoicePolicyListbox.SelectedIndex = $VoicePolicyListbox.findstring("$($SelectedCAP.VoicePolicy)")
			}
			if ($SelectedCAP.ClientPolicy -eq $null)
			{
				$ClientPolicyListbox.SelectedIndex = $ClientPolicyListbox.findstring("<Automatic>")
			}
			else
			{
				$ClientPolicyListbox.SelectedIndex = $ClientPolicyListbox.findstring("$($SelectedCAP.ClientPolicy)")
			}
			if ($SelectedCAP.LocationPolicy -eq $null)
			{
				$LocationPolicyListbox.SelectedIndex = $LocationPolicyListbox.findstring("<Automatic>")
			}
			else
			{
				$LocationPolicyListbox.SelectedIndex = $LocationPolicyListbox.findstring("$($SelectedCAP.LocationPolicy)")
			}
		}
	}
	else
	{
		$OutputBox.Text = "Select only 1 Analog Device or Common Area Phone to display it"
		#This de-selects all the policies
		$DialPlanListbox.SelectedIndex = -1
		$VoicePolicyListbox.SelectedIndex = -1
		$ClientPolicyListbox.SelectedIndex = -1
		$LocationPolicyListbox.SelectedIndex = -1
	}
	# write-host "AD Count = $($AD_count), CA Count = $($CA_Count)"
	handler_ValidateGo
}


# ============================================================================
# Event Handler that reads *all* the Analog Devices & Common Area Phones =====
# ============================================================================
function Update-DeviceList
{ 
	# Flush all lists first:
	$Global:ADs = @()
	$Global:CAPs = @()
	
	write-verbose "Updating device lists. Filtering on ""$($Global:FilterOu)"""
	
	# Build the Analog Device list
	try
	{
		if ($Global:FilterOu -ne "")
		{
			$Global:ADs = Invoke-Expression "Get-CsAnalogDevice -OU ""$Global:FilterOu"" -warningaction silentlycontinue"
		}
		else
		{
			$Global:ADs = Invoke-Expression "Get-CsAnalogDevice -warningaction silentlycontinue"
		}
	}
	catch 
	{
		# Most likely issue is no Analogs in the selected OU
		$Global:ADs = @()
	}
	
	# Build the Common Area Phones list
	try
	{
		if ($Global:FilterOu -ne "")
		{
			$Global:CAPs = Invoke-Expression "Get-CsCommonAreaPhone -OU ""$Global:FilterOu"" -warningaction silentlycontinue" 
		}
		else
		{
			$Global:CAPs = Invoke-Expression "Get-CsCommonAreaPhone -warningaction silentlycontinue" 
		}
	}
	catch 
	{
		# Most likely issue is no CAPs in the selected OU
		$Global:CAPs = @()
	}
}	

# ============================================================================
# Update the lists of ADs & CAPs (after Refresh or a Filter is applied ========
# ============================================================================
function Update-ADCAPLists
{
	$AnalogDeviceListbox.Items.Clear()
	$CommonAreaPhoneListbox.Items.Clear()
	
	$FilteredADList = Filter-Objects $Global:ADs
	#(Re)build the list of ADs:
	[void] $AnalogDeviceListbox.Items.Add("")	#Add the blank entry at the top
	$selectedIndex = 0
	$ItemIndex = 1
	foreach ($item in $FilteredADList)
	{   
		if ($item -ne $null) 
		{
			if ($global:LastTouchedObject -ne "")
			{
				if ($global:LastTouchedObject -match $item.DistinguishedName)
				{
					$selectedIndex = $ItemIndex
				}
			}
			[void] $AnalogDeviceListbox.Items.Add($item.Identity)
			$ItemIndex ++
		}
	}
	$AnalogDeviceListbox.SetSelected($selectedIndex,$true)
	$AnalogDeviceListbox.TopIndex = ($AnalogDeviceListbox.SelectedIndex)
	$AnalogDeviceCount.Text = $AnalogDeviceListbox.Items.Count - 1

	$FilteredCAPList = Filter-Objects $Global:CAPs	
	#(Re)build the list of CAPs:
	[void] $CommonAreaPhoneListbox.Items.Add("")	#Add the blank entry at the top
	$selectedIndex = 0
	$ItemIndex = 1
	foreach ($item in $FilteredCAPList)
	{   
		if ($item -ne $null) 
		{
			if ($global:LastTouchedObject -ne "")
			{
				if ($global:LastTouchedObject -match $item.DistinguishedName)
				{
					$selectedIndex = $ItemIndex
				}
			}
			[void] $CommonAreaPhoneListbox.Items.Add($item.Identity)
			$ItemIndex ++
		}
	}
	$CommonAreaPhoneListbox.SetSelected($selectedIndex,$true)
	$CommonAreaPhoneListbox.TopIndex = ($CommonAreaPhoneListbox.SelectedIndex)
	$CommonAreaPhoneCount.Text = $CommonAreaPhoneListbox.Items.Count - 1
}


# ============================================================================
# Add the listbox containing the Get-CsDialPlans =============================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the Dial Plans list box"
$DialPlanTitleBox = New-Object System.Windows.Forms.Label
$DialPlanTitleBox.Location = New-Object System.Drawing.Size(500,15)
$DialPlanTitleBox.Size = New-Object System.Drawing.Size(220,20)
$DialPlanTitleBox.Font = New-Object System.Drawing.Font("Arial", "10",[System.Drawing.FontStyle]::Bold)
$DialPlanTitleBox.Text = "Dial Plans"
$TabPageSet.Controls.Add($DialPlanTitleBox)

$DialPlanListbox = New-Object System.Windows.Forms.Listbox
$DialPlanListbox.Location = New-Object System.Drawing.Size(500,35)
$DialPlanListbox.Size = New-Object System.Drawing.Size(220,90)
$DialPlanListbox.TabIndex = 5
try
{
	$items = Invoke-Expression "Get-CsDialPlan -warningaction silentlycontinue -verbose:`$false | Select-Object Identity"
}
catch 
{
	$items = ""
	log ("Error populating Dial Plans: $_")
}
[void] $DialPlanListbox.Items.Add("<No Change>")
[void] $DialPlanListbox.Items.Add("<Automatic>")

foreach ($item in $items)
{
	$result = Filter-PolicyNames $item.Identity
	if ($result -ne $null) { [void] $DialPlanListbox.Items.Add($result) }
}
$DialPlanListbox.SetSelected(0,$true)
$TabPageSet.Controls.Add($DialPlanListbox)


# ============================================================================
# Add the listbox containing the Get-CsVoicePolicies =========================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the Voice Policies list box"
$VoicePoliciesTitleBox = New-Object System.Windows.Forms.Label
$VoicePoliciesTitleBox.Location = New-Object System.Drawing.Size(500,125)
$VoicePoliciesTitleBox.Size = New-Object System.Drawing.Size(220,20)
$VoicePoliciesTitleBox.Font = New-Object System.Drawing.Font("Arial", "10",[System.Drawing.FontStyle]::Bold)
$VoicePoliciesTitleBox.Text = "Voice Policies"
$TabPageSet.Controls.Add($VoicePoliciesTitleBox)

$VoicePolicyListbox = New-Object System.Windows.Forms.Listbox
$VoicePolicyListbox.Location = New-Object System.Drawing.Size(500,145)
$VoicePolicyListbox.Size = New-Object System.Drawing.Size(220,90)
$VoicePolicyListbox.TabIndex = 6
try
{
$items = Invoke-Expression "Get-CsVoicePolicy -warningaction silentlycontinue -verbose:`$false | Select-Object Identity"
}
catch 
{
	$items = ""
	log ("Error populating Voice Policies: $_")
}
[void] $VoicePolicyListbox.Items.Add("<No Change>")
[void] $VoicePolicyListbox.Items.Add("<Automatic>")
foreach ($item in $items)
{ 
	$result = Filter-PolicyNames $item.Identity
	if ($result -ne $null) { [void] $VoicePolicyListbox.Items.Add($result) }
}
$VoicePolicyListbox.SetSelected(0,$true)
$TabPageSet.Controls.Add($VoicePolicyListbox)


# ============================================================================
# Add the listbox containing the Get-CsClientPolicies ========================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the Client Policies list box"
$ClientPoliciesTitleBox = New-Object System.Windows.Forms.Label
$ClientPoliciesTitleBox.Location = New-Object System.Drawing.Size(500,235)
$ClientPoliciesTitleBox.Size = New-Object System.Drawing.Size(220,20)
$ClientPoliciesTitleBox.Font = New-Object System.Drawing.Font("Arial", "10",[System.Drawing.FontStyle]::Bold)
$ClientPoliciesTitleBox.Text = "Client Policies"
$TabPageSet.Controls.Add($ClientPoliciesTitleBox)

$ClientPolicyListbox = New-Object System.Windows.Forms.Listbox
$ClientPolicyListbox.Location = New-Object System.Drawing.Size(500,255)
$ClientPolicyListbox.Size = New-Object System.Drawing.Size(220,90)
$ClientPolicyListbox.TabIndex = 7

try
{
	$items = Invoke-Expression "Get-CsClientPolicy -warningaction silentlycontinue -verbose:`$false | Select-Object Identity"
}
catch 
{
	$items = ""
	log ("Error populating Client Policies: $_")
}
[void] $ClientPolicyListbox.Items.Add("<No Change>")
[void] $ClientPolicyListbox.Items.Add("<Automatic>")
foreach ($item in $items)
{   
  	$result = Filter-PolicyNames $item.Identity
	if ($result -ne $null) { [void] $ClientPolicyListbox.Items.Add($result) }
}
$ClientPolicyListbox.SetSelected(0,$true)
$TabPageSet.Controls.Add($ClientPolicyListbox)

# ============================================================================
# Add the listbox containing the Get-CsLocationPolicies ========================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the Location Policies list box"
$LocationPoliciesTitleBox = New-Object System.Windows.Forms.Label
$LocationPoliciesTitleBox.Location = New-Object System.Drawing.Size(500,345)
$LocationPoliciesTitleBox.Size = New-Object System.Drawing.Size(220,20)
$LocationPoliciesTitleBox.Font = New-Object System.Drawing.Font("Arial", "10",[System.Drawing.FontStyle]::Bold)
$LocationPoliciesTitleBox.Text = "Location Policies"
$TabPageSet.Controls.Add($LocationPoliciesTitleBox)

$LocationPolicyListbox = New-Object System.Windows.Forms.Listbox
$LocationPolicyListbox.Location = New-Object System.Drawing.Size(500,365)
$LocationPolicyListbox.Size = New-Object System.Drawing.Size(220,90)
$LocationPolicyListbox.TabIndex = 8

try
{
	$items = Invoke-Expression "Get-CsLocationPolicy -warningaction silentlycontinue -verbose:`$false | Select-Object Identity"
}
catch 
{
	$items = ""
	log ("Error populating Client Policies: $_")
}
[void] $LocationPolicyListbox.Items.Add("<No Change>")
[void] $LocationPolicyListbox.Items.Add("<Automatic>")
foreach ($item in $items)
{   
  	$result = Filter-PolicyNames $item.Identity
	if ($result -ne $null) { [void] $LocationPolicyListbox.Items.Add($result) }
}
$LocationPolicyListbox.SetSelected(0,$true)
$TabPageSet.Controls.Add($LocationPolicyListbox)


# ============================================================================
# Add the PIN text box =======================================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the PIN text box"
$PinTitleBox = New-Object System.Windows.Forms.Label
$PinTitleBox.Location = New-Object System.Drawing.Size(500,457)
$PinTitleBox.Size = New-Object System.Drawing.Size(65,20)
$PinTitleBox.Font = New-Object System.Drawing.Font("Arial", "10",[System.Drawing.FontStyle]::Bold)
$PinTitleBox.Text = "User PIN"
$PinTitleBox.Enabled = $False
$TabPageSet.Controls.Add($PinTitleBox)

$PinTextBox = New-Object System.Windows.Forms.TextBox
$PinTextBox.Location = New-Object System.Drawing.Size(570,455)
$PinTextBox.Size = New-Object System.Drawing.Size(100,20)
$PinTextBox.Multiline = $False
$PinTextBox.Font = New-Object System.Drawing.Font("Arial", "10")
$PinTextBox.ReadOnly = $True
$PinTextBox.TabIndex = 9
$TabPageSet.Controls.Add($PinTextBox)


# ============================================================================
# Validate PIN text values ===================================================
# ============================================================================
$PinTextBox.Add_TextChanged(
{ 
	# Only accept digits 0-9 in this field
	$PinTextBox.Text = [regex]::replace($PinTextBox.Text, "\D" , "")
	$PinTextBox.SelectionStart = $PinTextBox.Text.Length
})


# ============================================================================
# Event Handler that unlocks the GO button ===================================
# ============================================================================
function handler_ValidateGo 
{ 
	if ($TabControl.SelectedIndex -eq 0)
	{
		# "New"
		$valid = 1	#We start with the assumption that the test will be true. 
					# Any of those that fail will reset it to 0 & disable the Go Button
		if ($NewLineUriTextBox.Text -eq "") {$valid = 0}
		
		if (($NewOuTextBox.Text -eq "") -and ($NewDnTextBox.Text -eq "")) {$valid = 0} #At least 1 must be populated!
		if (($NewDnTextBox.Text -eq "") -and ($NewOuTextBox.Text -ne "") -and ($NewOuTextBox.Text -notlike "*OU=*")) {$valid = 0} #If OU, it must contain "OU="
		if (($NewDnTextBox.Text -ne "") -and ($NewDnTextBox.Text -notlike "*CN=*")) {$valid = 0} #If DN, it must contain "CN="

		if ($valid -eq 1)
		{
			$GoButton.Enabled = $True
		}
		else
		{
			$GoButton.Enabled = $False
		}
		$PinTitleBox.Enabled = $False	#Always disabled for New entities
		$PinTextBox.ReadOnly = $True		# "
	}
	else
	{
		# "Existing"
		$GoButton.Enabled = $False
		foreach ($AD in $AnalogDeviceListbox.SelectedItems)
		{
			if ($AD -ne "")
			{
				$GoButton.Enabled = $True
				$PinTitleBox.Enabled = $False	#Always disabled for Analogs
				$PinTextBox.ReadOnly = $True		# "
			}
		}
		foreach ($CA in $CommonAreaPhoneListbox.SelectedItems)
		{
			if ($CA -ne "") 
			{
				$GoButton.Enabled = $True
				$PinTitleBox.Enabled = $True		#Always enabled for CAPs
				$PinTextBox.ReadOnly = $False	# "
			}
		}
	}
}


# ============================================================================
# Add the 'Monitor Window' frame =============================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the Monitor window"
$StatusBox = New-Object System.Windows.Forms.RichTextBox
$StatusBox.Location = New-Object System.Drawing.Size(15,570)
$StatusBox.Size = New-Object System.Drawing.Size(500,65)
$StatusBox.Multiline = $True
$StatusBox.Font = New-Object System.Drawing.Font("Courier New", "9")
#$StatusBox.ForeColor = [System.Drawing.Color]::"Black"
$StatusBox.Wordwrap = $False
$StatusBox.ReadOnly = $True
$StatusBox.TabStop = $False
#$StatusBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
$StatusBox.Text = "Monitor Window"
$StatusBox.Add_TextChanged(
{
	# Scroll to the most recent entry when-ever text is added
	$StatusBox.SelectionStart = $StatusBox.Text.Length
	$StatusBox.ScrollToCaret()
}
)
$Form.Controls.Add($StatusBox)


# ============================================================================
# Create the LOGGING filename and Add the check-box ===========================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Create the logging filename and checkbox"
$date = (get-date).ToString('yyyyMMMdd') #Thanks Kathy: http://blogs.msdn.com/b/kathykam/archive/2006/09/29/.net-format-string-102_3a00_-datetime-format-string.aspx
$LogFile = "MAdCaP-" + $date + ".log"
$Log = $True #The logging code will refer to this Boolean

$LogCheckbox = New-Object System.Windows.Forms.Checkbox
$LogCheckbox.Name = "Log"
$LogCheckbox.Text = "Log to file $LogFile"
$LogCheckbox.Checked = $True
$LogCheckbox.TabIndex = 10
$LogCheckbox.Size = "290,30"
$LogCheckbox.Location = "540, 570"
$LogCheckbox.Add_CheckedChanged(
{
	if ($LogCheckbox.Checked -eq $True)
	{
		$Log = $True
	}
	else
	{
		$Log = $False
	}
})
$Form.Controls.Add($LogCheckbox)


# ============================================================================
# Function SEND ==============================================================
# 1) Logs the string to be sent ==============================================
# 2) Appends the error redirector " 2>&1" to it ==============================
# 3) Sends it! ===============================================================
# 4) Logs the response =======================================================
# 5) Returns the response in case the calling code wants it ==================
# ============================================================================
function Send () {
param ([String]$data)

Log ("Command Executed = $data")
$data += " 2>&1"	# Append the handler that will capture errors

try
{
	$response = Invoke-Expression "$data -verbose:`$false" 	#"$Response" will contain the object's name or blank if we succeeded, otherwise error text.
}
catch 
{
	$response = "Error caught by handler: $_"
}

Log ("Result  Received = $response")
return [string]$response
}


# ============================================================================
# Function LOG ===============================================================
# Writes to the status display, and to the log file if enabled ===============
# ============================================================================
function Log() {
param ([String]$data)

$CRLF = [System.Environment]::NewLine
$StatusBox.Text +=  $CRLF + $Data

if ($Log -eq $True)
{
	try
	{
		$Time = (get-date).ToString("HH:mm:ss") # Use "hh:mm:ss tt" if you want the time in 12-hour format: "01:09:42 PM"
		Write "$Time $data" | Out-File $Logfile -Append
	}
	catch
	{
		#Log the failure to screen if we're unable to write to the log file
		$StatusBox.Text +=  $CRLF + "MAdCaP ERROR: Unable to write to log file"
	}
}
}


# ============================================================================
# Add the GO button ==========================================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the GO! button"
$GoButton = New-Object System.Windows.Forms.Button
$GoButton.Name = "GO"
$GoButton.Text = "GO!"
$GoButton.TabIndex = 11
$GoButton.Size = "90,35"
$GoButton.Location = "540, 600"
$GoButton.Enabled = $False
$GoButton.Add_Click(
{
	if ($TabControl.SelectedIndex -eq 0)
	{
		# We're creating a NEW Object
		if ($RadiobuttonAD.Checked -eq $True)
		{
			# We're creating an Analog Device
			$transmit = "New-CsAnalogDevice "
			
			#Analog Fax
			$text = $NewAnalogFaxComboBox.SelectedItem
			if ($text -eq "True")
			{
				$transmit += "-AnalogFax:`$True "
			}
			else
			{
				$transmit += "-AnalogFax:`$False "
			}

			#Gateway
			$text = $NewPstnGatewayCombobox.SelectedItem
			$transmit += "-Gateway ""$text"" "
		}
		else
		{
			# It's a Common Area Phone
			$transmit = "New-CsCommonAreaPhone "
		
			#Description              
			$text = $NewDescriptionTextBox.Text
			if ($text -ne "") 
			{
				$transmit += "-Description ""$text"" "
			}
		}
		
		#If DN is populated, use it, otherwise use OU
		$text = $NewDnTextBox.Text
		if ($text -ne "")
		{
			$transmit += "-DN ""$text"" "
		}
		else
		{
			$text = $NewOuTextBox.Text
			$transmit += "-OU ""$text"" "
		}
		
		#LineURI
		$text = $NewLineUriTextBox.Text
		$transmit += "-LineUri ""tel:$text"" "
		
		#Registrar Pool
		$text = $NewRegistrarPoolCombobox.SelectedItem
		$transmit += "-RegistrarPool ""$text"" "
						
		#Add the Optional parameters if they exist:
		#Sip Address
		$text = $NewSipUriTextBox.Text
		if ($text -ne "")
		{
			$domain = $SipDomainCombobox.SelectedItem
			$transmit += "-SipAddress ""sip:$text@$domain"" "
		}
		#DisplayName
		$text = $NewDisplayNameTextBox.Text
		if ($text -ne "")
		{
			$transmit += "-DisplayName ""$text"" "
		}
		#DisplayNumber              
		$text = $NewDisplayNumberTextBox.Text
		if ($text -ne "") 
		{
			$transmit += "-DisplayNumber ""$text"" "
		}
		$transmit.trim() # Remove the space from the end
		$result = send ($transmit)
		
		#"$result" will be the new object's identity if we created it OK, otherwise it'll contain an error message.
		if (($result -notlike '*Error*') -and ($result -notlike '*The object * already exists*'))
		{
			if ($result -like "CN=*")
			{
				$global:LastTouchedObject = $result 
				$OKMessage = "New object created OK!`n`nClick Refresh on the ""Existing Object"" tab until the new object shows (it will be automatically selected for you) then set the required Policies and/or PIN."
				$discard = [windows.forms.messagebox]::show($OKMessage,'MAdCaP.ps1','Ok')
			}
			else
			{
				$global:LastTouchedObject = ""
			}
		}
		else
		{
			$global:LastTouchedObject = ""
			$discard = [windows.forms.messagebox]::show($result,'MAdCaP.ps1','Ok')
		}
	}
	else
	{
	# "Existing Object"
	# Do the analogs first:
		if ($AnalogDeviceListbox.SelectedItems –ne $Null)
		{
			foreach ($AnalogDevice in $AnalogDeviceListbox.SelectedItems)
			{
				$SelectedAD = $Global:ADs | where-object {$_.Identity -match $AnalogDevice}
				Grant-Policy ($SelectedAD)
				# Analog phones don't have a client PIN
			}
		}
					
		# Now the Common Area Phones:
		if ($CommonAreaPhoneListbox.SelectedItems –ne $Null)
		{
			foreach ($CommonAreaPhone in $CommonAreaPhoneListbox.SelectedItems)
			{  
				$SelectedCAP = $Global:CAPs | where-object {$_.Identity -match $CommonAreaPhone}
				Grant-Policy ($SelectedCAP)
				if ($PinTextBox.Text -ne "")
				{
					$PIN = $PinTextBox.Text
					$transmit = "Set-CsClientPin ""$SelectedCAP"" -PIN ""$PIN"""
					$discard = send ($transmit) # We don't use the value returned here, hence "$discard"
				}
			}
		}
	}              
})
$Form.Controls.Add($GoButton)


# ============================================================================
# Function Grant-Policy =======================================================
# ============================================================================
function Grant-Policy()
{
	param ([Object]$device)

	$DialPlan = $DialPlanListbox.SelectedItem
	$Result = Test-ForPolicyChange $device "DialPlan" $DialPlan
	if ($Result -ne $null)
	{
		$transmit = "Grant-CsDialplan ""$device"" -PolicyName ""$Result"""
		$discard = send ($transmit) # We don't use the value returned here, hence "$discard"
	}
	$VoicePolicy = $VoicePolicyListbox.SelectedItem
	$Result = Test-ForPolicyChange $device "VoicePolicy" $VoicePolicy
	if ($Result -ne $null)
	{
		$transmit = "Grant-CsVoicepolicy ""$device"" -PolicyName ""$Result"""
		$discard = send ($transmit) # We don't use the value returned here, hence "$discard"
	}
	$ClientPolicy = $ClientPolicyListbox.SelectedItem
	$Result = Test-ForPolicyChange $device "ClientPolicy" $ClientPolicy
	if ($Result -ne $null)
	{
		$transmit = "Grant-CsClientpolicy ""$device"" -PolicyName ""$Result"""
		$discard = send ($transmit) # We don't use the value returned here, hence "$discard"
	}
	$LocationPolicy = $LocationPolicyListbox.SelectedItem
	$Result = Test-ForPolicyChange $device "LocationPolicy" $LocationPolicy
	if ($Result -ne $null)
	{	
		$transmit = "Grant-CsLocationPolicy ""$device"" -PolicyName ""$Result"""
		$discard = send ($transmit) # We don't use the value returned here, hence "$discard"
	}
}
	
	
# ============================================================================
# Add the CANCEL button ======================================================
# ============================================================================
write-progress -id 1 -Activity "Loading" -Status "Add the Cancel button"
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Name = "Cancel"
$CancelButton.Text = "Cancel"
$CancelButton.TabIndex = 12
$CancelButton.Size = "90,35"
$CancelButton.Location = "650, 600"
$CancelButton.Add_Click(
{
	if ($Log -eq $True)
	{
		Log ("============================= Clean exit =======================")
	}
	$Form.Close()
}
)
$Form.Controls.Add($CancelButton)


# ===========================================================================
# Activate the form =========================================================
# ===========================================================================
write-progress -id 1 -Activity "Loading" -Status "Activate the form"
if ($ShowExisting)
{
	$TabControl.SelectedIndex = 1
	Update-DeviceList	#Reads all objects from Lync/SfB
	Update-ADCAPLists	#Populates the two lists on the Existing Items tab
	Update-Display		#Shows the contents of one item to the pane & the RH Policy lists
}
$Form.Add_Shown({$Form.Activate()})
if ($Log -eq $True)
{
	Log ("============================== Launched ========================")
}
write-progress -id 1 -Activity "Loading" -Status "Activate the form" -complete	
write-verbose "Showing dialog"
[void] $Form.ShowDialog()
write-verbose "Writing Settings"
WriteSettings $Configfile $Global:DefaultOU $Global:DefaultPool $Global:DefaultSIPDomain
# END


#References:
# Corky Caragol's OU picker: https://gallery.technet.microsoft.com/Lync-2013-Common-Area-57bc4ff1


#Code signing certificate kindly provided by Digicert:
# SIG # Begin signature block
# MIIceAYJKoZIhvcNAQcCoIIcaTCCHGUCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUz6q2P8nAMjdqHm8/yAMdIltN
# vlCgghenMIIFMDCCBBigAwIBAgIQA1GDBusaADXxu0naTkLwYTANBgkqhkiG9w0B
# AQsFADByMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYD
# VQQLExB3d3cuZGlnaWNlcnQuY29tMTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFz
# c3VyZWQgSUQgQ29kZSBTaWduaW5nIENBMB4XDTIwMDQxNzAwMDAwMFoXDTIxMDcw
# MTEyMDAwMFowbTELMAkGA1UEBhMCQVUxGDAWBgNVBAgTD05ldyBTb3V0aCBXYWxl
# czESMBAGA1UEBxMJUGV0ZXJzaGFtMRcwFQYDVQQKEw5HcmVpZyBTaGVyaWRhbjEX
# MBUGA1UEAxMOR3JlaWcgU2hlcmlkYW4wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAw
# ggEKAoIBAQC0PMhHbI+fkQcYFNzZHgVAuyE3BErOYAVBsCjZgWFMhqvhEq08El/W
# PNdtlcOaTPMdyEibyJY8ZZTOepPVjtHGFPI08z5F6BkAmyJ7eFpR9EyCd6JRJZ9R
# ibq3e2mfqnv2wB0rOmRjnIX6XW6dMdfs/iFaSK4pJAqejme5Lcboea4ZJDCoWOK7
# bUWkoqlY+CazC/Cb48ZguPzacF5qHoDjmpeVS4/mRB4frPj56OvKns4Nf7gOZpQS
# 956BgagHr92iy3GkExAdr9ys5cDsTA49GwSabwpwDcgobJ+cYeBc1tGElWHVOx0F
# 24wBBfcDG8KL78bpqOzXhlsyDkOXKM21AgMBAAGjggHFMIIBwTAfBgNVHSMEGDAW
# gBRaxLl7KgqjpepxA8Bg+S32ZXUOWDAdBgNVHQ4EFgQUzBwyYxT+LFH+GuVtHo2S
# mSHS/N0wDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoGCCsGAQUFBwMDMHcGA1Ud
# HwRwMG4wNaAzoDGGL2h0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9zaGEyLWFzc3Vy
# ZWQtY3MtZzEuY3JsMDWgM6Axhi9odHRwOi8vY3JsNC5kaWdpY2VydC5jb20vc2hh
# Mi1hc3N1cmVkLWNzLWcxLmNybDBMBgNVHSAERTBDMDcGCWCGSAGG/WwDATAqMCgG
# CCsGAQUFBwIBFhxodHRwczovL3d3dy5kaWdpY2VydC5jb20vQ1BTMAgGBmeBDAEE
# ATCBhAYIKwYBBQUHAQEEeDB2MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wTgYIKwYBBQUHMAKGQmh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydFNIQTJBc3N1cmVkSURDb2RlU2lnbmluZ0NBLmNydDAMBgNVHRMB
# Af8EAjAAMA0GCSqGSIb3DQEBCwUAA4IBAQCtV/Nu/2vgu+rHGFI6gssYWfYLEwXO
# eJqOYcYYjb7dk5sRTninaUpKt4WPuFo9OroNOrw6bhvPKdzYArXLCGbnvi40LaJI
# AOr9+V/+rmVrHXcYxQiWLwKI5NKnzxB2sJzM0vpSzlj1+fa5kCnpKY6qeuv7QUCZ
# 1+tHunxKW2oF+mBD1MV2S4+Qgl4pT9q2ygh9DO5TPxC91lbuT5p1/flI/3dHBJd+
# KZ9vYGdsJO5vS4MscsCYTrRXvgvj0wl+Nwumowu4O0ROqLRdxCZ+1X6a5zNdrk4w
# Dbdznv3E3s3My8Axuaea4WHulgAvPosFrB44e/VHDraIcNCx/GBKNYs8MIIFMDCC
# BBigAwIBAgIQBAkYG1/Vu2Z1U0O1b5VQCDANBgkqhkiG9w0BAQsFADBlMQswCQYD
# VQQGEwJVUzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGln
# aWNlcnQuY29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0Ew
# HhcNMTMxMDIyMTIwMDAwWhcNMjgxMDIyMTIwMDAwWjByMQswCQYDVQQGEwJVUzEV
# MBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29t
# MTEwLwYDVQQDEyhEaWdpQ2VydCBTSEEyIEFzc3VyZWQgSUQgQ29kZSBTaWduaW5n
# IENBMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA+NOzHH8OEa9ndwfT
# CzFJGc/Q+0WZsTrbRPV/5aid2zLXcep2nQUut4/6kkPApfmJ1DcZ17aq8JyGpdgl
# rA55KDp+6dFn08b7KSfH03sjlOSRI5aQd4L5oYQjZhJUM1B0sSgmuyRpwsJS8hRn
# iolF1C2ho+mILCCVrhxKhwjfDPXiTWAYvqrEsq5wMWYzcT6scKKrzn/pfMuSoeU7
# MRzP6vIK5Fe7SrXpdOYr/mzLfnQ5Ng2Q7+S1TqSp6moKq4TzrGdOtcT3jNEgJSPr
# CGQ+UpbB8g8S9MWOD8Gi6CxR93O8vYWxYoNzQYIH5DiLanMg0A9kczyen6Yzqf0Z
# 3yWT0QIDAQABo4IBzTCCAckwEgYDVR0TAQH/BAgwBgEB/wIBADAOBgNVHQ8BAf8E
# BAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwMweQYIKwYBBQUHAQEEbTBrMCQGCCsG
# AQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQwYIKwYBBQUHMAKGN2h0
# dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RD
# QS5jcnQwgYEGA1UdHwR6MHgwOqA4oDaGNGh0dHA6Ly9jcmw0LmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwOqA4oDaGNGh0dHA6Ly9jcmwz
# LmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcmwwTwYDVR0g
# BEgwRjA4BgpghkgBhv1sAAIEMCowKAYIKwYBBQUHAgEWHGh0dHBzOi8vd3d3LmRp
# Z2ljZXJ0LmNvbS9DUFMwCgYIYIZIAYb9bAMwHQYDVR0OBBYEFFrEuXsqCqOl6nED
# wGD5LfZldQ5YMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3zbcgPMA0GCSqG
# SIb3DQEBCwUAA4IBAQA+7A1aJLPzItEVyCx8JSl2qB1dHC06GsTvMGHXfgtg/cM9
# D8Svi/3vKt8gVTew4fbRknUPUbRupY5a4l4kgU4QpO4/cY5jDhNLrddfRHnzNhQG
# ivecRk5c/5CxGwcOkRX7uq+1UcKNJK4kxscnKqEpKBo6cSgCPC6Ro8AlEeKcFEeh
# emhor5unXCBc2XGxDI+7qPjFEmifz0DLQESlE/DmZAwlCEIysjaKJAL+L3J+HNdJ
# RZboWR3p+nRka7LrZkPas7CM1ekN3fYBIM6ZMWM9CBoYs4GbT8aTEAb8B4H6i9r5
# gkn3Ym6hU/oSlBiFLpKR6mhsRDKyZqHnGKSaZFHvMIIGajCCBVKgAwIBAgIQAwGa
# Ajr/WLFr1tXq5hfwZjANBgkqhkiG9w0BAQUFADBiMQswCQYDVQQGEwJVUzEVMBMG
# A1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEw
# HwYDVQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTEwHhcNMTQxMDIyMDAwMDAw
# WhcNMjQxMDIyMDAwMDAwWjBHMQswCQYDVQQGEwJVUzERMA8GA1UEChMIRGlnaUNl
# cnQxJTAjBgNVBAMTHERpZ2lDZXJ0IFRpbWVzdGFtcCBSZXNwb25kZXIwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCjZF38fLPggjXg4PbGKuZJdTvMbuBT
# qZ8fZFnmfGt/a4ydVfiS457VWmNbAklQ2YPOb2bu3cuF6V+l+dSHdIhEOxnJ5fWR
# n8YUOawk6qhLLJGJzF4o9GS2ULf1ErNzlgpno75hn67z/RJ4dQ6mWxT9RSOOhkRV
# fRiGBYxVh3lIRvfKDo2n3k5f4qi2LVkCYYhhchhoubh87ubnNC8xd4EwH7s2AY3v
# J+P3mvBMMWSN4+v6GYeofs/sjAw2W3rBerh4x8kGLkYQyI3oBGDbvHN0+k7Y/qpA
# 8bLOcEaD6dpAoVk62RUJV5lWMJPzyWHM0AjMa+xiQpGsAsDvpPCJEY93AgMBAAGj
# ggM1MIIDMTAOBgNVHQ8BAf8EBAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8E
# DDAKBggrBgEFBQcDCDCCAb8GA1UdIASCAbYwggGyMIIBoQYJYIZIAYb9bAcBMIIB
# kjAoBggrBgEFBQcCARYcaHR0cHM6Ly93d3cuZGlnaWNlcnQuY29tL0NQUzCCAWQG
# CCsGAQUFBwICMIIBVh6CAVIAQQBuAHkAIAB1AHMAZQAgAG8AZgAgAHQAaABpAHMA
# IABDAGUAcgB0AGkAZgBpAGMAYQB0AGUAIABjAG8AbgBzAHQAaQB0AHUAdABlAHMA
# IABhAGMAYwBlAHAAdABhAG4AYwBlACAAbwBmACAAdABoAGUAIABEAGkAZwBpAEMA
# ZQByAHQAIABDAFAALwBDAFAAUwAgAGEAbgBkACAAdABoAGUAIABSAGUAbAB5AGkA
# bgBnACAAUABhAHIAdAB5ACAAQQBnAHIAZQBlAG0AZQBuAHQAIAB3AGgAaQBjAGgA
# IABsAGkAbQBpAHQAIABsAGkAYQBiAGkAbABpAHQAeQAgAGEAbgBkACAAYQByAGUA
# IABpAG4AYwBvAHIAcABvAHIAYQB0AGUAZAAgAGgAZQByAGUAaQBuACAAYgB5ACAA
# cgBlAGYAZQByAGUAbgBjAGUALjALBglghkgBhv1sAxUwHwYDVR0jBBgwFoAUFQAS
# KxOYspkH7R7for5XDStnAs0wHQYDVR0OBBYEFGFaTSS2STKdSip5GoNL9B6Jwcp9
# MH0GA1UdHwR2MHQwOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydEFzc3VyZWRJRENBLTEuY3JsMDigNqA0hjJodHRwOi8vY3JsNC5kaWdpY2Vy
# dC5jb20vRGlnaUNlcnRBc3N1cmVkSURDQS0xLmNybDB3BggrBgEFBQcBAQRrMGkw
# JAYIKwYBBQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcw
# AoY1aHR0cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElE
# Q0EtMS5jcnQwDQYJKoZIhvcNAQEFBQADggEBAJ0lfhszTbImgVybhs4jIA+Ah+WI
# //+x1GosMe06FxlxF82pG7xaFjkAneNshORaQPveBgGMN/qbsZ0kfv4gpFetW7ea
# sGAm6mlXIV00Lx9xsIOUGQVrNZAQoHuXx/Y/5+IRQaa9YtnwJz04HShvOlIJ8Oxw
# YtNiS7Dgc6aSwNOOMdgv420XEwbu5AO2FKvzj0OncZ0h3RTKFV2SQdr5D4HRmXQN
# JsQOfxu19aDxxncGKBXp2JPlVRbwuwqrHNtcSCdmyKOLChzlldquxC5ZoGHd2vNt
# omHpigtt7BIYvfdVVEADkitrwlHCCkivsNRu4PQUCjob4489yq9qjXvc2EQwggbN
# MIIFtaADAgECAhAG/fkDlgOt6gAK6z8nu7obMA0GCSqGSIb3DQEBBQUAMGUxCzAJ
# BgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5k
# aWdpY2VydC5jb20xJDAiBgNVBAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBD
# QTAeFw0wNjExMTAwMDAwMDBaFw0yMTExMTAwMDAwMDBaMGIxCzAJBgNVBAYTAlVT
# MRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5j
# b20xITAfBgNVBAMTGERpZ2lDZXJ0IEFzc3VyZWQgSUQgQ0EtMTCCASIwDQYJKoZI
# hvcNAQEBBQADggEPADCCAQoCggEBAOiCLZn5ysJClaWAc0Bw0p5WVFypxNJBBo/J
# M/xNRZFcgZ/tLJz4FlnfnrUkFcKYubR3SdyJxArar8tea+2tsHEx6886QAxGTZPs
# i3o2CAOrDDT+GEmC/sfHMUiAfB6iD5IOUMnGh+s2P9gww/+m9/uizW9zI/6sVgWQ
# 8DIhFonGcIj5BZd9o8dD3QLoOz3tsUGj7T++25VIxO4es/K8DCuZ0MZdEkKB4YNu
# gnM/JksUkK5ZZgrEjb7SzgaurYRvSISbT0C58Uzyr5j79s5AXVz2qPEvr+yJIvJr
# GGWxwXOt1/HYzx4KdFxCuGh+t9V3CidWfA9ipD8yFGCV/QcEogkCAwEAAaOCA3ow
# ggN2MA4GA1UdDwEB/wQEAwIBhjA7BgNVHSUENDAyBggrBgEFBQcDAQYIKwYBBQUH
# AwIGCCsGAQUFBwMDBggrBgEFBQcDBAYIKwYBBQUHAwgwggHSBgNVHSAEggHJMIIB
# xTCCAbQGCmCGSAGG/WwAAQQwggGkMDoGCCsGAQUFBwIBFi5odHRwOi8vd3d3LmRp
# Z2ljZXJ0LmNvbS9zc2wtY3BzLXJlcG9zaXRvcnkuaHRtMIIBZAYIKwYBBQUHAgIw
# ggFWHoIBUgBBAG4AeQAgAHUAcwBlACAAbwBmACAAdABoAGkAcwAgAEMAZQByAHQA
# aQBmAGkAYwBhAHQAZQAgAGMAbwBuAHMAdABpAHQAdQB0AGUAcwAgAGEAYwBjAGUA
# cAB0AGEAbgBjAGUAIABvAGYAIAB0AGgAZQAgAEQAaQBnAGkAQwBlAHIAdAAgAEMA
# UAAvAEMAUABTACAAYQBuAGQAIAB0AGgAZQAgAFIAZQBsAHkAaQBuAGcAIABQAGEA
# cgB0AHkAIABBAGcAcgBlAGUAbQBlAG4AdAAgAHcAaABpAGMAaAAgAGwAaQBtAGkA
# dAAgAGwAaQBhAGIAaQBsAGkAdAB5ACAAYQBuAGQAIABhAHIAZQAgAGkAbgBjAG8A
# cgBwAG8AcgBhAHQAZQBkACAAaABlAHIAZQBpAG4AIABiAHkAIAByAGUAZgBlAHIA
# ZQBuAGMAZQAuMAsGCWCGSAGG/WwDFTASBgNVHRMBAf8ECDAGAQH/AgEAMHkGCCsG
# AQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDovL29jc3AuZGlnaWNlcnQuY29t
# MEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3J0MIGBBgNVHR8EejB4MDqgOKA2hjRodHRwOi8v
# Y3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3JsMDqg
# OKA2hjRodHRwOi8vY3JsNC5kaWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURS
# b290Q0EuY3JsMB0GA1UdDgQWBBQVABIrE5iymQftHt+ivlcNK2cCzTAfBgNVHSME
# GDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzANBgkqhkiG9w0BAQUFAAOCAQEARlA+
# ybcoJKc4HbZbKa9Sz1LpMUerVlx71Q0LQbPv7HUfdDjyslxhopyVw1Dkgrkj0bo6
# hnKtOHisdV0XFzRyR4WUVtHruzaEd8wkpfMEGVWp5+Pnq2LN+4stkMLA0rWUvV5P
# sQXSDj0aqRRbpoYxYqioM+SbOafE9c4deHaUJXPkKqvPnHZL7V/CSxbkS3BMAIke
# /MV5vEwSV/5f4R68Al2o/vsHOE8Nxl2RuQ9nRc3Wg+3nkg2NsWmMT/tZ4CMP0qqu
# AHzunEIOz5HXJ7cW7g/DvXwKoO4sCFWFIrjrGBpN/CohrUkxg0eVd3HcsRtLSxwQ
# nHcUwZ1PL1qVCCkQJjGCBDswggQ3AgEBMIGGMHIxCzAJBgNVBAYTAlVTMRUwEwYD
# VQQKEwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xMTAv
# BgNVBAMTKERpZ2lDZXJ0IFNIQTIgQXNzdXJlZCBJRCBDb2RlIFNpZ25pbmcgQ0EC
# EANRgwbrGgA18btJ2k5C8GEwCQYFKw4DAhoFAKB4MBgGCisGAQQBgjcCAQwxCjAI
# oAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIB
# CzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFC4KTPW3JPt+/tgIuVBK
# bNSCR3/NMA0GCSqGSIb3DQEBAQUABIIBAAa9zDmvLbcwmiqyro61gj0PYkis9A+q
# pshjh2PDrUZ0JeO5n15xRJQ1NerWO5x0E28w6vF+Kn2x2qxXsTTxG6bhbH2FWQu/
# rxuJD4+rlG7H149q8ZtLpTx/Axu3tUFST0B7ua69+mJBTNmySql6fMfAgvJAV9yh
# hQJHS5i7hOvVe5RCL5Z5tw02qCeot24fyjyVnP0a9H/8qBgHSxYbK5me/+PBFp1l
# QG6Sqyjj+T/+lBri9vfDf1WvzYHkw7z3jhoUyymAPGmaSmq6ocgKfjXfzr0+MCdg
# /2+Jo7wofoCO3vUFdycnNI4yciK30PdY4gaTyFrloy6s/u72Ec+XMXuhggIPMIIC
# CwYJKoZIhvcNAQkGMYIB/DCCAfgCAQEwdjBiMQswCQYDVQQGEwJVUzEVMBMGA1UE
# ChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYD
# VQQDExhEaWdpQ2VydCBBc3N1cmVkIElEIENBLTECEAMBmgI6/1ixa9bV6uYX8GYw
# CQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcN
# AQkFMQ8XDTIwMDUwNTExMzc0NFowIwYJKoZIhvcNAQkEMRYEFMdZZnVh4nY8rKvL
# pKq8FojHt9l7MA0GCSqGSIb3DQEBAQUABIIBAHrlpHRWX2fdNyQs/HC/c0UsT8kw
# jrJZOrFNyvq43mC+pQcdr1M9Ie6uH1dFg6eJOAtFLlGIdRKFw1XSogQVp1BfqTKk
# OlCYP0IbOmTXYwY506uvu/iFBLs6EomnwmLUvFVpKNtdoaGgVogpduoFS7BJU0Iq
# rQMH9jXwwi9rDp3UWbJ658TEa6Wc0BrbMSvzK6FKrDA2A+5v0RBCosti3P23YFsA
# NNjYrfN/Pf0EMIi/dLAhOcjarzox998fY3kyXriCKwL4WgPMfFnnHRzT5pbSw8da
# C5UMBhVa9pMp7j1V3FdiOStPMiVlcgsHKIttd9SpRHWAMpUNSbt45tHXGps=
# SIG # End signature block
