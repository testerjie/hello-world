

$formChildForm_Load={
	#TODO: Initialize Form Controls here
	
}


Function TopologyTypeChoose
{
	
	If ($combobox1.SelectedIndex -eq 0)
	{
		
		Write-Host "We start use $($combobox1.SelectedItem)"  -ForegroundColor Cyan -WarningAction Continue
		
		
	}
	
	ElseIf ($combobox1.SelectedIndex -eq 1)
	{
		
		Write-Host "We start use" $($combobox1.SelectedItem)  -ForegroundColor Cyan -WarningAction Continue
		
	}
	
	ElseIf ($combobox1.SelectedIndex -eq 2)
	{
		Write-Host "We start use:" $($combobox1.SelectedItem)  -ForegroundColor Cyan -WarningAction Continue
		
	}
	ElseIf ($combobox1.SelectedIndex -eq 3)
	{
		Write-Host "We start use:" $($combobox1.SelectedItem) -ForegroundColor Cyan -WarningAction Continue
		
	}
}

#region Control Helper Functions
function Update-ComboBox
{
<#
	.SYNOPSIS
		This functions helps you load items into a ComboBox.
	
	.DESCRIPTION
		Use this function to dynamically load items into the ComboBox control.
	
	.PARAMETER ComboBox
		The ComboBox control you want to add items to.
	
	.PARAMETER Items
		The object or objects you wish to load into the ComboBox's Items collection.
	
	.PARAMETER DisplayMember
		Indicates the property to display for the items in this control.
		
	.PARAMETER ValueMember
		Indicates the property to use for the value of the control.
	
	.PARAMETER Append
		Adds the item(s) to the ComboBox without clearing the Items collection.
	
	.EXAMPLE
		Update-ComboBox $combobox1 "Red", "White", "Blue"
	
	.EXAMPLE
		Update-ComboBox $combobox1 "Red" -Append
		Update-ComboBox $combobox1 "White" -Append
		Update-ComboBox $combobox1 "Blue" -Append
	
	.EXAMPLE
		Update-ComboBox $combobox1 (Get-Process) "ProcessName"
	
	.NOTES
		Additional information about the function.
#>
	
	param
	(
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		[System.Windows.Forms.ComboBox]
		$ComboBox,
		[Parameter(Mandatory = $true)]
		[ValidateNotNull()]
		$Items,
		[Parameter(Mandatory = $false)]
		[string]$DisplayMember,
		[Parameter(Mandatory = $false)]
		[string]$ValueMember,
		[switch]
		$Append
	)
	
	if (-not $Append)
	{
		$ComboBox.Items.Clear()
	}
	
	if ($Items -is [Object[]])
	{
		$ComboBox.Items.AddRange($Items)
	}
	elseif ($Items -is [System.Collections.IEnumerable])
	{
		$ComboBox.BeginUpdate()
		foreach ($obj in $Items)
		{
			$ComboBox.Items.Add($obj)
		}
		$ComboBox.EndUpdate()
	}
	else
	{
		$ComboBox.Items.Add($Items)
	}
	
	$ComboBox.DisplayMember = $DisplayMember
	$ComboBox.ValueMember = $ValueMember
}
#endregion
Function TopologyContentChoose
{
	
	if ($combobox2.SelectedIndex -eq 0)
	{
		$DBMap = @{ }
		$xmldata = [xml](Get-Content -Path ".\topologyconfig.xml")
		switch ($combobox1.SelectedIndex)
		{
			0 {
				$computernames = $xmldata.Config.prem2013.ServerName;
				$Username = $xmldata.Config.prem2013.User
				$Password = $xmldata.Config.prem2013.Password
				$Usernameforedge = $xmldata.Config.prem2013.Useredge
				$Passwordforedge = $xmldata.Config.prem2013.Passwordedge
				break
			}
			1 {
				$computernames = $xmldata.Config.prem2016.ServerName;
				$Username = $xmldata.Config.prem2016.User
				$Password = $xmldata.Config.prem2016.Password
				$Usernameforedge = $xmldata.Config.prem2016.Useredge
				$Passwordforedge = $xmldata.Config.prem2016.Passwordedge
				break
			}
			2 {
				$computernames = $xmldata.Config.S4BSE.ServerName;
				$Username = $xmldata.Config.S4BSE.User
				$Password = $xmldata.Config.S4BSE.Password
				break
			}
			3 {
				$computernames = $xmldata.Config.S4BSEW171.ServerName;
				$Username = $xmldata.Config.S4BSEW171.User
				$Password = $xmldata.Config.S4BSEW171.Password
				break
			}
		}
		Write-Host $computernames
		
		$pass = ConvertTo-SecureString -AsPlainText $Password -Force
		$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $pass
		$ocscore = Invoke-Command -ComputerName $computernames[0] -ScriptBlock { Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | ForEach-Object { Get-ItemProperty $_.PsPath } | where { $_.DisplayName -and ($_.DisplayName -match ", Core Components*") } } -Credential $Cred
		
		#Get the expected dbversion map
		[System.String]$version = $ocscore.DisplayVersion
		Write-Host $version
		[System.String]$wave = $version.Split('.')[0]
		Write-Host $wave
		
		if ($wave -eq 5)
		{
			$DBData = [xml](Get-Content ".\w15DBVersion.xml")
			
			$DBs = $DBData.DbVersions.Db
			foreach ($DB in $DBs)
			{
				$DBMap.Add($DB.DbName, $DB.VersionSchema.Value + '.' + $DB.VersionSproc.Value + '.' + $DB.VersionUpgrade.Value)
			}
			$DBMap | Out-Default
		}
		elseif ($wave -eq 6)
		{
			$DBData = [xml](Get-Content -Path ".\w16DBVersion.xml")
			Write-Host $DBData
			$DBs = $DBData.DbVersions.Db
			foreach ($DB in $DBs)
			{
				$DBMap.Add($DB.DbName, $DB.VersionSchema.Value + '.' + $DB.VersionSproc.Value + '.' + $DB.VersionUpgrade.Value)
			}
			$DBMap | Out-Default
		}
		elseif ($wave -eq 7)
		{
			$DBData = [xml](Get-Content -Path ".\w17DBVersion.xml")
			Write-Host $DBData
			$DBs = $DBData.DbVersions.Db
			foreach ($DB in $DBs)
			{
				$DBMap.Add($DB.DbName, $DB.VersionSchema.Value + '.' + $DB.VersionSproc.Value + '.' + $DB.VersionUpgrade.Value)
			}
			$DBMap | Out-Default
		}
		
		
		
		foreach ($computername in $computernames)
		{
			if ($computername -contains "edge")
			{
				Write-Host -ForegroundColor Yellow -BackgroundColor Blue $computername is NA
				break
			}
			# disable firewall on domain
			Write-Host -ForegroundColor White -BackgroundColor Black "*******Now need to disable firewall...*******"
			Invoke-Command -ComputerName $computername -ScriptBlock { Set-NetFirewallProfile -Profile Domain -Enabled False} -Credential $Cred
			#start to connect db
			Write-Host "We start check:"$($combobox1.SelectedItem)"'s" $computername "'s"$($combobox2.SelectedItem) -ForegroundColor Cyan -WarningAction Continue
			
			
			$instances = @("xds", "rtcdyn", "rtc", "lyss")
			
			foreach ($instance in $instances)
			{
				$dbty = "RTCLOCAL"
				if ($instance -eq "lyss")
				{
					$dbty="LYNCLOCAL"
				}
				
				$ConnectionStrings = "Data Source=$($computername)\$($dbty);Initial Catalog=$instance;Integrated Security=SSPI"
				Write-Host $ConnectionStrings
				Write-Host -ForegroundColor Green $instance Installed Version is ..
				$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
				$SqlConnection.ConnectionString = $ConnectionStrings
				#$SqlConnection.Open()
				$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
				$SqlCmd.CommandText = "select * from $($instance).dbo.DbConfigInt"
				$SqlCmd.Connection = $SqlConnection
				$sqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
				$sqlAdapter.SelectCommand = $SqlCmd
				$set = New-Object System.Data.DataSet
				$sqlresult = $sqlAdapter.Fill($set) | Out-Null
				$? | Out-Default
				if ($? -gt 0)
				{
					Write-Host -ForegroundColor Red $instance Installed Version is null
				}
				
				elseif ($sqlresult -eq 0)
				{
					Write-Host -ForegroundColor Red $instance Installed Version is 0.0.0
					
				}
				
				$table = $set.Tables[0] | Out-Default
				$table = $set.Tables[0] | Format-Table -AutoSize -Property Name,Version
				foreach ($list in $table)
				{
					Write-Host $list
				}
			
				
				#[string]$list.Item("DbVersionSchema") + '.'+ [string]$list.Item("DbVersionSproc")+'.'+ [string]$list.Item("DbVersionUpgrade")
				
				
				
			}
			
		}
		
	}
	





	elseif ($combobox2.SelectedIndex -eq 1)
	{
	
	Write-Host "We start check:"$($combobox1.SelectedItem)"'s"$($combobox2.SelectedItem)
	
	$xmldata = [xml](Get-Content -Path ".\topologyconfig.xml")
	switch ($combobox1.SelectedIndex)
	{
		0 {
			$computernames = $xmldata.Config.prem2013.ServerName;
			$Username = $xmldata.Config.prem2013.User
			$Password = $xmldata.Config.prem2013.Password
			$Usernameforedge = $xmldata.Config.prem2013.Useredge
			$Passwordforedge = $xmldata.Config.prem2013.Passwordedge
			break
		}
		1 {
			$computernames = $xmldata.Config.prem2016.ServerName;
			$Username = $xmldata.Config.prem2016.User
			$Password = $xmldata.Config.prem2016.Password
			$Usernameforedge = $xmldata.Config.prem2016.Useredge
			$Passwordforedge = $xmldata.Config.prem2016.Passwordedge
			break
		}
		2 {
			$computernames = $xmldata.Config.S4BSE.ServerName;
			$Username = $xmldata.Config.S4BSE.User
			$Password = $xmldata.Config.S4BSE.Password
			break
		}
		3 {
			$computernames = $xmldata.Config.S4BSEW171.ServerName;
			$Username = $xmldata.Config.S4BSEW171.User
			$Password = $xmldata.Config.S4BSEW171.Password
			$Usernameforedge = $xmldata.Config.prem2016.Useredge
			$Passwordforedge = $xmldata.Config.prem2016.Passwordedge
			break
		}
	}
	
	######### For each server#########
	Write-Host -ForegroundColor Green ("--------------PatchVersion check result list :-------------- `n")
	foreach ($computername in $computernames)
	{
		if ($computername -contains "edge")
		{
			Enable-PSRemoting -Force
			Set-Item wsman:\localhost\client\trustedhosts * -Force
			Restart-Service WinRM
			$Username = $Usernameforedge;
			$Password = $Passwordforedge;
		}
		$pass = ConvertTo-SecureString -AsPlainText $Password -Force
		$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $pass
		#$Checkarrys = @("Skype*","Lync*", "Unified*", "Fabric*")
		$programVer = @() #initialize array for Program information
		Write-Host -ForegroundColor Green ("$computername version test result: ")
		
		
		$programVer += Invoke-Command -ComputerName $computername -ScriptBlock { Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | ForEach-Object { Get-ItemProperty $_.PsPath } | where { $_.DisplayName -and ($_.DisplayName -match "Skype*") } } -Credential $Cred
		$programVer += Invoke-Command -ComputerName $computername -ScriptBlock { Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | ForEach-Object { Get-ItemProperty $_.PsPath } | where { $_.DisplayName -and ($_.DisplayName -match "Lync*") } } -Credential $Cred
		$programVer += Invoke-Command -ComputerName $computername -ScriptBlock { Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | ForEach-Object { Get-ItemProperty $_.PsPath } | where { $_.DisplayName -and ($_.DisplayName -match "Unified*") } } -Credential $Cred
		$programVer += Invoke-Command -ComputerName $computername -ScriptBlock { Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | ForEach-Object { Get-ItemProperty $_.PsPath } | where { $_.DisplayName -and ($_.DisplayName -match "Fabric*") } } -Credential $Cred
		
		
		$programVer = $programVer | Sort-Object DisplayName -Unique
		
		# Need to check if it is an empty or null returned
		# or maybe count or length
		
		foreach ($prog in $programVer)
		{
			Write-Host -ForegroundColor Blue -BackgroundColor White  ("{0,-75} {1}" -f $($prog.DisplayName), $($prog.DisplayVersion)) | Out-Default
		}
	}
}

elseif ($combobox2.SelectedIndex -eq 2)
{
	Write-Host "We start check:"$($combobox1.SelectedItem)"'s"$($combobox2.SelectedItem) -ForegroundColor DarkCyan
	$flag = 0
	$CurrentyDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
	$xmldata = [xml](Get-Content -Path ".\topologyconfig.xml")
	switch ($combobox1.SelectedIndex)
	{
		0 {
			$computernames = $xmldata.Config.prem2013.ServerName;
			$Username = $xmldata.Config.prem2013.User
			$Password = $xmldata.Config.prem2013.Password
			$Usernameforedge = $xmldata.Config.prem2013.Useredge
			$Passwordforedge = $xmldata.Config.prem2013.Passwordedge
			break
		}
		1 {
			$computernames = $xmldata.Config.prem2016.ServerName;
			$Username = $xmldata.Config.prem2016.User
			$Password = $xmldata.Config.prem2016.Password
			$Usernameforedge = $xmldata.Config.prem2016.Useredge
			$Passwordforedge = $xmldata.Config.prem2016.Passwordedge
			break
		}
		2 {
			$computernames = $xmldata.Config.S4BSE.ServerName;
			$Username = $xmldata.Config.S4BSE.User
			$Password = $xmldata.Config.S4BSE.Password
			break
		}
		3 {
			$computernames = $xmldata.Config.S4BSEW171.ServerName;
			$Username = $xmldata.Config.S4BSEW171.User
			$Password = $xmldata.Config.S4BSEW171.Password
			break
		}
	}
	
	
	
	
	
	Write-Host -ForegroundColor Green ("--------------Services check result list :-------------- `n")
	foreach ($computername in $computernames)
	{
		
		#DONE: Filter server edge
		if ($computername -contains "edge")
		{
			$Username = $Usernameforedge;
			$Password = $Passwordforedge;
		}
		
		$pass = ConvertTo-SecureString -AsPlainText $Password -Force
		$Cred = New-Object System.Management.Automation.PSCredential -ArgumentList $Username, $pass
		Invoke-Command -ComputerName $computername -ScriptBlock { Get-CsWindowsService | Format-Table -AutoSize } -credential $Cred | Out-Default
		
		
		$ServiceStatuses = Invoke-Command -ComputerName $computername -ScriptBlock { Get-CsWindowsService } -credential $Cred
		
		foreach ($ServiceStatus in $ServiceStatuses)
		{
			if ($ServiceStatus.ToString -Contains ("stopped", "suspended"))
			{
				$flag = 1
			}
			
		}
		
		if ($flag -eq 0)
		{
			Write-Host -ForegroundColor Green ("$computername Service test pass ")
		}
		else
		{
			Write-Host -ForegroundColor Red ("$computername Service test failed ")
			$flag = 0
		}
		
	}
}

}


$buttonOK_Click={
	#TODO: Place custom script here
	
	TopologyTypeChoose
	TopologyContentChoose
}
