<#
.SYNOPSIS
    Run-AvigilonApplianceSetup: Script to automate video appliance configuration.

.DESCRIPTION
    Automates the configuration of Avigilon HD video appliances.
    WARNING: Has only been tested on Avigilon HDVAs.

.NOTES
    Version:        3.0
    Author:         Eddie Miller
    Organization:   Central Technologies
    Creation Date:  1/13/2016
    Purpose/Change: 3/7/2016 - Full GUI tool
    TODO:           Better error handling
#>

Function Install-Software {

    [CmdletBinding()]
    Param(
        #District parameter passed from Create-GUI
        [String]$District
    )
    
    Begin {
        Try {
            #Default CentraStage and Avigilon Control Center Server installers 
            $DefaultCentraStageInstaller = "\\drobo\Software\CentraStage Clients\AgentSetup_TEMP.exe"
            $DefaultACCSInstaller = Get-ChildItem "\\drobo\Software\Avigilon Software\~All Avigilon Software Versions\" | Sort-Object -Descending | Select-Object -First 1 | Get-ChildItem | Where-Object { $_.Name -like "*Server*" }
            $DefaultACCCInstaller = Get-ChildItem "\\drobo\Software\Avigilon Software\~All Avigilon Software Versions\" | Sort-Object -Descending | Select-Object -First 1 | Get-ChildItem | Where-Object { $_.Name -like "*Client*" }

            #Path to CentraStage clients and Avigilon software
            $CentraStagePath = "\\drobo\Software\CentraStage Clients\"
            $ACCSPath = "\\drobo\Software\Avigilon Software\"
        }
        Catch {
            Write-Error -Message "Error in accessing software.`n$_.Exception.Message"
        }
    }
    Process {
        Try {
            #Installing software...
            #If "Not Listed" was selected in listbox, run the default installers
            If($District -eq "Not Listed") {
                & ($DefaultCentraStageInstaller) | Out-Null
                & ($DefaultACCSInstaller) | Out-Null
                & ($DefaultACCCInstaller) | Out-Null
            }
            #Else, run the selected installers
            Else {
                $CentraStageInstaller = Get-ChildItem $CentraStagePath -File -Filter "*$District*".Split()[0]
                & ($CentraStageInstaller.FullName) | Out-Null

                #Grab all the Avigilon installers in that district's folder and run each one
                $ACCSInstaller = Get-ChildItem $ACCSPath | Where-Object { $_ -like "*$District*".Split()[0] } `
                | Get-ChildItem -File

                $ACCSInstaller| ForEach-Object { & ($_.FullName) | Out-Null }
            }
        }
        Catch {
            Write-Error -Message "Error installing software.`n$_.Exception.Message"
        }
    }
    End { }

}

Function Remove-StartupItem {

    Begin {
        #Full path of program (link) to remove from startup
        If(Test-Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Avigilon Control Center 5 Client.lnk") {
            $StartupItem = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Avigilon Control Center 5 Client.lnk"
            Remove-Item $StartupItem
        }
        Else {
            #Do nothing - ACC Client not in startup
        }
    }
    Process { }
    End { }

}

Function Disable-Firewall {

    #Disabling the Windows Firewall...
    netsh advfirewall set allprofiles state off

}

Function Set-NetworkAdapterName {

    Begin {
        Try {
            #Identify network adapters
            $NICifnameCameras = Get-WmiObject -Class Win32_NetworkAdapter | Where-Object { $_.Manufacturer -like "Intel*" }
            $NICifnameWAN = Get-WmiObject -Class Win32_NetworkAdapter | Where-Object { $_.Manufacturer -like "Real*" }
        }
        Catch {
            Write-Error -Message "Error retrieving network adapters.`n$_.Exception.Message"
        }
    }
    Process {
        #Renaming network adapters...
        #THIS HAS ONLY BEEN TESTED & WORKS ON THE AVIGILON HDVAs - only 1 Intel adapter & 1 Realtek adapter
        $NICifnameCameras.NetConnectionID = "Cameras"
        $NICifnameWAN.NetConnectionID = "WAN"
        $NICifnameCameras.Put() | Out-Null
        $NICifnameWAN.Put() | Out-Null
    }
    End { }

}

Function Set-NetworkIPConfig {

    [CmdletBinding()]
    Param(
        #Reusing Networking variable from previous version
        #Pass values to this function as an array
        [array]$Networking
    )

    Begin {
        #Obtain the index for the Camera NIC
        $Index = (Get-WmiObject -Class Win32_NetworkAdapter | Where-Object { $_.NetConnectionID -eq "Cameras" }).InterfaceIndex
        $NetInterface = Get-WmiObject -Class Win32_NetworkAdapterConfiguration | Where-Object { $_.InterfaceIndex -eq $Index }
    }
    Process {
        Try {
            #Configuring networking...
            #Networking[0] = IP address
            #Networking[1] = Subnet mask
            #Networking[2] = Gateway
            #Networking[3] = DNS1
            #Networking[4] = DNS2
            $NetInterface.EnableStatic($Networking[0], $Networking[1]) | Out-Null
            $NetInterface.SetGateways($Networking[2]) | Out-Null

            If (!$Networking[4]) { $NetInterface.SetDNSServerSearchOrder(@($Networking[3])) | Out-Null }
            Else { $NetInterface.SetDNSServerSearchOrder(@($Networking[3], $Networking[4])) | Out-Null }

            $NetInterface.SetDynamicDNSRegistration("TRUE")   
        }
        Catch {
            Write-Error -Message "Error setting IP configuration`n$_.Exception.Message"
        }
    }
    End { }

}

Function Run-Updates {

    Begin {
        #New COM object for Windows Update settings
        $WUSettings = (New-Object -ComObject "Microsoft.Update.AutoUpdate").Settings

        #Set NotificationLevel to Scheduled Install, check, & install Windows Updates
        #1 = Disabled
        #2 = Notify before download
        #3 = Notify before installation
        #4 = Scheduled installation
    }
    Process {
        Try {
            #Disabling automatic Windows Updates
            $WUSettings.NotificationLevel = 1
            $WUSettings.Save()
        }
        Catch {
            Write-Error -Message "Error changing Windows Update settings`n$_.Exception.Message"
        }
    }
    End {
        #Running Windows Updates...
        wuauclt /ShowWUAutoScan
    }
}

Function Remove-ScheduledTasks {
    
    #For now, just removed the Disk Defrag scheduled task - may be causing performance issues
    schtasks /d /tn \Microsoft\Windows\Defrag\ScheduledDefrag /f

}

Function Create-GUI {

    #XAML for GUI creation and declaration of variables
    Begin {
        #Generated from Visual Studio 2015
        $inputXML = @"
<Window x:Class="AvigilonApplianceConfig.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:AvigilonApplianceConfig"
        mc:Ignorable="d"
        Title="Avigilon Appliance Configuration" Height="426" Width="525">
    <Grid>
        <Image x:Name="Central_Technologies_Logo" HorizontalAlignment="Left" Height="100" VerticalAlignment="Top" Width="100" Source="\\drobo\Public\Scripts\Central Logo.jpg"/>
        <TextBlock x:Name="TextDescription" HorizontalAlignment="Left" Height="35" Margin="100,36,0,0" TextWrapping="Wrap" Text="This tool will install the necessary components for deploying an Avigilon HD Video Appliance." VerticalAlignment="Top" Width="374"/>
        <ListBox x:Name="DistrictListBox" HorizontalAlignment="Left" Height="160" Margin="23,125,0,0" VerticalAlignment="Top" Width="126"/>
        <Button x:Name="ButtonBegin" Content="Begin" HorizontalAlignment="Left" Margin="412,350,0,0" VerticalAlignment="Top" Width="75"/>
        <TextBlock x:Name="TextSelectDistrict" HorizontalAlignment="Left" Height="17" Margin="23,103,0,0" TextWrapping="Wrap" Text="Select a school district:" VerticalAlignment="Top" Width="126"/>
        <TextBlock x:Name="TextNetworkDescription" HorizontalAlignment="Left" Height="17" Margin="195,103,0,0" TextWrapping="Wrap" Text="Enter networking information for the Cameras adapter:" VerticalAlignment="Top" Width="292"/>
        <TextBlock x:Name="TextIPAddress" HorizontalAlignment="Left" Margin="214,127,0,0" TextWrapping="Wrap" Text="IP Address:" VerticalAlignment="Top"/>
        <TextBlock x:Name="TextSubnetMask" HorizontalAlignment="Left" Margin="214,153,0,0" TextWrapping="Wrap" Text="Subnet Mask:" VerticalAlignment="Top"/>
        <TextBlock x:Name="TextGateway" HorizontalAlignment="Left" Margin="214,179,0,0" TextWrapping="Wrap" Text="(Optional) Gateway:" VerticalAlignment="Top"/>
        <TextBlock x:Name="TextDNS1" HorizontalAlignment="Left" Margin="214,205,0,0" TextWrapping="Wrap" Text="(Optional) DNS 1:" VerticalAlignment="Top"/>
        <TextBlock x:Name="TextDNS2" HorizontalAlignment="Left" Margin="214,231,0,0" TextWrapping="Wrap" Text="(Optional) DNS 2:" VerticalAlignment="Top"/>
        <TextBox x:Name="InputIPAddress" HorizontalAlignment="Left" Height="22" Margin="355,124,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="132" Text="0.0.0.0" TextAlignment="Right"/>
        <TextBox x:Name="InputSubnetMask" HorizontalAlignment="Left" Height="22" Margin="355,150,0,0" TextWrapping="Wrap" Text="0.0.0.0" VerticalAlignment="Top" Width="132" TextAlignment="Right"/>
        <TextBox x:Name="InputDNS1" HorizontalAlignment="Left" Height="22" Margin="355,200,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="132" TextAlignment="Right"/>
        <TextBox x:Name="InputGateway" HorizontalAlignment="Left" Height="22" Margin="355,175,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="132" TextAlignment="Right"/>
        <TextBox x:Name="InputDNS2" HorizontalAlignment="Left" Height="22" Margin="355,226,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="132" TextAlignment="Right"/>
    </Grid>
</Window>
"@       
 
        #Replacing characters in XML so it will run in PowerShell
        $inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
        
        #Loading the Presentation Framework to display forms
        [void][System.Reflection.Assembly]::LoadWithPartialName('PresentationFramework')
        [xml]$XAML = $inputXML
        
        #Read XAML
        $reader = (New-Object System.Xml.XmlNodeReader $xaml) 
        Try {
            $Form = [Windows.Markup.XamlReader]::Load($reader)
        }
        Catch {
            Write-Error "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .Net is installed."
        }

        #Loading message box assembly
        [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

        #Regex string for checking validity of IP address and subnet mask
        $Regex = "^(?:(?:0?0?\d|0?[1-9]\d|1\d\d|2[0-5][0-5]|2[0-4]\d)\.){3}(?:0?0?\d|0?[1-9]\d|1\d\d|2[0-5][0-5]|2[0-4]\d)$"

        #Generate list of school districts based on folders for Avigilon Software
        #Not the most elegant solution but it works
        $DistrictList = Get-ChildItem "\\drobo\Software\Avigilon Software" -Exclude "*~*" | ForEach-Object { $_.Name.Substring(11) }

        #Create variables for the items on form
        $xaml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) }
        
        #Populate listbox with school districts
        [void]$WPFDistrictListBox.Items.Add("Not Listed")
        $DistrictList | ForEach-Object { [void]$WPFDistrictListBox.Items.Add($_) }
    }
    #Displays GUI then calls other functions when Begin is clicked
    Process {
        #When Begin is clicked check that the form is completed and correct, then run functions
        $WPFButtonBegin.Add_Click({
            If(($WPFInputIPAddress.Text -notmatch $Regex) -or ($WPFInputSubnetMask.Text -notmatch $Regex) -or (!$WPFDistrictListBox.SelectedItem)) {
                #Error handling if IP/subnet is not valid or if no district was selected
                [System.Windows.Forms.MessageBox]::Show("Invalid input.`nSelect a district and then enter networking information in the format 'XXX.XXX.XXX.XXX'", "Error")
            }
            Else {
                $WPFButtonBegin.IsEnabled = $False
                [System.Windows.Forms.MessageBox]::Show("Running.")
                Install-Software -District $WPFDistrictListBox.SelectedItem
                Remove-StartupItem
                Disable-Firewall
                Set-NetworkAdapterName
                Set-NetworkIPConfig -Networking @($WPFInputIPAddress.Text, $WPFInputSubnetMask.Text, $WPFInputGateway.Text, $WPFInputDNS1.Text, $WPFInputDNS2.Text)
                Remove-ScheduledTasks
                Run-Updates
                [System.Windows.Forms.MessageBox]::Show("Done!")
            }
        })

        #Display the form
        $Form.ShowDialog() | Out-Null
    }
    #Cleanup
    End { }

}

Create-GUI
