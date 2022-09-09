# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# THIS CODE AND ANY ASSOCIATED INFORMATION ARE PROVIDED “AS IS” WITHOUT
# WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT
# LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS
# FOR A PARTICULAR PURPOSE. THE ENTIRE RISK OF USE, INABILITY TO USE, OR 
# RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# AUTHOR: Mihai Filip
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# Dependencies: MicrosoftTeams, MgGraph
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. 
# FOR DEBUGGING N PROD
$ErrorActionPreference = "SilentlyContinue"

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. #
#----Functions
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. #
function Get-TenantIdFromDomain {
  param (
    [string]$Domain
  )

  $url = "https://login.windows.net/$($Domain)/.well-known/openid-configuration"

  $tenantId = (Invoke-WebRequest $url | ConvertFrom-Json).token_endpoint.split('/')[3]

  return $tenantId
}

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. #
#----UI
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. #

Add-Type -AssemblyName PresentationCore, PresentationFramework

$Xaml = @"
<Window xmlns:syncfusion="http://schemas.syncfusion.com/wpf"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" Name="MainWindow" Title="TMS-NINJA v1.0" Background="#141414" WindowStartupLocation="
CenterScreen" ResizeMode="CanResizeWithGrip" FontFamily="Segoe UI" Height="844" Width="1262" BorderThickness="1" BorderBrush="#292929" AllowsTransparency="True" WindowStyle="None">

    <Grid>
        <DockPanel VerticalAlignment="Top" Height="580">

            <DockPanel Name="TitleBar" DockPanel.Dock="Top" Background="#141414" Height="24">

                <TextBlock Name="Title" Margin="7,0,0,0" HorizontalAlignment="Left" VerticalAlignment="Center" Background="#141414" Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold">
                    TMS-NINJA v1.0
                </TextBlock>

                <Button Name="btnCloseScreen" HorizontalAlignment="Right" VerticalAlignment="Center" BorderThickness="0" Height="20" Width="208" Cursor="Hand">

                    <Button.Template>
                        <ControlTemplate>
                            <Image Name="imgClose" Margin="0,0,5,0" Source="https://i.postimg.cc/G2xkjsq0/close.png" Height="14" Width="14" HorizontalAlignment="Right" VerticalAlignment="Center" ToolTip="Quit"/>
                        </ControlTemplate>
                    </Button.Template>
                </Button>

            </DockPanel>

            <StackPanel Name="stackAuth">
                <Grid Name="grdLogo">
                    <Border BorderThickness="0" BorderBrush="#292929">
                        <StackPanel>
                            <Image Name="Logo" Source="https://i.postimg.cc/sDQBJNjP/tms-ninja-logo-white.png" HorizontalAlignment="Center" VerticalAlignment="Top" Width="200" Height="200"></Image>
                        </StackPanel>
                    </Border>
                </Grid>

                <Grid Name="grdAuthFields">
                    <Border BorderThickness="0" BorderBrush="#292929">
                        <StackPanel>
                            <Label Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold" Margin="10,0,10,0">Username:</Label>
                            <TextBox Name="txtUsername" Padding="5" Margin="15,5,15,5"></TextBox>
                            <Label Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold" Margin="10,0,10,0">Password:</Label>
                            <PasswordBox Name="txtPassword" Padding="5" Margin="15,5,15,5"></PasswordBox>
                            <Label Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold" Margin="10,0,10,0">Client ID:</Label>
                            <PasswordBox Name="txtClientId" Padding="5" Margin="15,5,15,5"></PasswordBox>
                            <Label Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold" Margin="10,0,10,0">Client secret:</Label>
                            <PasswordBox Name="txtClientSecret" Padding="5" Margin="15,5,15,5"></PasswordBox>
                            <Button Name="btnSignIn" Content="Connect" Padding="5" Margin="5, 20, 5, 10" Width="150" Cursor="Hand"></Button>
                        </StackPanel>
                    </Border>
                </Grid>

                <Grid Name="grdAuthStatus">
                    <Label Name="lblAuthStatus" Content="" HorizontalAlignment="Center" Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold" Background="#141414"></Label>
                </Grid>

                <Grid Name="grdApp" Visibility="Hidden">
                    <Border BorderThickness="0" BorderBrush="#292929">
                        <DockPanel>
                            <Border BorderThickness="0" BorderBrush="#bdbdbd">
                                <Grid Name="grdNav" HorizontalAlignment="Left" Width="180">
                                    <StackPanel>
                                        <Border BorderThickness="0" BorderBrush="#bdbdbd">
                                            <Label Name="lblDirectRouting" Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold" Background="#141414" Padding="7">------ DIRECT ROUTING ------</Label>
                                        </Border>
                                        <Border BorderThickness="0" BorderBrush="#bdbdbd">
                                            <Label Name="lblDREnableUser" Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold" Cursor="Hand" Background="#141414" Padding="7">Enable user</Label>
                                        </Border>
                                        <Border BorderThickness="0" BorderBrush="#bdbdbd">
                                            <Label Name="lblDRConfigureVR" Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold" Cursor="Hand" Background="#141414" Padding="7">Configure voice routing</Label>
                                        </Border>
                                        <Border BorderThickness="0" BorderBrush="#bdbdbd">
                                            <Label Name="lblDRDisableUser" Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold" Cursor="Hand" Background="#141414" Padding="7">Disable user</Label>
                                        </Border>
                                        <Border BorderThickness="0" BorderBrush="#bdbdbd">
                                            <Label Name="lblDREnd" Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold" Background="#141414" Padding="7">---------------------------------</Label>
                                        </Border>
                                    </StackPanel>
                                </Grid>
                            </Border>

                            <Border BorderThickness="0" BorderBrush="#bdbdbd">
                                <Grid Name="grdActivity" HorizontalAlignment="Right" Background="#292929" Width="620">
                                        <Grid Name="grdDREnableUser" Margin="10,10,10,10">
                                        <StackPanel>
                                            <Border BorderThickness="0.5" BorderBrush="#bdbdbd" Padding="5" Margin="5">
                                                <Label Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold" FontSize="14">Direct Routing \ Enable user</Label>
                                            </Border>
                                            <Label Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold" Margin="0,5,0,0">User Principal Name:</Label>
                                            <TextBox Name="txtDREnableUPN" Padding="5" Margin="5" HorizontalContentAlignment="Left"></TextBox>
                                            <Label Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold">Phone Number:</Label>
                                            <TextBox Name="txtDREnablePhoneNumber" Padding="5" Margin="5" HorizontalContentAlignment="Left"></TextBox>
                                            <Button Name="btnDREnable" Content="Enable" FontFamily="Segoe UI" FontWeight="DemiBold" Margin="5,10,450,10" Height="30" Cursor="Hand"></Button>
                                            <TextBox Name="txtDREnableOutput" Margin="5,10,5,0" Height="265" IsEnabled="False" Background="#292929" FontFamily="Segoe UI" Foreground="#bdbdbd"></TextBox>
                                        </StackPanel>
                                    </Grid>
                                    <Grid Name="grdDRConfigureVR" Margin="10,10,10,10" Visibility="Hidden">
                                        <StackPanel>
                                            <Border BorderThickness="0.5" BorderBrush="#bdbdbd" Padding="5" Margin="5">
                                                <Label Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold" FontSize="14">Direct Routing \ Configure voice routing</Label>
                                            </Border>
                                            <Label Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold" Margin="0,5,0,0">User Principal Name:</Label>
                                            <TextBox Name="txtDRConfigureVRUPN" Padding="5" Margin="5" HorizontalContentAlignment="Left"></TextBox>
                                            <Label Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold">Voice routing policy:</Label>
                                            <ListBox Name="lstDRConfigureVRPPolicy" Padding="5" Margin="5" Height="100" HorizontalContentAlignment="Left" BorderThickness="1"/>
                                            <Button Name="btnDRConfigureVRAssign" Content="Assign" FontFamily="Segoe UI" FontWeight="DemiBold" Margin="5,10,450,10" Height="30" Cursor="Hand"></Button>
                                            <TextBox Name="txtDRConfigureVROutput" Margin="5,10,5,0" Height="193" IsEnabled="False" Background="#292929" FontFamily="Segoe UI" Foreground="#bdbdbd"></TextBox>
                                        </StackPanel>
                                    </Grid>
                                    <Grid Name="grdDRDisableUser" Margin="10,10,10,10" Visibility="Hidden">
                                        <StackPanel>
                                            <Border BorderThickness="0.5" BorderBrush="#bdbdbd" Padding="5" Margin="5">
                                                <Label Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold" FontSize="14">Direct Routing \ Disable user</Label>
                                            </Border>
                                            <Label Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold" Margin="0,5,0,0">User Principal Name:</Label>
                                            <TextBox Name="txtDRDisableUPN" Padding="5" Margin="5" HorizontalContentAlignment="Left"></TextBox>
                                            <CheckBox Name="chkDRRemoveVRP" Content="Remove voice routing policy" Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold" Margin="5"></CheckBox>
                                            <Button Name="btnDRDisable" Content="Disable" FontFamily="Segoe UI" FontWeight="DemiBold" Margin="5,10,450,10" Height="30" Cursor="Hand"></Button>
                                            <TextBox Name="txtDRDisableOutput" Margin="5,10,5,0" Height="304" IsEnabled="False" Background="#292929" FontFamily="Segoe UI" Foreground="#bdbdbd"></TextBox>
                                        </StackPanel>
                                    </Grid>
                                </Grid>
                            </Border>
                        </DockPanel>
                    </Border>

                    <DockPanel Margin="0,531,0,0" Background="#141414" Height="24">

                        <TextBlock Name="txtSignedInUser" Margin="7,0,0,0" HorizontalAlignment="Center" VerticalAlignment="Center" Background="#141414" Foreground="#bdbdbd" FontFamily="Segoe UI" FontWeight="DemiBold" FontSize="11">
                            Signed in as: user@domain.com
                        </TextBlock>

                    </DockPanel>
                </Grid>
            </StackPanel>

        </DockPanel>
    </Grid>
</Window>
"@

# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. #
#----Handlers and Logic
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. #

$Window = [Windows.Markup.XamlReader]::Parse($Xaml)

[xml]$xml = $Xaml

$xml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name $_.Name -Value $Window.FindName($_.Name) }

$Window.Width = 400
$Window.Height = 580
$Window.ResizeMode = 'NoResize'

$lblDREnableUser.Background = '#292929'


$btnCloseScreen.Add_Click({
    $Window.Close()

    Start-Sleep -s 3

    Disconnect-MicrosoftTeams
    Disconnect-MgGraph
})

$TitleBar.Add_MouseDown({
    $Window.DragMove()
})

$lblDREnableUser.Add_MouseEnter({
    $lblDREnableUser.Background = '#292929'
})

$lblDREnableUser.Add_MouseLeave({
    if ($grdDREnableUser.Visibility -eq 'Visible') {
        $lblDREnableUser.Background = '#292929'
    } else {
        $lblDREnableUser.Background = '#141414'
    }
})

$lblDRDisableUser.Add_MouseEnter({
    $lblDRDisableUser.Background = '#292929'
})

$lblDRDisableUser.Add_MouseLeave({
    if ($grdDRDisableUser.Visibility -eq 'Visible') {
        $lblDRDisableUser.Background = '#292929'
    } else {
        $lblDRDisableUser.Background = '#141414'
    }
})

$lblDRDisableUser.Add_MouseLeftButtonUp({
    $grdDRConfigureVR.Visibility = 'Hidden'
    $grdDREnableUser.Visibility = 'Hidden'
    $grdDRDisableUser.Visibility = 'Visible'
    $lblDRConfigureVR.Background = '#141414'
    $lblDREnableUser.Background = '#141414'
})

$lblDREnableUser.Add_MouseLeftButtonUp({
    $grdDRConfigureVR.Visibility = 'Hidden'
    $grdDRDisableUser.Visibility = 'Hidden'
    $grdDREnableUser.Visibility = 'Visible'
    $lblDRConfigureVR.Background = '#141414'
    $lblDRDisableUser.Background = '#141414'
})

$lblDRConfigureVR.Add_MouseEnter({
    $lblDRConfigureVR.Background = '#292929'
})

$lblDRConfigureVR.Add_MouseLeave({
    if ($grdDRConfigureVR.Visibility -eq 'Visible') {
        $lblDRConfigureVR.Background = '#292929'
    } else {
        $lblDRConfigureVR.Background = '#141414'
    }
})

$lblDRConfigureVR.Add_MouseLeftButtonUp({
    $grdDREnableUser.Visibility = 'Hidden'
    $grdDRDisableUser.Visibility = 'Hidden'
    $grdDRConfigureVR.Visibility = 'Visible'
    $lblDREnableUser.Background = '#141414'
    $lblDRDisableUser.Background = '#141414'
})

$btnSignIn.Add_Click({
    $ClientID = $txtClientId.Password 
    $ClientSecret = $txtClientSecret.Password 
    $ClientSecret = [Net.WebUtility]::URLEncode($ClientSecret)
    $TenantID = Get-TenantIdFromDomain -Domain $txtUsername.Text.split('@')[1] 
    $Username = $txtUsername.Text
    $Password = $txtPassword.Password
    $Password = [Net.WebUtility]::URLEncode($Password)

    $URI = "https://login.microsoftonline.com/$TenantID/oauth2/v2.0/token"
    $Body = "client_id=$ClientID&client_secret=$ClientSecret&grant_type=password&username=$Username&password=$Password"

    $RequestParameters = @{
      URI = $URI
      Method = "POST"
      ContentType = "application/x-www-form-urlencoded"
    }
    $GraphToken = (Invoke-RestMethod @RequestParameters -Body "$Body&scope=https://graph.microsoft.com/.default").access_token
    $TeamsToken = (Invoke-RestMethod @RequestParameters -Body "$Body&scope=48ac35b8-9aa8-4d74-927d-1f4a14a0b239/.default").access_token

    Connect-MicrosoftTeams -AccessTokens @($GraphToken, $TeamsToken)
    Connect-MgGraph -AccessToken $GraphToken

    $usr = get-csonlineuser -identity $Username
    $mgUser = Get-MgUser -Filter "userPrincipalName eq '$($Username)'"

    if ($usr) {
        Write-Host 'Teams auth valid.'
        $lblAuthStatus.Foreground = "#BDF2D5"
        $lblAuthStatus.Content = "TEAMS AUTH SUCCESSFUL"

        if ($mgUser) {
            Write-Host 'Graph auth valid.'
            $lblAuthStatus.Foreground = "#BDF2D5"
            $lblAuthStatus.Content = "TEAMS AND GRAPH AUTH SUCCESSFUL"
            $grdLogo.Visibility = 'Hidden'
            $grdAuthFields.Visibility = 'Hidden'
            $grdAuthStatus.Visibility = 'Hidden'
            $Window.Width = 800
            $Window.Left -= 200
            $grdApp.VerticalAlignment = 'Top'
            $grdApp.Margin = '0,-540,0,0'
            $grdApp.Height = 554
            $grdApp.Visibility = 'Visible'
            $txtSignedInUser.Text = "Signed in as: $($Username)"

            $VRPs = (Get-CsOnlineVoiceRoutingPolicy)

            foreach ($VRP in $VRPs) {
                if ($VRP.Identity -ne 'Global') {
                    $lstDRConfigureVRPPolicy.Items.Add($VRP.Identity.split(':')[1])
                } 
            }
        } else {
            Write-Host 'Graph auth invalid.'+
            $lblAuthStatus.Foreground = "Salmon"
            $lblAuthStatus.Content = "GRAPH AUTH FAILED"
        }

    } else {
        Write-Host 'Teams auth failed.'
        $lblAuthStatus.Foreground = "Salmon"
        $lblAuthStatus.Content = "TEAMS AUTH FAILED"
    }
})

$btnDREnable.Add_Click({
    $txtDREnableOutput.Foreground = "White"
    $txtDREnableOutput.Text = ''

    $DRUser = (Get-CsOnlineUser -Identity $txtDREnableUPN.Text)
 
    if ($DRUser) {
        Set-CsPhoneNumberAssignment -Identity $txtDREnableUPN.Text -EnterpriseVoiceEnabled $true
        $isEVEnabled = (Get-CsOnlineUser -Identity $txtDREnableUPN.Text).EnterpriseVoiceEnabled

        if ($isEVEnabled) {
            $txtDREnableOutput.Text += "The user has been successfully enabled for Enterprise Voice."

            Set-CsPhoneNumberAssignment -Identity $txtDREnableUPN.Text -PhoneNumber $txtDREnablePhoneNumber.Text -PhoneNumberType "DirectRouting"

            $DRPhoneNumber = (Get-CsOnlineUser -Identity $txtDREnableUPN.Text).lineUri
            if ($DRPhoneNumber -eq "tel:$($txtDREnablePhoneNumber.Text)") {
                $txtDREnableOutput.Foreground = "#BDF2D5"
                $txtDREnableOutput.Text = "Successfully enabled $($txtDREnableUPN.Text) for Direct Routing with the phone number $($txtDREnablePhoneNumber.Text)"
            } else {
                $txtDREnableOutput.Foreground = "Salmon"
                $txtDREnableOutput.Text = "Something went wrong with assigning the phone number."
            }
        } else {
            $txtDREnableOutput.Foreground = "Salmon"
            $txtDREnableOutput.Text = "Something went wrong with enabling the user for Enterprise Voice. Please check licenses."
        }
    } else {
        $txtDREnableOutput.Foreground = "Salmon"
        $txtDREnableOutput.Text = "The user does not exist"
    }

})

$btnDRConfigureVRAssign.Add_Click({
    $txtDRConfigureVROutput.Foreground = "White"
    $txtDRConfigureVROutput.Text = ''

    $DRUser = (Get-CsOnlineUser -Identity $txtDRConfigureVRUPN.Text)

    if ($DRUser) {
        $SelectedVRP = $lstDRConfigureVRPPolicy.SelectedItem

        if ($SelectedVRP) {
            Grant-CsOnlineVoiceRoutingPolicy -Identity $txtDRConfigureVRUPN.Text -PolicyName $SelectedVRP

            $assignedVRP = (Get-CsOnlineUser -Identity $txtDRConfigureVRUPN.Text).OnlineVoiceRoutingPolicy.Name

            if ($assignedVRP -eq $SelectedVRP) {
                $txtDRConfigureVROutput.Foreground = "#BDF2D5"
                $txtDRConfigureVROutput.Text += "Successfully assigned the voice routing policy $($SelectedVRP) to $($txtDRConfigureVRUPN.Text)`n"
            } else {
                $txtDRConfigureVROutput.Foreground = "Salmon"
                $txtDRConfigureVROutput.Text = "Something went wrong with assigning the voice routing policy."
            }

        } else {
            $txtDRConfigureVROutput.Text = "No voice routing policy selected. Please select a policy and retry."
        }

    } else {
        $txtDRConfigureVROutput.Foreground = "Salmon"
        $txtDRConfigureVROutput.Text = "The user does not exist."
    }
})

$btnDRDisable.Add_Click({
    $txtDRDisableOutput.Foreground = "White"
    $txtDRDisableOutput.Text = ""

    $DRUser = (Get-CsOnlineUser -Identity $txtDRDisableUPN.Text)
    $RemoveVRP = $chkDRRemoveVRP.IsChecked

    write-host $RemoveVRP

    if ($DRUser) {
        if ($RemoveVRP) {
            Grant-CsOnlineVoiceRoutingPolicy -Identity $txtDRDisableUPN.Text -PolicyName $null
        }
        Set-CsPhoneNumberAssignment -Identity $txtDRDisableUPN.Text -EnterpriseVoiceEnabled $false

        Remove-CsPhoneNumberAssignment -Identity $txtDRDisableUPN.Text -RemoveAll

        $DRUserLineUri = (Get-CsOnlineUser -Identity $txtDRDisableUPN.Text).lineUri

        if (!$DRUserLineUri) {
            if ($RemoveVRP) {
                $txtDRDisableOutput.Foreground = "#BDF2D5"
                $txtDRDisableOutput.Text = "Successfully disabled enterprise voice, removed line uri and voice routing policy."
            } else {
                $txtDRDisableOutput.Foreground = "#BDF2D5"
                $txtDRDisableOutput.Text = "Successfully disabled enterprise voice and removed line uri."
            }

        } else {
            $txtDRDisableOutput.Foreground = "Salmon"
            $txtDRDisableOutput.Text = "Something went wrong with disabling the user."
        }
    } else {
        $txtDRDisableOutput.Foreground = "Salmon"
        $txtDRDisableOutput.Text = "The user does not exist."
    }
})

$Window.ShowDialog()
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. #
#----END
# ...- --- .. -.. - .... . ...- .. .-.. .-.. .- .. -. #