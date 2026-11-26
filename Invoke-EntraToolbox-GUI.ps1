#Requires -Version 5.1
<#
.SYNOPSIS
    Entra Toolbox: a WPF parameter and output interface for the bundled Entra scripts.
.DESCRIPTION
    Run on Windows in an STA session. Task implementations remain in the feature
    scripts. Microsoft Graph submodules are checked per task, never auto-installed.
    Uses delegated sign-in; required scopes follow the selected script.
.PARAMETER TenantId
    Optional tenant ID used for sign-in and checked before each operation.
.PARAMETER ScriptsRoot
    Root containing the feature scripts, normally this launcher's Features directory.
.PARAMETER ValidateOnly
    Build and lay out the interface without authentication or showing a window.
.EXAMPLE
    powershell.exe -STA -File .\Invoke-EntraToolbox-GUI.ps1
.EXAMPLE
    pwsh.exe -STA -File .\Invoke-EntraToolbox-GUI.ps1 -TenantId 'your-tenant-id'
.NOTES
        Deployment: copy the entire EntraToolbox folder, preserving these paths
        relative to the GUI directory:
            Features\Add-MembersToAADGroup\Add-MembersToEntraGroup.ps1
            Features\Create-AADGroups\Create-EntraSecurityGroupsFromCsv.ps1
            Features\Get-PrimaryPCFromEntraUsers\Get-PrimaryPCFromEntraUsers.ps1
            Features\Get-DevicesLastUser.ps1
            Features\Invoke-EntraDeviceCleanup.ps1
            Features\ExtensionAttributes\Set-ExtensionAttribute.ps1
            Shared\Invoke-EntraToolboxOperation.ps1
        Use the matching updated versions of these scripts, including their optional
        GUI integration parameters. Do not flatten the subfolders. -ScriptsRoot can
        point to another directory containing the same structure as Features.
        Shared is always resolved beside the GUI, independently of -ScriptsRoot.
        Tests\Test-EntraToolbox.ps1 is optional for offline validation, not required to run.
        See README.md for standalone feature commands and prerequisites.
        The Microsoft Graph modules below must also be installed on the destination.

    Version 1.0. Requires Microsoft.Graph.Authentication. Task-specific dependencies:
    Microsoft.Graph.Groups, Microsoft.Graph.Users,
    Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Reports.
    Install missing modules with Install-Module <module> -Scope CurrentUser.
#>
[CmdletBinding()]
param(
    [string]$TenantId,
    [string]$ScriptsRoot,
    [switch]$ValidateOnly
)

if (-not $PSBoundParameters.ContainsKey('ScriptsRoot')) {
    $ScriptsRoot = Join-Path $PSScriptRoot 'Features'
}

if ([Environment]::OSVersion.Platform -ne 'Win32NT') { throw 'Entra Toolbox requires Windows.' }
if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') { throw 'Start PowerShell with -STA to use Entra Toolbox.' }
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
$script:EntraRoot = $ScriptsRoot
$script:RequestedTenant = $TenantId
$script:ConnectedTenant = ''
# Excel needs a BOM to read UTF-8 (æ/ø/å); PowerShell 7 'UTF8' omits it.
$script:Utf8BomEncoding = if ($PSVersionTable.PSVersion.Major -ge 6) { 'utf8BOM' } else { 'UTF8' }
$script:IsDark = $true
$script:Job = $null
$script:Runspace = $null
$script:CloseRequested = $false

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Entra Toolbox" Width="1280" Height="900" MinWidth="1050" MinHeight="720"
        WindowStartupLocation="CenterScreen" Background="{DynamicResource BgBrush}"
        FontFamily="Segoe UI" FontSize="13" Foreground="{DynamicResource TextBrush}">
 <Window.Resources>
  <SolidColorBrush x:Key="BgBrush" Color="#191c22"/>
  <SolidColorBrush x:Key="SurfaceBrush" Color="#252b34"/>
  <SolidColorBrush x:Key="InputBgBrush" Color="#20252d"/>
  <SolidColorBrush x:Key="TextBrush" Color="#e6e9ef"/>
  <SolidColorBrush x:Key="SubTextBrush" Color="#9ca8b7"/>
  <SolidColorBrush x:Key="BorderBrush2" Color="#39424e"/>
  <SolidColorBrush x:Key="AccentBrush" Color="#dc3545"/>
  <SolidColorBrush x:Key="AccentHover" Color="#bd2130"/>
  <SolidColorBrush x:Key="SelectionBrush" Color="#49272e"/>
  <SolidColorBrush x:Key="SuccessBrush" Color="#36c18f"/>
  <SolidColorBrush x:Key="WarningBrush" Color="#fbbf24"/>
  <SolidColorBrush x:Key="ErrorBrush" Color="#ff8791"/>
  <Style TargetType="TextBlock"><Setter Property="Foreground" Value="{DynamicResource TextBrush}"/></Style>
  <Style TargetType="Button">
   <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
   <Setter Property="Background" Value="{DynamicResource SurfaceBrush}"/>
   <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush2}"/>
   <Setter Property="Padding" Value="12,7"/><Setter Property="Margin" Value="0,0,8,0"/>
   <Setter Property="Cursor" Value="Hand"/>
   <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button">
    <Border x:Name="ButtonBorder" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="4" Padding="{TemplateBinding Padding}">
     <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
    </Border>
    <ControlTemplate.Triggers>
     <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="ButtonBorder" Property="BorderBrush" Value="{DynamicResource AccentBrush}"/></Trigger>
     <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="ButtonBorder" Property="BorderBrush" Value="{DynamicResource AccentBrush}"/></Trigger>
     <Trigger Property="IsPressed" Value="True"><Setter TargetName="ButtonBorder" Property="Opacity" Value="0.7"/></Trigger>
     <Trigger Property="IsEnabled" Value="False"><Setter TargetName="ButtonBorder" Property="Opacity" Value="0.4"/></Trigger>
    </ControlTemplate.Triggers>
   </ControlTemplate></Setter.Value></Setter>
  </Style>
  <Style x:Key="PrimaryBtn" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
   <Setter Property="Background" Value="{DynamicResource AccentBrush}"/><Setter Property="BorderBrush" Value="{DynamicResource AccentBrush}"/><Setter Property="Foreground" Value="White"/>
  </Style>
  <Style TargetType="TextBox">
   <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/><Setter Property="Background" Value="{DynamicResource InputBgBrush}"/>
   <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush2}"/><Setter Property="CaretBrush" Value="{DynamicResource TextBrush}"/>
   <Setter Property="Padding" Value="8,6"/><Setter Property="MinHeight" Value="32"/>
  </Style>
  <Style TargetType="CheckBox"><Setter Property="Foreground" Value="{DynamicResource TextBrush}"/><Setter Property="Margin" Value="0,6,12,6"/><Setter Property="VerticalContentAlignment" Value="Center"/></Style>
  <Style TargetType="ComboBox">
   <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/><Setter Property="Background" Value="{DynamicResource InputBgBrush}"/>
   <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush2}"/><Setter Property="MinHeight" Value="32"/>
   <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ComboBox">
    <Grid>
     <ToggleButton Focusable="False" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}">
      <ToggleButton.Template><ControlTemplate TargetType="ToggleButton"><Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="3"><TextBlock Text="&#xE70D;" FontFamily="Segoe MDL2 Assets" HorizontalAlignment="Right" Margin="0,0,10,0" VerticalAlignment="Center"/></Border></ControlTemplate></ToggleButton.Template>
     </ToggleButton>
     <ContentPresenter Content="{TemplateBinding SelectionBoxItem}" Margin="9,5,30,5" IsHitTestVisible="False" VerticalAlignment="Center"/>
     <Popup Name="PART_Popup" IsOpen="{TemplateBinding IsDropDownOpen}" Placement="Bottom" AllowsTransparency="True" Focusable="False">
      <Border Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="1" MinWidth="{Binding ActualWidth, RelativeSource={RelativeSource TemplatedParent}}">
       <ScrollViewer MaxHeight="300"><ItemsPresenter/></ScrollViewer>
      </Border>
     </Popup>
    </Grid>
   </ControlTemplate></Setter.Value></Setter>
  </Style>
  <Style TargetType="ComboBoxItem">
   <Setter Property="Foreground" Value="{DynamicResource TextBrush}"/><Setter Property="Padding" Value="9,7"/>
   <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ComboBoxItem"><Border x:Name="ItemBorder" Padding="{TemplateBinding Padding}" Background="Transparent"><ContentPresenter/></Border><ControlTemplate.Triggers><Trigger Property="IsHighlighted" Value="True"><Setter TargetName="ItemBorder" Property="Background" Value="{DynamicResource SelectionBrush}"/></Trigger><Trigger Property="IsSelected" Value="True"><Setter TargetName="ItemBorder" Property="Background" Value="{DynamicResource SelectionBrush}"/></Trigger></ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter>
  </Style>
  <Style x:Key="NavItem" TargetType="RadioButton">
   <Setter Property="Foreground" Value="{DynamicResource SubTextBrush}"/><Setter Property="Margin" Value="0,2"/><Setter Property="Cursor" Value="Hand"/>
   <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="RadioButton"><Border x:Name="NavBorder" Background="Transparent" BorderBrush="Transparent" BorderThickness="3,0,0,0" Padding="12,12"><ContentPresenter/></Border><ControlTemplate.Triggers>
    <Trigger Property="IsChecked" Value="True"><Setter TargetName="NavBorder" Property="Background" Value="{DynamicResource SelectionBrush}"/><Setter TargetName="NavBorder" Property="BorderBrush" Value="{DynamicResource AccentBrush}"/><Setter Property="Foreground" Value="{DynamicResource TextBrush}"/></Trigger>
    <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="NavBorder" Property="Background" Value="{DynamicResource SelectionBrush}"/></Trigger>
   </ControlTemplate.Triggers></ControlTemplate></Setter.Value></Setter>
  </Style>
  <Style TargetType="TabControl"><Setter Property="Template"><Setter.Value><ControlTemplate TargetType="TabControl"><ContentPresenter ContentSource="SelectedContent"/></ControlTemplate></Setter.Value></Setter></Style>
  <Style TargetType="DataGrid">
   <Setter Property="Background" Value="{DynamicResource InputBgBrush}"/><Setter Property="Foreground" Value="{DynamicResource TextBrush}"/>
   <Setter Property="BorderBrush" Value="{DynamicResource BorderBrush2}"/><Setter Property="RowBackground" Value="{DynamicResource InputBgBrush}"/>
   <Setter Property="AlternatingRowBackground" Value="{DynamicResource SurfaceBrush}"/><Setter Property="GridLinesVisibility" Value="Horizontal"/>
   <Setter Property="HorizontalGridLinesBrush" Value="{DynamicResource BorderBrush2}"/><Setter Property="HeadersVisibility" Value="Column"/>
   <Setter Property="IsReadOnly" Value="True"/><Setter Property="AutoGenerateColumns" Value="True"/><Setter Property="CanUserAddRows" Value="False"/>
    <Setter Property="EnableRowVirtualization" Value="True"/><Setter Property="ColumnWidth" Value="Auto"/><Setter Property="MinRowHeight" Value="30"/>
  </Style>
  <Style TargetType="DataGridColumnHeader"><Setter Property="Background" Value="{DynamicResource SurfaceBrush}"/><Setter Property="Foreground" Value="{DynamicResource SubTextBrush}"/><Setter Property="Padding" Value="10,8"/><Setter Property="BorderBrush" Value="{DynamicResource BorderBrush2}"/><Setter Property="BorderThickness" Value="0,0,1,1"/></Style>
  <Style TargetType="DataGridCell"><Setter Property="Padding" Value="8,5"/><Setter Property="BorderThickness" Value="0"/><Style.Triggers><Trigger Property="IsSelected" Value="True"><Setter Property="Background" Value="{DynamicResource SelectionBrush}"/><Setter Property="Foreground" Value="{DynamicResource TextBrush}"/></Trigger></Style.Triggers></Style>
  <Style TargetType="ToolTip"><Setter Property="Background" Value="{DynamicResource SurfaceBrush}"/><Setter Property="Foreground" Value="{DynamicResource TextBrush}"/><Setter Property="BorderBrush" Value="{DynamicResource BorderBrush2}"/></Style>
 </Window.Resources>
 <Grid>
  <Grid.RowDefinitions><RowDefinition Height="64"/><RowDefinition Height="*"/></Grid.RowDefinitions>
  <Border Background="{DynamicResource SurfaceBrush}" BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="0,0,0,1">
   <Grid Margin="20,0,12,0"><Grid.ColumnDefinitions><ColumnDefinition Width="230"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
    <StackPanel Orientation="Horizontal" VerticalAlignment="Center"><TextBlock Text="Entra Toolbox" FontSize="22" FontWeight="SemiBold"/><TextBlock Text="v1.0" Margin="10,7,0,0" Foreground="{DynamicResource SubTextBrush}"/></StackPanel>
    <Button x:Name="ThemeButton" Grid.Column="2" VerticalAlignment="Center" Content="&#xE708;" FontFamily="Segoe MDL2 Assets" ToolTip="Switch light / dark theme"/>
   </Grid>
  </Border>
  <Grid Grid.Row="1"><Grid.ColumnDefinitions><ColumnDefinition Width="220"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
   <Border Background="{DynamicResource SurfaceBrush}" BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="0,0,1,0">
    <Grid>
     <Grid.RowDefinitions><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
     <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
      <StackPanel x:Name="Navigation" Margin="10,18,10,12"/>
     </ScrollViewer>
     <Border x:Name="ConnectionPanel" Grid.Row="1" BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="0,1,0,0" Padding="14,12">
      <StackPanel>
       <TextBlock Text="Microsoft Graph" FontSize="10" FontWeight="SemiBold" Foreground="{DynamicResource SubTextBrush}" Margin="0,0,0,7"/>
       <StackPanel Orientation="Horizontal" Margin="0,0,0,7">
        <Ellipse x:Name="ConnectionStatusLight" Width="8" Height="8" Fill="{DynamicResource ErrorBrush}" VerticalAlignment="Center" Margin="0,0,7,0" ToolTip="Not connected"/>
        <TextBlock x:Name="ConnectionText" Text="Not connected" FontSize="11" VerticalAlignment="Center"/>
       </StackPanel>
       <TextBlock x:Name="AccountText" Text="" FontSize="11" FontWeight="SemiBold" TextTrimming="CharacterEllipsis" Margin="0,0,0,4"/>
       <TextBlock x:Name="TenantText" Text="" FontSize="11" Foreground="{DynamicResource SubTextBrush}" TextTrimming="CharacterEllipsis" Margin="0,0,0,10"/>
       <StackPanel Orientation="Horizontal"><Button x:Name="ConnectButton" Content="Sign in" ToolTip="Sign in to Microsoft Graph"/><Button x:Name="DisconnectButton" Content="Sign out" ToolTip="Sign out of Microsoft Graph" IsEnabled="False"/></StackPanel>
      </StackPanel>
     </Border>
    </Grid>
   </Border>
   <Grid Grid.Column="1" Margin="22,18,22,14"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
    <StackPanel Margin="0,0,0,16"><TextBlock x:Name="PageTitle" FontSize="23" FontWeight="SemiBold"/><TextBlock x:Name="PageDescription" Margin="0,5,0,0" Foreground="{DynamicResource SubTextBrush}" TextWrapping="Wrap"/></StackPanel>
    <TabControl x:Name="PagesHost" Grid.Row="1"/>
   </Grid>
  </Grid>
 </Grid>
</Window>
'@
$script:Window = [Windows.Markup.XamlReader]::Load([Xml.XmlNodeReader]::new($xaml))
$script:PagesHost = $script:Window.FindName('PagesHost')
$script:Navigation = $script:Window.FindName('Navigation')
$script:Pages = [Collections.Generic.List[object]]::new()
# Menu sections are created up front so menu order is independent of page creation order.
$script:NavSections = [ordered]@{}
foreach ($sectionName in @('User-Device Affinity','Groups','Users','Devices')) {
    $heading = [Windows.Controls.TextBlock]::new(); $heading.Text = $sectionName.ToUpperInvariant(); $heading.FontSize = 10; $heading.FontWeight = 'SemiBold'
    $heading.Margin = [Windows.Thickness]::new(14,$(if ($script:NavSections.Count) { 12 } else { 0 }),0,6)
    $heading.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty,'SubTextBrush')
    $items = [Windows.Controls.StackPanel]::new()
    [void]$script:Navigation.Children.Add($heading); [void]$script:Navigation.Children.Add($items)
    $script:NavSections[$sectionName] = $items
}

function New-EntraButton {
    param([string]$Text, [string]$Tip, [switch]$Primary, [string]$Glyph)
    $button = [Windows.Controls.Button]::new()
    if ($Primary) { $button.Style = $script:Window.FindResource('PrimaryBtn') }
    if ($Glyph) {
        $button.Content = $Glyph
        $button.FontFamily = [Windows.Media.FontFamily]::new('Segoe MDL2 Assets')
        $button.Width = 36
        $button.Padding = [Windows.Thickness]::new(6)
    } else { $button.Content = $Text }
    $button.ToolTip = $Tip
    [Windows.Automation.AutomationProperties]::SetName($button, $Tip)
    return $button
}

function Update-EntraCheckListSummary {
    param($Panel)
    $state = $Panel.Tag
    $selected = @($Panel.Children | Where-Object { $_.IsChecked } | ForEach-Object { [string]$_.Content })
    $state.Summary.Text = if ($selected.Count) { $selected -join ', ' } else { $state.Placeholder }
    $state.Summary.Opacity = if ($selected.Count) { 1 } else { 0.7 }
    $state.Toggle.ToolTip = if ($selected.Count) { "Selected: $($selected -join ', ')" } else { $state.Tip }
}

function Add-EntraField {
    param($Page, [string]$Key, [string]$Label, [string]$Kind = 'Text', [string[]]$Options, [string]$Default, [string]$Tip)
    $field = [Windows.Controls.StackPanel]::new()
    $field.Margin = [Windows.Thickness]::new(0,0,14,10)
    $labelBlock = [Windows.Controls.TextBlock]::new()
    $labelBlock.Text = $Label
    $labelBlock.Margin = [Windows.Thickness]::new(0,0,0,4)
    if ($Page.CompactForm) {
        $field.Margin = [Windows.Thickness]::new(0,0,14,4)
        $labelBlock.Margin = [Windows.Thickness]::new(0,0,0,2)
    }
    if ($Kind -eq 'Choice') {
        $inputControl = [Windows.Controls.ComboBox]::new()
        foreach ($option in $Options) { [void]$inputControl.Items.Add($option) }
        $inputControl.SelectedIndex = 0
    } elseif ($Kind -eq 'Check') {
        $inputControl = [Windows.Controls.CheckBox]::new()
        $inputControl.Content = $Label
    } elseif ($Kind -eq 'CheckList') {
        # Multi-select dropdown: a ComboBox-styled toggle opening a popup of checkboxes (ComboBox itself closes on every click).
        $inputControl = [Windows.Controls.StackPanel]::new()
        foreach ($option in $Options) {
            $check = [Windows.Controls.CheckBox]::new()
            $check.Content = $option
            $check.Margin = [Windows.Thickness]::new(9,7,9,7)
            $check.Padding = [Windows.Thickness]::new(6,0,0,0)
            $check.Tag = $inputControl
            $check.Add_Checked({ param($sender,$eventArgs) Update-EntraCheckListSummary $sender.Tag })
            $check.Add_Unchecked({ param($sender,$eventArgs) Update-EntraCheckListSummary $sender.Tag })
            [void]$inputControl.Children.Add($check)
        }
        $toggle = [Windows.Markup.XamlReader]::Parse(@'
<ToggleButton xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" MinHeight="32" HorizontalContentAlignment="Stretch">
 <ToggleButton.Template><ControlTemplate TargetType="ToggleButton">
  <Border Background="{DynamicResource InputBgBrush}" BorderBrush="{DynamicResource BorderBrush2}" BorderThickness="1" CornerRadius="3">
   <Grid><TextBlock Text="&#xE70D;" FontFamily="Segoe MDL2 Assets" HorizontalAlignment="Right" Margin="0,0,10,0" VerticalAlignment="Center" Foreground="{DynamicResource TextBrush}"/>
   <ContentPresenter Margin="9,5,30,5" VerticalAlignment="Center"/></Grid>
  </Border>
 </ControlTemplate></ToggleButton.Template>
</ToggleButton>
'@)
        $summary = [Windows.Controls.TextBlock]::new()
        $summary.TextTrimming = 'CharacterEllipsis'
        $summary.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, 'TextBrush')
        $toggle.Content = $summary
        $toggle.ToolTip = $Tip
        [Windows.Automation.AutomationProperties]::SetName($toggle, $Label)
        $popupBorder = [Windows.Controls.Border]::new()
        $popupBorder.BorderThickness = [Windows.Thickness]::new(1)
        $popupBorder.SetResourceReference([Windows.Controls.Border]::BackgroundProperty, 'InputBgBrush')
        $popupBorder.SetResourceReference([Windows.Controls.Border]::BorderBrushProperty, 'BorderBrush2')
        $popupScroll = [Windows.Controls.ScrollViewer]::new()
        $popupScroll.MaxHeight = 300
        $popupScroll.VerticalScrollBarVisibility = 'Auto'
        $popupScroll.Content = $inputControl
        $popupBorder.Child = $popupScroll
        $minWidthBinding = [Windows.Data.Binding]::new('ActualWidth'); $minWidthBinding.Source = $toggle
        [void]$popupBorder.SetBinding([Windows.FrameworkElement]::MinWidthProperty, $minWidthBinding)
        $popup = [Windows.Controls.Primitives.Popup]::new()
        $popup.PlacementTarget = $toggle
        $popup.Placement = 'Bottom'
        $popup.StaysOpen = $false
        $popup.AllowsTransparency = $true
        $popup.Child = $popupBorder
        $popup.Tag = $toggle
        $openBinding = [Windows.Data.Binding]::new('IsChecked'); $openBinding.Source = $toggle; $openBinding.Mode = 'TwoWay'
        [void]$popup.SetBinding([Windows.Controls.Primitives.Popup]::IsOpenProperty, $openBinding)
        # While open, the outside click that closes the popup must not immediately re-toggle it.
        $popup.Add_Opened({ param($sender,$eventArgs) $sender.Tag.IsHitTestVisible = $false })
        $popup.Add_Closed({ param($sender,$eventArgs) $sender.Tag.IsHitTestVisible = $true })
        $popupBorder.Tag = $toggle
        $popupBorder.Add_PreviewKeyDown({ param($sender,$eventArgs) if ($eventArgs.Key -eq 'Escape') { $sender.Tag.IsChecked = $false; [void]$sender.Tag.Focus(); $eventArgs.Handled = $true } })
        $toggle.Add_IsEnabledChanged({ param($sender,$eventArgs) if (-not $sender.IsEnabled) { $sender.IsChecked = $false } })
        $inputControl.Tag = @{ Toggle = $toggle; Summary = $summary; Popup = $popup; Tip = $Tip; Placeholder = $(if ($Default) { $Default } else { 'None selected' }) }
        Update-EntraCheckListSummary $inputControl
    } else {
        $inputControl = [Windows.Controls.TextBox]::new()
        $inputControl.Text = $Default
        if ($Kind -eq 'Lines') {
            $inputControl.AcceptsReturn = $true
            $inputControl.Height = if ($Page.CompactForm) { 60 } else { 74 }
            $inputControl.VerticalScrollBarVisibility = 'Auto'
        }
    }
    $inputControl.ToolTip = $Tip
    [Windows.Automation.AutomationProperties]::SetName($inputControl, $Label)
    if ($Kind -ne 'Check') { [void]$field.Children.Add($labelBlock) }
    if ($Kind -in @('File', 'Save')) {
        $fileGrid = [Windows.Controls.Grid]::new()
        [void]$fileGrid.ColumnDefinitions.Add([Windows.Controls.ColumnDefinition]::new())
        $browseColumn = [Windows.Controls.ColumnDefinition]::new(); $browseColumn.Width = [Windows.GridLength]::Auto
        [void]$fileGrid.ColumnDefinitions.Add($browseColumn)
        $browse = New-EntraButton -Glyph ([string][char]0xE8B7) -Tip "Browse: $Label"
        $browse.Margin = [Windows.Thickness]::new(6,0,0,0)
        [Windows.Controls.Grid]::SetColumn($browse,1)
        $browse.Tag = @{ Kind=$Kind; Key=$Key; Label=$Label; Input=$inputControl }
        $browse.Add_Click({
            param($sender,$eventArgs)
            $fieldInfo = $sender.Tag
            if ($fieldInfo.Kind -eq 'Save') { $dialog = [Microsoft.Win32.SaveFileDialog]::new(); $dialog.OverwritePrompt = $true; $dialog.DefaultExt = '.csv' }
            else { $dialog = [Microsoft.Win32.OpenFileDialog]::new() }
            $dialog.Filter = if ($fieldInfo.Key -eq 'MemberFile') { 'Text files (*.txt)|*.txt|All files (*.*)|*.*' } else { 'CSV files (*.csv)|*.csv|All files (*.*)|*.*' }
            $dialog.Title = $fieldInfo.Label
            if ($dialog.ShowDialog($script:Window)) { $fieldInfo.Input.Text = $dialog.FileName }
        })
        [void]$fileGrid.Children.Add($inputControl); [void]$fileGrid.Children.Add($browse)
        [void]$field.Children.Add($fileGrid)
    } elseif ($Kind -eq 'CheckList') {
        [void]$field.Children.Add($inputControl.Tag.Toggle)
        [void]$field.Children.Add($inputControl.Tag.Popup)
    } else { [void]$field.Children.Add($inputControl) }
    [void]$Page.Form.Children.Add($field)
    $Page.Controls[$Key] = $inputControl
    $Page.Defaults[$Key] = @{ Kind=$Kind; Value=$Default }
    $Page.Fields[$Key] = $field
}

function Write-EntraActivity {
    param($Page, [string]$Message, [string]$Level = 'INFO')
    if ([string]::IsNullOrWhiteSpace($Message)) { return }
    $stamp = Get-Date -Format 'HH:mm:ss'
    $brushKey = switch ($Level) { 'ERROR' { 'ErrorBrush' } 'WARN' { 'WarningBrush' } 'OK' { 'SuccessBrush' } default { 'SubTextBrush' } }
    $paragraph = [Windows.Documents.Paragraph]::new()
    $paragraph.Margin = [Windows.Thickness]::new(0)
    $prefix = [Windows.Documents.Run]::new("$stamp  [$Level]  ")
    $prefix.SetResourceReference([Windows.Documents.TextElement]::ForegroundProperty, $brushKey)
    [void]$paragraph.Inlines.Add($prefix)
    [void]$paragraph.Inlines.Add([Windows.Documents.Run]::new($Message))
    [void]$Page.Log.Document.Blocks.Add($paragraph)
    if ($Page.Log.Document.Blocks.Count -gt 5000) { $Page.Log.Document.Blocks.Remove($Page.Log.Document.Blocks.FirstBlock) }
    $Page.Log.ScrollToEnd()
    $Page.ActivitySummary.Text = "$stamp  [$Level]  $($Message -replace '\r?\n',' ')"
    $Page.ActivitySummary.ToolTip = $Message
    if ($Level -in @('ERROR','WARN')) { $Page.Activity.IsExpanded = $true }
}

function Get-EntraOutcomeCategory {
    param($Result)
    $values = @($Result.PSObject.Properties | Where-Object { $_.Name -in @('Status','Membership','ResolutionStatus','LookupError') } | ForEach-Object { [string]$_.Value } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $outcome = $values -join ' '
    if ($outcome -match '(?i)\bwhatif\b|\bpreview\b') { return 'Preview' }
    if ($outcome -match '(?i)fail|error|not found|missing|blocked') { return 'Failed' }
    if ($outcome -match '(?i)skip|already\s+(?:a\s+)?member|already\s+exists|already\s+disabled|same as target') { return 'Skipped' }
    if ($outcome -match '(?i)\b(?:added|created|reused|ok|deleted|removed|disabled|usersloggedonresolved|primaryuserresolved)\b') { return 'Success' }
    return 'Other'
}

function Format-EntraResultWarning {
    param($Result)
    $deviceNameProperty = $Result.PSObject.Properties['DeviceName']
    if ($deviceNameProperty -and -not [string]::IsNullOrWhiteSpace([string]$deviceNameProperty.Value)) {
        $lookupErrorProperty = $Result.PSObject.Properties['LookupError']
        $errorMessage = if ($lookupErrorProperty) { [string]$lookupErrorProperty.Value } else { '' }
        if ([string]::IsNullOrWhiteSpace($errorMessage)) {
            $errorMessage = @($Result.PSObject.Properties | Where-Object { $_.Name -in @('Status','Membership','ResolutionStatus') -and [string]$_.Value -match '(?i)fail|error|not found|missing|blocked' } | ForEach-Object { [string]$_.Value }) -join '; '
        }
        if ([string]::IsNullOrWhiteSpace($errorMessage)) { $errorMessage = 'Unspecified device warning' }
        return "DeviceName: $($deviceNameProperty.Value); Error: $errorMessage"
    }
    return ($Result | Out-String -Width 180).Trim()
}

function Add-EntraRunSummaryEntry {
    param($Job, [string]$Level, [string]$Message)
    if ($Job.Kind -eq 'Run' -and $Level -in @('INFO','WARN','ERROR') -and -not [string]::IsNullOrWhiteSpace($Message)) {
        [void]$Job.ActivityEntries.Add([pscustomobject]@{ Time=(Get-Date -Format 'HH:mm:ss'); Level=$Level; Message=$Message.Trim() })
    }
}

function Format-EntraRunSummary {
    param($Job)
    $totals = "Run summary$(if ($Job.Cancelled) { ' (partial)' }): Processed $($Job.Counts.Processed); Success $($Job.Counts.Success); Failed $($Job.Counts.Failed); Skipped $($Job.Counts.Skipped); Preview $($Job.Counts.Preview); Other $($Job.Counts.Other)"
    if ($Job.ExecutionErrors) { $totals += "; Execution errors $($Job.ExecutionErrors)" }
    $summary = ''
    $activityEntries = $Job.ActivityEntries.ToArray()
    $information = @($activityEntries | Where-Object Level -eq 'INFO')
    $warnings = @($activityEntries | Where-Object Level -eq 'WARN')
    $errors = @($activityEntries | Where-Object Level -eq 'ERROR')
    if ($information.Count) { $summary += "`nActivity:`n" + (($information | ForEach-Object { "  $($_.Time)  [INFO]  $($_.Message)" }) -join "`n") }
    if ($warnings.Count) { $summary += "`nWarnings ($($warnings.Count)):`n" + (($warnings | ForEach-Object { "  $($_.Time)  [WARN]  $($_.Message)" }) -join "`n") }
    if ($errors.Count) { $summary += "`nErrors ($($errors.Count)):`n" + (($errors | ForEach-Object { "  $($_.Time)  [ERROR]  $($_.Message)" }) -join "`n") }
    return ($summary.Trim() + "`n" + $totals).Trim()
}

function Get-EntraLines {
    param([string]$Text)
    return @($Text -split '\r?\n' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Read-EntraCsv {
    param([string]$Path, [string[]]$Headers, [char]$Delimiter = ',', [switch]$AnyHeader)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "File not found: $Path" }
    $rows = @(Import-Csv -LiteralPath $Path -Delimiter $Delimiter -ErrorAction Stop)
    if ($rows.Count -eq 0) { throw 'The CSV contains no data rows.' }
    $names = $rows[0].PSObject.Properties.Name
    $matched = @($Headers | Where-Object { $_ -in $names })
    if (($AnyHeader -and $matched.Count -eq 0) -or (-not $AnyHeader -and $matched.Count -ne $Headers.Count)) {
        throw "Required CSV headers: $($Headers -join ', ')"
    }
    return $rows
}

function Export-EntraResults {
    param($Page, [string]$Path)
    # -NoTypeInformation is required on Windows PowerShell 5.1 but obsolete (warns) on PowerShell 7.
    $exportOptions = @{ Encoding = $script:Utf8BomEncoding }
    if ($PSVersionTable.PSVersion.Major -lt 6) { $exportOptions.NoTypeInformation = $true }
    $Page.Results | Export-Csv -LiteralPath $Path @exportOptions -ErrorAction Stop
}

function Update-EntraForm {
    param($Page)
    $controls = $Page.Controls
    switch ($Page.Id) {
        'Members' {
            $fromFile = $controls.Source.SelectedItem -eq 'Text file'
            $Page.Fields.MemberNames.Visibility = if ($fromFile) { 'Collapsed' } else { 'Visible' }
            $Page.Fields.MemberFile.Visibility = if ($fromFile) { 'Visible' } else { 'Collapsed' }
        }
        'PrimaryPC' {
            $sourceFields = @{ 'Group name'='GroupName'; 'Group ID'='GroupId'; 'User CSV'='CsvPath'; 'User UPNs'='UserPrincipalName' }
            foreach ($source in $sourceFields.Keys) {
                $Page.Fields[$sourceFields[$source]].Visibility = if ($controls.Source.SelectedItem -eq $source) { 'Visible' } else { 'Collapsed' }
            }
            $signInOnly = $controls.DeviceSource.SelectedItem -eq 'Sign-in logs only'
            $deviceModes = if ($signInOnly) { @('Most logged-on PCs','All logged-on PCs') } else { @('Latest','All') }
            if ($controls.DeviceMode.Items[0] -ne $deviceModes[0]) {
                $allDevices = $controls.DeviceMode.SelectedItem -in @('All','All logged-on PCs')
                $controls.DeviceMode.Items.Clear()
                foreach ($mode in $deviceModes) { [void]$controls.DeviceMode.Items.Add($mode) }
                $controls.DeviceMode.SelectedIndex = if ($allDevices) { 1 } else { 0 }
            }
            $Page.Fields.SignInFallback.Visibility = if ($signInOnly) { 'Collapsed' } else { 'Visible' }
            if ($signInOnly) { $controls.SignInFallback.IsChecked = $false }
            $controls.SignInDays.IsEnabled = $signInOnly -or [bool]$controls.SignInFallback.IsChecked
            Update-EntraRunAvailability $Page
        }
        'LastUser' {
            $Page.Fields.DeviceName.Visibility = if ($controls.Source.SelectedItem -eq 'Device names') { 'Visible' } else { 'Collapsed' }
            $Page.Fields.CsvPath.Visibility = if ($controls.Source.SelectedItem -eq 'CSV file') { 'Visible' } else { 'Collapsed' }
            $Page.Fields.GroupNameOrId.Visibility = if ($controls.Source.SelectedItem -eq 'Entra group') { 'Visible' } else { 'Collapsed' }
        }
        'DeviceCleanup' {
            $sourceFields = @{ 'CSV file'='CsvPath'; 'Entra lookup'='InactivityDays'; 'Device names'='DeviceNames' }
            foreach ($source in $sourceFields.Keys) {
                $Page.Fields[$sourceFields[$source]].Visibility = if ($controls.Source.SelectedItem -eq $source) { 'Visible' } else { 'Collapsed' }
            }
        }
        'Attributes' {
            $fromFile = $controls.Source.SelectedItem -eq 'CSV file'
            $Page.Fields.UserPrincipalName.Visibility = if ($fromFile) { 'Collapsed' } else { 'Visible' }
            $Page.Fields.CsvPath.Visibility = if ($fromFile) { 'Visible' } else { 'Collapsed' }
            $controls.Value.IsEnabled = $controls.Action.SelectedItem -eq 'Set'
        }
    }
    if ($Page.CompactForm) {
        $Page.Form.RowDefinitions.Clear()
        $fieldIndex = 0
        foreach ($field in $Page.Form.Children) {
            if ($field.Visibility -eq 'Collapsed') { continue }
            $rowIndex = [int][Math]::Floor($fieldIndex / 2)
            if ($fieldIndex % 2 -eq 0) {
                $row = [Windows.Controls.RowDefinition]::new()
                $row.Height = [Windows.GridLength]::Auto
                [void]$Page.Form.RowDefinitions.Add($row)
            }
            [Windows.Controls.Grid]::SetRow($field,$rowIndex)
            [Windows.Controls.Grid]::SetColumn($field,($fieldIndex % 2))
            $fieldIndex++
        }
    }
}

function Get-EntraInvocation {
    param($Page)
    $controls = $Page.Controls
    $parameters = @{ SkipGraphConnect=$true; PassThru=$true }
    $summary = ''
    $preview = $false
    switch ($Page.Id) {
        'DeviceCleanup' {
            $parameters.Action = 'Inspect'
            switch ([string]$controls.Source.SelectedItem) {
                'Entra lookup' {
                    $days = 0
                    if (-not [int]::TryParse($controls.InactivityDays.Text.Trim(), [ref]$days) -or $days -lt 1 -or $days -gt 3650) { throw 'Inactive days must be a whole number from 1 to 3650.' }
                    $parameters.Source = 'Entra'
                    $parameters.InactivityDays = $days
                }
                'Device names' {
                    $names = @(Get-EntraLines $controls.DeviceNames.Text)
                    if (-not $names.Count) { throw 'Enter at least one device name.' }
                    $parameters.Source = 'Names'
                    $parameters.DeviceName = [string[]]$names
                }
                default {
                    $path = $controls.CsvPath.Text.Trim()
                    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw 'Select an existing device CSV.' }
                    $parameters.CsvPath = $path
                }
            }
        }
        'Members' {
            $target = $controls.TargetGroupName.Text.Trim()
            if (-not $target) { throw 'Enter a target group name.' }
            $parameters.TargetGroupName = $target
            $parameters.MemberType = [string]$controls.MemberType.SelectedItem
            if ($controls.Source.SelectedItem -eq 'Text file') {
                $path = $controls.MemberFile.Text.Trim()
                if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw 'Select a valid member text file.' }
                $entries = @(Get-Content -LiteralPath $path | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                $parameters.MemberFile = $path
            } else { $entries = @(Get-EntraLines $controls.MemberNames.Text); $parameters.MemberNames = [string[]]$entries }
            if ($entries.Count -eq 0) { throw 'Enter at least one member.' }
            $preview = [bool]$controls.WhatIf.IsChecked
            $parameters.WhatIf = $preview; $parameters.Confirm = $false
            $summary = "Add $($entries.Count) $($parameters.MemberType.ToLowerInvariant()) to group '$target'."
        }
        'Groups' {
            $path = $controls.CsvPath.Text.Trim()
            $rows = @(Read-EntraCsv -Path $path -Headers @('GroupName','Description') -Delimiter ';')
            if (@($rows | Where-Object { [string]::IsNullOrWhiteSpace($_.GroupName) }).Count) { throw 'Every CSV row needs a GroupName.' }
            $parameters.CsvPath = $path
            $summary = "Create or reuse $($rows.Count) security group(s) and add any nested group memberships in the CSV."
        }
        'PrimaryPC' {
            switch ($controls.Source.SelectedItem) {
                'Group name' {
                    $parameters.GroupName = $controls.GroupName.Text.Trim()
                    if (-not $parameters.GroupName) { throw 'Enter a group name.' }
                }
                'Group ID' {
                    $parameters.GroupId = $controls.GroupId.Text.Trim()
                    $parsedId = [guid]::Empty
                    if (-not [guid]::TryParse($parameters.GroupId,[ref]$parsedId)) { throw 'Group ID must be a GUID.' }
                }
                'User CSV' {
                    $parameters.CsvPath = $controls.CsvPath.Text.Trim()
                    [void](Read-EntraCsv -Path $parameters.CsvPath -Headers @('UserPrincipalName'))
                }
                'User UPNs' {
                    $upns = @(Get-EntraLines $controls.UserPrincipalName.Text)
                    if (-not $upns.Count) { throw 'Enter at least one user UPN.' }
                    $parameters.UserPrincipalName = [string[]]$upns
                }
                default { throw 'Choose a group name, group ID, user CSV, or user UPNs.' }
            }
            $parameters.NonInteractive = $true
            $parameters.DeviceSource = switch ($controls.DeviceSource.SelectedItem) {
                'Entra owned/registered devices' { 'Entra' }
                'Sign-in logs only' { 'SignInLogs' }
                default { throw 'Choose Entra devices or sign-in logs only.' }
            }
            $parameters.DeviceMode = switch ($controls.DeviceMode.SelectedItem) {
                'Latest' { 'Latest' }
                'All' { 'All' }
                'Most logged-on PCs' { 'Most' }
                'All logged-on PCs' { 'All' }
                default { throw 'Choose a device selection.' }
            }
            $parameters.OnlyIntuneDevices = [bool]$controls.OnlyIntuneDevices.IsChecked
            $parameters.Verbose = [bool]$controls.Verbose.IsChecked
            $parameters.SignInFallback = $parameters.DeviceSource -eq 'Entra' -and [bool]$controls.SignInFallback.IsChecked
            if ($parameters.SignInFallback -or $parameters.DeviceSource -eq 'SignInLogs') {
                $days = 0
                if (-not [int]::TryParse($controls.SignInDays.Text,[ref]$days) -or $days -lt 1 -or $days -gt 365) { throw 'Sign-in days must be 1-365.' }
                $parameters.SignInDays = $days
            }
        }
        'LastUser' {
            switch ($controls.Source.SelectedItem) {
                'Device names' { $names = @(Get-EntraLines $controls.DeviceName.Text); if (-not $names.Count) { throw 'Enter at least one device name.' }; $parameters.DeviceName = [string[]]$names }
                'CSV file' { $parameters.CsvPath = $controls.CsvPath.Text.Trim(); [void](Read-EntraCsv -Path $parameters.CsvPath -Headers @('DeviceId','ManagedDeviceId','Id','DeviceName','ManagedDeviceName','Name','displayName') -AnyHeader) }
                'Entra group' { $groupNameOrId = $controls.GroupNameOrId.Text.Trim(); if (-not $groupNameOrId) { throw 'Enter an Entra group name or ID.' }; $parameters.GroupNameOrId = $groupNameOrId; $summary = "Report on direct device members of Entra group '$groupNameOrId'." }
                'All devices' { }
                default { throw 'Choose device names, an Entra group, a CSV, or explicitly select All devices.' }
            }
            $parameters.UserMode = switch ($controls.UserMode.SelectedItem) {
                'Latest logged-on user' { 'Latest' }
                'Most logged-on user(s)' { 'Most' }
                'All logged-on users' { 'All' }
                default { throw 'Choose latest, most, or all logged-on users.' }
            }
            $selectedAttributes = @($controls.ExtensionAttributeNumbers.Children | Where-Object { $_.IsChecked } |
                ForEach-Object { [int]([string]$_.Content -replace '^extensionAttribute', '') })
            if ($selectedAttributes.Count) { $parameters.ExtensionAttributeNumbers = [int[]]$selectedAttributes }
        }
        'Attributes' {
            if ($controls.Source.SelectedItem -eq 'CSV file') {
                $parameters.CsvPath = $controls.CsvPath.Text.Trim()
                $users = @(Read-EntraCsv -Path $parameters.CsvPath -Headers @('UserPrincipalName'))
                if (@($users | Where-Object { [string]::IsNullOrWhiteSpace($_.UserPrincipalName) }).Count) { throw 'Every CSV row needs a UserPrincipalName.' }
            } else {
                $users = @(Get-EntraLines $controls.UserPrincipalName.Text)
                if (-not $users.Count) { throw 'Enter at least one UPN.' }
                $parameters.UserPrincipalName = [string[]]$users
            }
            $parameters.AttributeNumber = [int]$controls.AttributeNumber.SelectedItem
            $parameters.NonInteractive = $true
            if ($controls.Action.SelectedItem -eq 'Clear') { $parameters.Clear = $true; $displayValue = '(clear)' }
            else { if ([string]::IsNullOrWhiteSpace($controls.Value.Text)) { throw 'Enter a value or select Clear.' }; $parameters.Value = $controls.Value.Text; $displayValue = $parameters.Value }
            $preview = [bool]$controls.WhatIf.IsChecked; $parameters.WhatIf = $preview; $parameters.Confirm = $false
            $summary = "Update $($users.Count) user(s): extensionAttribute$($parameters.AttributeNumber) = $displayValue."
        }
    }
    if ($Page.Id -in @('PrimaryPC','LastUser')) {
        $path = $controls.OutputPath.Text.Trim()
        if (-not $path -or -not [IO.Path]::IsPathRooted($path) -or [IO.Path]::GetExtension($path) -ine '.csv') { throw 'Choose an absolute CSV output path.' }
        $parent = Split-Path -Parent $path
        if (-not (Test-Path -LiteralPath $parent -PathType Container)) { throw 'The output folder does not exist.' }
        $parameters.OutputPath = $path
    }
    return @{ Parameters=$parameters; Summary=$summary; Preview=$preview }
}

function Get-EntraRequirements {
    param($Page, [hashtable]$Parameters)
    $modules = @('Microsoft.Graph.Authentication') + $Page.Modules
    $scopes = @($Page.Scopes)
    if ($Page.Id -eq 'DeviceCleanup' -and (($Parameters['Action'] -eq 'Delete' -and -not $Parameters['WhatIf']) -or
        ($Parameters.Count -eq 0 -and -not $Page.Controls.Preview.IsChecked))) {
        $scopes += @('Directory.AccessAsUser.All','DeviceManagementManagedDevices.ReadWrite.All','DeviceManagementServiceConfig.ReadWrite.All')
    }
    if ($Page.Id -eq 'DeviceCleanup' -and $Parameters['Action'] -eq 'Disable' -and -not $Parameters['WhatIf']) { $scopes += 'Directory.AccessAsUser.All' }
    if ($Page.Id -eq 'Members') {
        if ($Parameters.MemberType -eq 'Users') { $modules += 'Microsoft.Graph.Users' }
        if ($Parameters.MemberType -eq 'Devices') { $modules += 'Microsoft.Graph.Identity.DirectoryManagement' }
    }
    if ($Page.Id -eq 'PrimaryPC' -and ($Parameters['SignInFallback'] -or $Parameters['DeviceSource'] -eq 'SignInLogs' -or
        ($Parameters.Count -eq 0 -and ($Page.Controls.SignInFallback.IsChecked -or $Page.Controls.DeviceSource.SelectedItem -eq 'Sign-in logs only')))) {
        $modules += 'Microsoft.Graph.Reports'; $scopes += 'AuditLog.Read.All'
    }
    if ($Page.Id -eq 'PrimaryPC' -and ($Parameters['OnlyIntuneDevices'] -or ($Parameters.Count -eq 0 -and $Page.Controls.OnlyIntuneDevices.IsChecked))) { $scopes += 'DeviceManagementManagedDevices.Read.All' }
    if ($Page.Id -eq 'LastUser' -and $Parameters.GroupNameOrId) { $scopes += @('Device.Read.All','Group.Read.All','GroupMember.Read.All') }
    return @{ Modules=@($modules | Select-Object -Unique); Scopes=$scopes }
}

function Update-EntraResults {
    param($Page)
    $term = $Page.Filter.Text.Trim()
    if ($Page.Id -eq 'DeviceCleanup') {
        try {
            $visible = @(Select-EntraCleanupRows -Rows @($Page.Results.ToArray()) -Name $Page.Controls.NameFilter.Text `
                -JoinType ([string]$Page.Controls.JoinFilter.SelectedItem) -Managed ([string]$Page.Controls.ManagedFilter.SelectedItem) `
                -From $Page.Controls.DateFrom.Text -To $Page.Controls.DateTo.Text -DateMode ([string]$Page.Controls.DateMode.SelectedItem) -Search $term)
        } catch { $Page.Status.Text = $_.Exception.Message; $Page.Table.ItemsSource = @(); Update-EntraCleanupSelection $Page; return }
        $Page.Table.ItemsSource = $visible
        $Page.Count.Text = "$($visible.Count) shown / $($Page.Results.Count) results"
        $Page.Export.IsEnabled = $Page.Results.Count -gt 0
        Update-EntraCleanupSelection $Page
        return
    }
    $visible = @($Page.Results | Where-Object {
        if (-not $term) { $true } else { ($_.PSObject.Properties.Value -join ' ').IndexOf($term, [StringComparison]::OrdinalIgnoreCase) -ge 0 }
    })
    $Page.Table.ItemsSource = $visible
    $Page.Count.Text = "$($visible.Count) shown / $($Page.Results.Count) results"
    $Page.Export.IsEnabled = $Page.Results.Count -gt 0
}

function Select-EntraCleanupRows {
    param([array]$Rows, [string]$Name, [string]$JoinType = 'All', [string]$Managed = 'All',
        [string]$From, [string]$To, [string]$DateMode = 'All dates', [string]$Search)
    $startDate = $null; $endDate = $null
    foreach ($value in @($From,$To)) {
        if ($value -and $value -notmatch '^\d{4}-\d{2}-\d{2}$') { throw 'Activity dates must use yyyy-MM-dd (UTC).' }
    }
    if ($From) { $startDate = [DateTime]::ParseExact($From,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture) }
    if ($To) { $endDate = [DateTime]::ParseExact($To,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture).AddDays(1) }
    if ($startDate -and $endDate -and $startDate -ge $endDate) { throw 'Activity from date must not be after the to date.' }
    foreach ($row in $Rows) {
        if ($Name -and $row.'Device Name'.IndexOf($Name.Trim(),[StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
        if ($JoinType -and $JoinType -ne 'All' -and $row.'Join Type' -ne $JoinType) { continue }
        if ($Managed -and $Managed -ne 'All' -and $row.'Intune managed' -ne $Managed) { continue }
        $date = $row.'Activity date'
        if ($DateMode -eq 'Unknown dates' -and $null -ne $date) { continue }
        if ($DateMode -eq 'Known dates' -and $null -eq $date) { continue }
        if (($startDate -or $endDate) -and ($null -eq $date -or ($startDate -and $date -lt $startDate) -or ($endDate -and $date -ge $endDate))) { continue }
        if ($Search -and ($row.'Device Name',$row.'Join Type',$row.Enabled,$row.'Intune managed',$row.'Device ID',$row.Status -join ' ').IndexOf($Search,[StringComparison]::OrdinalIgnoreCase) -lt 0) { continue }
        $row
    }
}

function Get-EntraCleanupEligibleRows {
    param($Page)
    if (-not $script:ConnectedTenant) { return }
    foreach ($row in $Page.Table.ItemsSource) {
        if ($row.CanDelete -and $row.TenantId -eq $script:ConnectedTenant -and $row.Status -ne 'Deleted') { $row }
    }
}

function Get-EntraCleanupSelection {
    param($Page)
    $eligible = @(Get-EntraCleanupEligibleRows $Page)
    @($Page.Table.SelectedItems | Where-Object { $_ -in $eligible })
}

function Get-EntraCleanupDisableSelection {
    param($Page)
    @(Get-EntraCleanupSelection $Page | Where-Object { $_.Enabled -ne 'No' })
}

function Update-EntraCleanupSelection {
    param($Page)
    if (-not $Page.ContainsKey('DeleteSelected')) { return }
    $selected = @(Get-EntraCleanupSelection $Page)
    $Page.DeleteSelected.IsEnabled = -not $script:Job -and $selected.Count -gt 0
    $Page.DeleteSelected.ToolTip = "Delete or preview $($selected.Count) selected eligible device(s)"
    if ($Page.ContainsKey('DisableSelected')) {
        $disable = @(Get-EntraCleanupDisableSelection $Page)
        $Page.DisableSelected.IsEnabled = -not $script:Job -and $disable.Count -gt 0
        $Page.DisableSelected.ToolTip = "Disable or preview $($disable.Count) selected enabled device(s)"
    }
    if ($Page.ContainsKey('SelectAll')) {
        $eligible = @(Get-EntraCleanupEligibleRows $Page)
        $Page.SelectAll.IsEnabled = -not $script:Job -and $eligible.Count -gt $selected.Count
        $Page.SelectAll.ToolTip = "Select all visible eligible devices ($($eligible.Count) available)"
    }
    if ($Page.ContainsKey('ClearSelection')) {
        $Page.ClearSelection.IsEnabled = -not $script:Job -and $Page.Table.SelectedItems.Count -gt 0
        $Page.ClearSelection.ToolTip = "Clear selection ($($Page.Table.SelectedItems.Count) selected)"
    }
}

function Get-EntraCleanupDeletionInvocation {
    param($Page)
    $selected = @(Get-EntraCleanupSelection $Page)
    if (-not $selected.Count -or -not $script:ConnectedTenant) { throw 'Select eligible devices in the currently connected tenant.' }
    $preview = [bool]$Page.Controls.Preview.IsChecked
    if (-not $preview -and -not $Page.Controls.Acknowledge.IsChecked) { throw 'Acknowledge remaining cloud Entra deletion before deleting.' }
    $deviceLabel = if ($selected.Count -eq 1) { 'device' } else { 'devices' }
    return @{
        Parameters = @{ Action='Delete'; Device=$selected; ExpectedTenantId=$script:ConnectedTenant; AcknowledgeCloudDeletion=$true; SkipGraphConnect=$true; PassThru=$true; WhatIf=$preview; Confirm=$false }
        Preview = $preview
        Summary = "Delete $($selected.Count) selected ${deviceLabel}?"
    }
}

function Get-EntraCleanupDisableInvocation {
    param($Page)
    $selected = @(Get-EntraCleanupDisableSelection $Page)
    if (-not $selected.Count -or -not $script:ConnectedTenant) { throw 'Select enabled eligible devices in the currently connected tenant.' }
    $preview = [bool]$Page.Controls.Preview.IsChecked
    $deviceLabel = if ($selected.Count -eq 1) { 'device' } else { 'devices' }
    return @{
        Parameters = @{ Action='Disable'; Device=$selected; ExpectedTenantId=$script:ConnectedTenant; SkipGraphConnect=$true; PassThru=$true; WhatIf=$preview; Confirm=$false }
        Preview = $preview
        Summary = "Disable $($selected.Count) selected ${deviceLabel} in Entra? Users can no longer authenticate with the device until it is re-enabled."
    }
}

function Get-EntraConfirmationMessage {
    param($Page, [hashtable]$Invocation)
    if ($Page.Id -eq 'DeviceCleanup' -and $Invocation.Parameters['Action'] -eq 'Delete') {
        return $Invocation.Summary
    }
    return "$($Invocation.Summary)`n`nTenant: $script:ConnectedTenant`nChanges take effect immediately. Continue?"
}

function Update-EntraLayout {
    param($Page)
    if ($Page.Layout.ActualHeight -gt 0) {
        if ($Page.CompactForm) {
            if ($Page.Activity.IsExpanded) {
                $footerChrome = $Page.Footer.ActualHeight - $Page.Log.ActualHeight
                $available = $Page.Layout.ActualHeight - $Page.Form.DesiredSize.Height - $Page.Actions.ActualHeight - $Page.ResultsBar.ActualHeight - $Page.Table.MinHeight - $footerChrome - 4
                $Page.Log.Height = [Math]::Max(24,[Math]::Min(150,[Math]::Floor($available)))
            }
        } else {
            $available = $Page.Layout.ActualHeight - $Page.Footer.ActualHeight - $Page.Actions.ActualHeight - $Page.ResultsBar.ActualHeight - 125
            $Page.FormScroll.MaxHeight = [Math]::Max(70,[Math]::Min(285,$available))
        }
    }
}

function Update-EntraRunAvailability {
    param($Page, [bool]$Busy = [bool]$script:Job)
    $hasInput = $true
    if ($Page.Id -eq 'PrimaryPC') {
        $sourceFields = @{ 'Group name'='GroupName'; 'Group ID'='GroupId'; 'User CSV'='CsvPath'; 'User UPNs'='UserPrincipalName' }
        $selectedField = $sourceFields[[string]$Page.Controls.Source.SelectedItem]
        $hasInput = $selectedField -and -not [string]::IsNullOrWhiteSpace($Page.Controls[$selectedField].Text)
    }
    $Page.Run.IsEnabled = -not $Busy -and $Page.Form.IsEnabled -and $Page.Available -and $hasInput
}

function Set-EntraBusy {
    param([bool]$Busy)
    foreach ($page in $script:Pages) {
        $page.Form.IsEnabled = -not $Busy
        Update-EntraRunAvailability -Page $page -Busy $Busy
        $page.Reset.IsEnabled = -not $Busy
        $page.Stop.IsEnabled = $Busy -and $script:Job -and $script:Job.Page -eq $page
        if ($page.Id -eq 'DeviceCleanup') {
            $page.Filter.IsEnabled = -not $Busy
            $page.Table.IsEnabled = -not $Busy
            Update-EntraCleanupSelection $page
        }
    }
    $script:Window.FindName('ConnectButton').IsEnabled = -not $Busy
    $script:Window.FindName('DisconnectButton').IsEnabled = -not $Busy -and [bool]$script:ConnectedTenant
}

function Set-EntraConnectionState {
    param([bool]$Connected, [string]$Account, [string]$Tenant, [string]$StatusText, [string]$Detail)
    $script:ConnectedTenant = if ($Connected) { $Tenant } else { '' }
    $connectionLabel = if ($StatusText) { $StatusText } elseif ($Connected) { 'Connected' } else { 'Not connected' }
    if ($Connected -and $script:RequestedTenant -and $Tenant -ne $script:RequestedTenant) {
        $Detail = "The Graph session belongs to a different tenant than requested ($script:RequestedTenant). Sign out and reconnect before continuing."
    }
    $connectionText = $script:Window.FindName('ConnectionText')
    $connectionText.Text = $connectionLabel
    $connectionText.ToolTip = if ($Detail) { $Detail } else { $connectionLabel }
    $accountText = $script:Window.FindName('AccountText')
    $accountText.Text = if ($Connected) { $Account } else { '' }
    $accountText.ToolTip = $accountText.Text
    $tenantText = $script:Window.FindName('TenantText')
    $tenantText.Text = if ($Connected) { $Tenant } else { '' }
    $tenantText.ToolTip = $tenantText.Text
    $statusLight = $script:Window.FindName('ConnectionStatusLight')
    $statusLight.SetResourceReference([Windows.Shapes.Shape]::FillProperty, $(if ($Connected) { 'SuccessBrush' } else { 'ErrorBrush' }))
    $statusLight.ToolTip = $connectionText.ToolTip
    $script:Window.FindName('DisconnectButton').IsEnabled = -not $script:Job -and [bool]$script:ConnectedTenant
}

function Initialize-EntraConnection {
    $page = $script:Pages[$script:PagesHost.SelectedIndex]
    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
        $context = Get-MgContext -ErrorAction Stop
        if ($context) {
            Set-EntraConnectionState -Connected $true -Account $context.Account -Tenant $context.TenantId
        } else { Set-EntraConnectionState -Connected $false }
        if ($context -and $script:RequestedTenant -and $context.TenantId -ne $script:RequestedTenant) {
            Write-EntraActivity $page $script:Window.FindName('ConnectionText').ToolTip 'WARN'
        }
    } catch {
        $message = "Unable to check the Microsoft Graph connection: $($_.Exception.Message)"
        Set-EntraConnectionState -Connected $false -StatusText 'Connection check failed' -Detail $message
        Write-EntraActivity $page $message 'ERROR'
    }
}

$workerPath = Join-Path $PSScriptRoot 'Shared\Invoke-EntraToolboxOperation.ps1'
if (-not (Test-Path -LiteralPath $workerPath -PathType Leaf)) { throw "Shared worker script not found: $workerPath" }
$script:WorkerSource = Get-Content -LiteralPath $workerPath -Raw

function Start-EntraOperation {
    param($Page, [string]$Kind = 'Run', [hashtable]$PreparedInvocation)
    if ($script:Job) { return }
    try {
        $invocation = @{ Parameters=@{}; Summary=''; Preview=$false }
        if ($Kind -eq 'Run') {
            if ($PreparedInvocation -and $Page.Id -eq 'DeviceCleanup') { $invocation = $PreparedInvocation }
            else { $invocation = Get-EntraInvocation $Page }
        }
        else {
            if ($Page.Id -eq 'Members') { $invocation.Parameters.MemberType = [string]$Page.Controls.MemberType.SelectedItem }
            if ($Page.Id -eq 'PrimaryPC') {
                $invocation.Parameters.DeviceSource = if ($Page.Controls.DeviceSource.SelectedItem -eq 'Sign-in logs only') { 'SignInLogs' } else { 'Entra' }
                $invocation.Parameters.SignInFallback = [bool]$Page.Controls.SignInFallback.IsChecked
                $invocation.Parameters.OnlyIntuneDevices = [bool]$Page.Controls.OnlyIntuneDevices.IsChecked
            }
        }
        $requirements = Get-EntraRequirements -Page $Page -Parameters $invocation.Parameters
        if ($Kind -eq 'Disconnect') { $requirements = @{ Modules=@('Microsoft.Graph.Authentication'); Scopes=@() } }
        foreach ($module in $requirements.Modules) {
            if (-not (Get-Module -ListAvailable -Name $module)) { throw "Missing module. Run: Install-Module $module -Scope CurrentUser" }
        }
        if ($Kind -eq 'Run') {
            if (-not (Test-Path -LiteralPath $Page.Path -PathType Leaf)) { throw "Script not found: $($Page.Path)" }
            if ($invocation.Summary -and -not $invocation.Preview) {
                $message = Get-EntraConfirmationMessage -Page $Page -Invocation $invocation
                if (-not $script:ConnectedTenant) { throw 'Sign in before starting a write operation so the target tenant can be confirmed.' }
                if ([Windows.MessageBox]::Show($script:Window,$message,'Confirm Entra changes','YesNo','Warning') -ne 'Yes') { return }
            }
            if ($invocation.Parameters['OutputPath'] -and (Test-Path -LiteralPath $invocation.Parameters['OutputPath'])) {
                if ([Windows.MessageBox]::Show($script:Window,'Replace the existing output CSV?','Confirm overwrite','YesNo','Question') -ne 'Yes') { return }
            }
        }
        if (-not $script:Runspace) {
            $script:Runspace = [Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
            $script:Runspace.ApartmentState = 'STA'
            # A reused STA worker thread can deadlock Runspace.Dispose() on PowerShell 7 when the window closes.
            $script:Runspace.ThreadOptions = 'UseNewThread'
            $script:Runspace.Open()
        }
        $powershell = [Management.Automation.PowerShell]::Create()
        $powershell.Runspace = $script:Runspace
        $queue = [Collections.Concurrent.ConcurrentQueue[object]]::new()
        $state = [hashtable]::Synchronized(@{ Fatal=''; Account=''; Tenant=''; Connected=$null })
        [void]$powershell.AddScript($script:WorkerSource).AddArgument($Page.Path).AddArgument($invocation.Parameters).AddArgument($requirements.Modules).AddArgument($requirements.Scopes).AddArgument($Kind).AddArgument($script:ConnectedTenant).AddArgument($script:RequestedTenant).AddArgument($queue).AddArgument($state)
        if ($Kind -eq 'Run' -and $invocation.Parameters['Action'] -notin @('Delete','Disable')) {
            $Page.Results.Clear(); $Page.Filter.Text = ''; $Page.LastOutput = $null
            $Page.OpenOutput.IsEnabled = $false
            Update-EntraResults $Page
        }
        $handle = $powershell.BeginInvoke()
        $script:Job = @{ PowerShell=$powershell; Handle=$handle; Queue=$queue; State=$state; Page=$Page; Kind=$Kind; Parameters=$invocation.Parameters; Index=@{ Information=0; Warning=0; Error=0; Verbose=0; Progress=0 }; Counts=@{ Processed=0; Success=0; Failed=0; Skipped=0; Preview=0; Other=0 }; ActivityEntries=[Collections.Generic.List[object]]::new(); ExecutionErrors=0; HadWarnings=$false; HadErrors=$false; Cancelled=$false; StopHandle=$null }
        $Page.Progress.IsIndeterminate = $true
        $Page.Status.Text = if ($Kind -eq 'Connect') { 'Signing in...' } elseif ($Kind -eq 'Disconnect') { 'Signing out...' } else { 'Running...' }
        Add-EntraRunSummaryEntry $script:Job 'INFO' "$Kind started."
        Write-EntraActivity $Page "$Kind started."
        Set-EntraBusy $true
        $script:Timer.Start()
    } catch { Write-EntraActivity $Page $_.Exception.Message 'ERROR'; $Page.Status.Text = 'Not started' }
}

function Stop-EntraOperation {
    if ($script:Job -and -not $script:Job.Cancelled) {
        $script:Job.Cancelled = $true
        $script:Job.Page.Status.Text = 'Stopping; completed changes are not undone...'
        $script:Job.StopHandle = $script:Job.PowerShell.BeginStop($null,$null)
    }
}

function Receive-EntraOperation {
    $job = $script:Job
    if (-not $job) { return }
    foreach ($streamName in @('Information','Warning','Error','Verbose','Progress')) {
        $stream = $job.PowerShell.Streams.$streamName
        $limit = [Math]::Min($stream.Count, $job.Index[$streamName] + 150)
        while ($job.Index[$streamName] -lt $limit) {
            $record = $stream[$job.Index[$streamName]]
            $job.Index[$streamName]++
            if ($streamName -eq 'Progress') { $job.Page.Status.Text = "$($record.Activity): $($record.StatusDescription)"; continue }
            if ($streamName -eq 'Error') { $job.ExecutionErrors++ }
            $level = switch ($streamName) { 'Error' { 'ERROR'; $job.HadErrors=$true } 'Warning' { 'WARN'; $job.HadWarnings=$true } 'Verbose' { 'DETAIL' } default { 'INFO' } }
            Add-EntraRunSummaryEntry $job $level ([string]$record)
            Write-EntraActivity $job.Page ([string]$record) $level
        }
    }
    $result = $null
    $received = 0
    while ($received -lt 150 -and $job.Queue.TryDequeue([ref]$result)) {
        $received++
        if ($job.Kind -eq 'Run') {
            $job.Counts.Processed++
            $job.Counts[(Get-EntraOutcomeCategory $result)]++
        }
        if ($job.Page.Id -eq 'DeviceCleanup' -and $job.Parameters['Action'] -in @('Delete','Disable')) {
            for ($index = 0; $index -lt $job.Page.Results.Count; $index++) {
                if ($job.Page.Results[$index].ObjectId -eq $result.ObjectId) { $job.Page.Results[$index] = $result; break }
            }
        } else { [void]$job.Page.Results.Add($result) }
        $outcome = @($result.PSObject.Properties | Where-Object { $_.Name -in @('Status','Membership','ResolutionStatus','LookupError') } | ForEach-Object { [string]$_.Value }) -join ' '
        if ($outcome -match '(?i)fail|not found|missing') {
            $job.HadWarnings=$true
            $warningMessage = Format-EntraResultWarning $result
            Add-EntraRunSummaryEntry $job 'WARN' $warningMessage
            Write-EntraActivity $job.Page $warningMessage 'WARN'
        }
    }
    if ($received) { Update-EntraResults $job.Page }
    if (-not $job.Handle.IsCompleted -or -not $job.Queue.IsEmpty) { return }
    foreach ($streamName in @('Information','Warning','Error','Verbose','Progress')) { if ($job.Index[$streamName] -lt $job.PowerShell.Streams.$streamName.Count) { return } }
    $script:Timer.Stop()
    try { [void]$job.PowerShell.EndInvoke($job.Handle) } catch { if (-not $job.Cancelled) { $job.ExecutionErrors++; $job.HadErrors=$true; Add-EntraRunSummaryEntry $job 'ERROR' $_.Exception.Message; Write-EntraActivity $job.Page $_.Exception.Message 'ERROR' } }
    if ($job.StopHandle) { try { $job.PowerShell.EndStop($job.StopHandle) } catch { } }
    $job.PowerShell.Dispose()
    if ($null -ne $job.State.Connected) {
        Set-EntraConnectionState -Connected $job.State.Connected -Account $job.State.Account -Tenant $job.State.Tenant
    }
    $status = if ($job.Cancelled) { 'Stopped; partial changes may remain' } elseif ($job.State.Fatal -or $job.HadErrors) { 'Finished with errors' } elseif ($job.HadWarnings) { 'Finished with warnings' } else { 'Completed' }
    $level = if ($job.State.Fatal -or $job.HadErrors) { 'ERROR' } elseif ($job.HadWarnings -or $job.Cancelled) { 'WARN' } else { 'OK' }
    $job.Page.Status.Text = $status
    $job.Page.Progress.IsIndeterminate = $false
    Write-EntraActivity $job.Page $status $level
    if ($job.Kind -eq 'Run' -and $job.Parameters['OutputPath'] -and (Test-Path -LiteralPath $job.Parameters['OutputPath'] -PathType Leaf) -and -not $job.State.Fatal -and -not $job.Cancelled) {
        $job.Page.LastOutput = $job.Parameters['OutputPath']; $job.Page.OpenOutput.IsEnabled = $true
    }
    if ($job.Kind -eq 'Run') {
        $summary = Format-EntraRunSummary $job
        $summaryLevel = if ($job.Counts.Failed -or $job.ExecutionErrors -or $job.Cancelled) { 'WARN' } else { 'INFO' }
        Write-EntraActivity $job.Page $summary $summaryLevel
    }
    $script:Job = $null
    Set-EntraBusy $false
    if ($script:CloseRequested) { $script:Window.Close() }
}

function Add-EntraPage {
    param([string]$Id, [string]$Title, [string]$Description, [string]$RelativePath, [string[]]$Modules, [string[]]$Scopes,
        [Parameter(Mandatory)][ValidateSet('User-Device Affinity','Groups','Users','Devices')][string]$Section)
    $page = @{ Id=$Id; Title=$Title; Description=$Description; Path=(Join-Path $script:EntraRoot $RelativePath); Modules=$Modules; Scopes=$Scopes; Controls=@{}; Fields=@{}; Defaults=@{}; Results=[Collections.Generic.List[object]]::new(); LastOutput=$null; Available=$true }
    $page.CompactForm = $Id -in @('Members','PrimaryPC','Attributes')
    $layout = [Windows.Controls.Grid]::new()
    foreach ($height in @('Auto','Auto','Auto','*','Auto')) { $row = [Windows.Controls.RowDefinition]::new(); $row.Height = [Windows.GridLengthConverter]::new().ConvertFromString($height); [void]$layout.RowDefinitions.Add($row) }
    $formScroll = [Windows.Controls.ScrollViewer]::new(); $formScroll.MaxHeight = 285; $formScroll.VerticalScrollBarVisibility = 'Auto'
    $page.FormScroll = $formScroll
    $page.Layout = $layout
    if ($page.CompactForm) {
        $page.Form = [Windows.Controls.Grid]::new()
        [void]$page.Form.ColumnDefinitions.Add([Windows.Controls.ColumnDefinition]::new())
        [void]$page.Form.ColumnDefinitions.Add([Windows.Controls.ColumnDefinition]::new())
        $formScroll.MaxHeight = [double]::PositiveInfinity
        $page.Form.Tag = $page
        $page.Form.Add_SizeChanged({ param($sender,$eventArgs) Update-EntraLayout $sender.Tag })
    } else { $page.Form = [Windows.Controls.Primitives.UniformGrid]::new(); $page.Form.Columns = 2 }
    $formScroll.Content = $page.Form; [void]$layout.Children.Add($formScroll)
    $actions = [Windows.Controls.StackPanel]::new(); $actions.Orientation = 'Horizontal'; $actions.Margin = [Windows.Thickness]::new(0,5,0,12)
    if ($page.CompactForm) { $actions.Margin = [Windows.Thickness]::new(0,5,0,4) }
    $page.Run = New-EntraButton -Text 'Run' -Tip "Run $Title" -Primary
    $page.Stop = New-EntraButton -Tip 'Stop operation; completed changes are not undone' -Glyph ([string][char]0xE71A); $page.Stop.IsEnabled = $false
    $page.Reset = New-EntraButton -Tip 'Reset parameters and results' -Glyph ([string][char]0xE72C)
    $page.OpenOutput = New-EntraButton -Tip 'Open generated CSV' -Glyph ([string][char]0xE8A5); $page.OpenOutput.IsEnabled = $false
    foreach ($button in @($page.Run,$page.Stop,$page.Reset,$page.OpenOutput)) { [void]$actions.Children.Add($button) }
    $page.Status = [Windows.Controls.TextBlock]::new(); $page.Status.Text = 'Ready'; $page.Status.VerticalAlignment = 'Center'; $page.Status.TextWrapping = 'Wrap'; $page.Status.MaxWidth = 340
    [void]$actions.Children.Add($page.Status)
    [Windows.Controls.Grid]::SetRow($actions,1); [void]$layout.Children.Add($actions)
    $resultsBar = [Windows.Controls.Grid]::new(); $resultsBar.Margin = [Windows.Thickness]::new(0,0,0,8)
    foreach ($width in @('*','Auto','Auto')) { $column = [Windows.Controls.ColumnDefinition]::new(); $column.Width = [Windows.GridLengthConverter]::new().ConvertFromString($width); [void]$resultsBar.ColumnDefinitions.Add($column) }
    $page.Filter = [Windows.Controls.TextBox]::new(); $page.Filter.ToolTip = 'Filter result rows'; $page.Filter.MaxWidth = 350; $page.Filter.HorizontalAlignment = 'Left'; $page.Filter.MinWidth = 200
    $page.Count = [Windows.Controls.TextBlock]::new(); $page.Count.VerticalAlignment = 'Center'; $page.Count.Margin = [Windows.Thickness]::new(12,0,12,0)
    $page.Export = New-EntraButton -Tip 'Export result rows to CSV' -Glyph ([string][char]0xE74E); $page.Export.IsEnabled = $false
    [Windows.Controls.Grid]::SetColumn($page.Count,1); [Windows.Controls.Grid]::SetColumn($page.Export,2)
    foreach ($control in @($page.Filter,$page.Count,$page.Export)) { [void]$resultsBar.Children.Add($control) }
    [Windows.Controls.Grid]::SetRow($resultsBar,2); [void]$layout.Children.Add($resultsBar)
    $page.Table = [Windows.Controls.DataGrid]::new(); $page.Table.MinHeight = 100
    if ($page.CompactForm) { $page.Table.MinHeight = 80 }
    $page.Table.Add_AutoGeneratingColumn({
        param($sender,$eventArgs)
        $eventArgs.Column.MaxWidth = 420; $eventArgs.Column.MinWidth = 110
        if ($eventArgs.Column -is [Windows.Controls.DataGridTextColumn]) {
            $textStyle = [Windows.Style]::new([Windows.Controls.TextBlock])
            [void]$textStyle.Setters.Add([Windows.Setter]::new([Windows.FrameworkElement]::MarginProperty,[Windows.Thickness]::new(7,5,7,5)))
            [void]$textStyle.Setters.Add([Windows.Setter]::new([Windows.Controls.TextBlock]::TextTrimmingProperty,[Windows.TextTrimming]::CharacterEllipsis))
            [void]$textStyle.Setters.Add([Windows.Setter]::new([Windows.FrameworkElement]::ToolTipProperty,[Windows.Data.Binding]::new($eventArgs.PropertyName)))
            $eventArgs.Column.ElementStyle = $textStyle
        }
    })
    [Windows.Controls.Grid]::SetRow($page.Table,3); [void]$layout.Children.Add($page.Table)
    $page.Activity = [Windows.Controls.Expander]::new(); $page.Activity.Margin = [Windows.Thickness]::new(0,10,0,0); $page.Activity.SetResourceReference([Windows.Controls.Control]::ForegroundProperty,'TextBrush')
    $header = [Windows.Controls.StackPanel]::new(); $header.Orientation = 'Horizontal'
    $activityLabel = [Windows.Controls.TextBlock]::new(); $activityLabel.Text = 'Activity'; $activityLabel.FontWeight = 'SemiBold'; $activityLabel.Margin = [Windows.Thickness]::new(0,0,12,0)
    $page.ActivitySummary = [Windows.Controls.TextBlock]::new(); $page.ActivitySummary.Text = 'Ready'; $page.ActivitySummary.MaxWidth = 540; $page.ActivitySummary.TextTrimming = 'CharacterEllipsis'
    [void]$header.Children.Add($activityLabel); [void]$header.Children.Add($page.ActivitySummary); $page.Activity.Header = $header
    $logLayout = [Windows.Controls.DockPanel]::new()
    $logActions = [Windows.Controls.StackPanel]::new(); $logActions.Orientation = 'Horizontal'; $logActions.Margin = [Windows.Thickness]::new(0,6,0,6)
    if ($page.CompactForm) { $logActions.Margin = [Windows.Thickness]::new(0,3,0,3) }
    $clear = New-EntraButton -Tip 'Clear activity log' -Glyph ([string][char]0xE74D)
    $copy = New-EntraButton -Tip 'Copy activity log' -Glyph ([string][char]0xE8C8)
    $save = New-EntraButton -Tip 'Save activity log' -Glyph ([string][char]0xE74E)
    foreach ($button in @($clear,$copy,$save)) { [void]$logActions.Children.Add($button) }
    [Windows.Controls.DockPanel]::SetDock($logActions,'Top'); [void]$logLayout.Children.Add($logActions)
    $page.Log = [Windows.Controls.RichTextBox]::new(); $page.Log.IsReadOnly = $true; $page.Log.Height = 150; $page.Log.VerticalScrollBarVisibility = 'Auto'; $page.Log.FontFamily = [Windows.Media.FontFamily]::new('Consolas'); $page.Log.FontSize = 11
    $page.Log.SetResourceReference([Windows.Controls.Control]::BackgroundProperty,'InputBgBrush'); $page.Log.SetResourceReference([Windows.Controls.Control]::ForegroundProperty,'TextBrush'); $page.Log.SetResourceReference([Windows.Controls.Control]::BorderBrushProperty,'BorderBrush2')
    [void]$logLayout.Children.Add($page.Log); $page.Activity.Content = $logLayout
    $page.Progress = [Windows.Controls.ProgressBar]::new(); $page.Progress.Height = 3; $page.Progress.BorderThickness = [Windows.Thickness]::new(0); $page.Progress.Margin = [Windows.Thickness]::new(0,8,0,0); $page.Progress.SetResourceReference([Windows.Controls.Control]::ForegroundProperty,'AccentBrush'); $page.Progress.SetResourceReference([Windows.Controls.Control]::BackgroundProperty,'SurfaceBrush')
    $footer = [Windows.Controls.StackPanel]::new(); [void]$footer.Children.Add($page.Activity); [void]$footer.Children.Add($page.Progress)
    $page.Footer = $footer; $page.Actions = $actions; $page.ResultsBar = $resultsBar
    $layout.Tag = $page; $footer.Tag = $page
    $layout.Add_SizeChanged({ param($sender,$eventArgs) Update-EntraLayout $sender.Tag })
    $footer.Add_SizeChanged({ param($sender,$eventArgs) Update-EntraLayout $sender.Tag })
    [Windows.Controls.Grid]::SetRow($footer,4); [void]$layout.Children.Add($footer)
    $tab = [Windows.Controls.TabItem]::new(); $tab.Content = $layout; [void]$script:PagesHost.Items.Add($tab)
    $nav = [Windows.Controls.RadioButton]::new(); $nav.Content = $Title; $nav.GroupName = 'EntraNavigation'; $nav.Style = $script:Window.FindResource('NavItem'); $nav.Tag = $script:Pages.Count
    $page.Navigation = $nav
    $nav.Add_Checked({ param($sender,$eventArgs) $script:PagesHost.SelectedIndex = [int]$sender.Tag })
    $page.Section = $Section
    [void]$script:NavSections[$Section].Children.Add($nav)
    foreach ($control in @($page.Run,$page.Filter,$page.OpenOutput,$page.Reset,$page.Export,$clear,$copy,$save)) { $control.Tag = $page }
    $page.Run.Add_Click({ param($sender,$eventArgs) Start-EntraOperation $sender.Tag })
    $page.Stop.Add_Click({ Stop-EntraOperation })
    $page.Filter.Add_TextChanged({ param($sender,$eventArgs) Update-EntraResults $sender.Tag })
    $page.OpenOutput.Add_Click({ param($sender,$eventArgs) $page=$sender.Tag; try { if ($page.LastOutput -and (Test-Path -LiteralPath $page.LastOutput -PathType Leaf)) { Start-Process -FilePath $page.LastOutput -ErrorAction Stop } } catch { Write-EntraActivity $page $_.Exception.Message 'ERROR' } })
    $page.Reset.Add_Click({
        param($sender,$eventArgs)
        $page = $sender.Tag
        foreach ($key in $page.Controls.Keys) {
            $control = $page.Controls[$key]
            switch ($page.Defaults[$key].Kind) {
                'Choice' { $control.SelectedIndex=0 }
                'Check' { $control.IsChecked=$false }
                'CheckList' { foreach ($check in $control.Children) { $check.IsChecked=$false }; $control.Tag.Toggle.IsChecked=$false }
                default { $control.Text=$page.Defaults[$key].Value }
            }
        }
        $page.Results.Clear(); $page.Filter.Text=''; $page.LastOutput=$null; $page.OpenOutput.IsEnabled=$false; $page.Status.Text='Ready'
        if ($page.Id -eq 'DeviceCleanup') { $page.Controls.Preview.IsChecked = $true }
        Update-EntraForm $page; Update-EntraResults $page
    })
    $page.Export.Add_Click({
        param($sender,$eventArgs)
        $page = $sender.Tag
        $dialog = [Microsoft.Win32.SaveFileDialog]::new(); $dialog.Filter = 'CSV files (*.csv)|*.csv'; $dialog.FileName = "$($page.Id)-Results.csv"
        if ($dialog.ShowDialog($script:Window)) {
            try { Export-EntraResults $page $dialog.FileName; Write-EntraActivity $page "Results exported: $($dialog.FileName)" 'OK' }
            catch { Write-EntraActivity $page $_.Exception.Message 'ERROR' }
        }
    })
    $clear.Add_Click({ param($sender,$eventArgs) $page=$sender.Tag; $page.Log.Document.Blocks.Clear(); $page.ActivitySummary.Text='Log cleared' })
    $copy.Add_Click({ param($sender,$eventArgs) $page=$sender.Tag; try { [Windows.Clipboard]::SetText([Windows.Documents.TextRange]::new($page.Log.Document.ContentStart,$page.Log.Document.ContentEnd).Text) } catch { Write-EntraActivity $page $_.Exception.Message 'ERROR' } })
    $save.Add_Click({
        param($sender,$eventArgs)
        $page = $sender.Tag
        $dialog = [Microsoft.Win32.SaveFileDialog]::new(); $dialog.Filter='Text files (*.txt)|*.txt'; $dialog.FileName="$($page.Id)-Activity.txt"
        if ($dialog.ShowDialog($script:Window)) {
            try { [Windows.Documents.TextRange]::new($page.Log.Document.ContentStart,$page.Log.Document.ContentEnd).Text | Set-Content -LiteralPath $dialog.FileName -Encoding $script:Utf8BomEncoding -ErrorAction Stop }
            catch { Write-EntraActivity $page $_.Exception.Message 'ERROR' }
        }
    })
    [void]$script:Pages.Add($page)
    return $page
}

$members = Add-EntraPage -Id 'Members' -Section 'Groups' -Title 'Add Group Members' -Description 'Groups, users and devices' -RelativePath 'Add-MembersToAADGroup\Add-MembersToEntraGroup.ps1' -Modules @('Microsoft.Graph.Groups') -Scopes @('GroupMember.ReadWrite.All','Group.Read.All','User.Read.All','Device.Read.All')
Add-EntraField $members 'TargetGroupName' 'Target group name'
Add-EntraField $members 'MemberType' 'Member type' 'Choice' @('Groups','Users','Devices')
Add-EntraField $members 'Source' 'Member source' 'Choice' @('Names','Text file')
Add-EntraField $members 'WhatIf' 'WhatIf preview (no membership changes)' 'Check'
Add-EntraField $members 'MemberNames' 'Member names / UPNs' 'Lines' -Tip 'One display name or UPN per line.'
Add-EntraField $members 'MemberFile' 'Member text file' 'File'

$groups = Add-EntraPage -Id 'Groups' -Section 'Groups' -Title 'Create Security Groups' -Description 'CSV group creation and nested membership' -RelativePath 'Create-AADGroups\Create-EntraSecurityGroupsFromCsv.ps1' -Modules @('Microsoft.Graph.Groups') -Scopes @('Group.ReadWrite.All','Directory.ReadWrite.All')
Add-EntraField $groups 'CsvPath' 'Security groups CSV' 'File' -Tip 'Semicolon-delimited: GroupName;Description. Optional: MemberGroup, MemberGroupName, GroupMember or Member.'
$previewButton = New-EntraButton -Text 'Preview CSV' -Tip 'Preview security groups CSV'
[void]$groups.Form.Children.Add($previewButton)
$previewButton.Add_Click({
    try {
        $rows = @(Read-EntraCsv -Path $groups.Controls.CsvPath.Text.Trim() -Headers @('GroupName','Description') -Delimiter ';')
        $groups.Results.Clear(); foreach ($row in $rows) { [void]$groups.Results.Add($row) }; Update-EntraResults $groups
        $groups.Status.Text = "$($rows.Count) CSV rows previewed"
    } catch { Write-EntraActivity $groups $_.Exception.Message 'ERROR' }
})

$desktop = [Environment]::GetFolderPath('Desktop')
if (-not $desktop) { $desktop = $HOME }
$primary = Add-EntraPage -Id 'PrimaryPC' -Section 'User-Device Affinity' -Title 'Find Primary PCs' -Description 'Find Windows PCs through Entra associations or successful Windows Sign In events' -RelativePath 'Get-PrimaryPCFromEntraUsers\Get-PrimaryPCFromEntraUsers.ps1' -Modules @('Microsoft.Graph.Groups','Microsoft.Graph.Users','Microsoft.Graph.Identity.DirectoryManagement') -Scopes @('User.Read.All','GroupMember.Read.All','Device.Read.All')
Add-EntraField $primary 'Source' 'User source' 'Choice' @('Choose a source','Group name','Group ID','User CSV','User UPNs')
Add-EntraField $primary 'GroupName' 'Group name'
Add-EntraField $primary 'GroupId' 'Group ID' -Tip 'Entra group object ID (GUID).'
Add-EntraField $primary 'CsvPath' 'User CSV' 'File' -Tip 'CSV with UserPrincipalName column.'
Add-EntraField $primary 'UserPrincipalName' 'User UPNs' 'Lines' -Tip 'One UPN per line.'
Add-EntraField $primary 'DeviceSource' 'Device discovery' 'Choice' @('Entra owned/registered devices','Sign-in logs only') -Tip 'Sign-in logs only bypasses owned/registered devices and counts successful Windows Sign In events. Devices must still resolve in Entra for Windows validation and export.'
Add-EntraField $primary 'DeviceMode' 'Device selection' 'Choice' @('Latest','All') -Tip 'Entra: Latest active Intune-managed PC or All associated Windows devices. Sign-in logs: Most returns all tied PCs; All returns every eligible PC, ordered by successful event count then newest sign-in.'
Add-EntraField $primary 'SignInFallback' 'Use Windows Sign In fallback' 'Check' -Tip 'Requires AuditLog.Read.All and Entra ID P1/P2 sign-in logs.'
Add-EntraField $primary 'SignInDays' 'Sign-in lookback days' -Default '30' -Tip 'Requires AuditLog.Read.All, Microsoft.Graph.Reports and applicable Entra ID P1/P2 licensing. Available history depends on tenant retention and access.'
Add-EntraField $primary 'OnlyIntuneDevices' 'Only devices in Intune' 'Check' -Tip 'Confirm device presence in Intune inventory. Requires DeviceManagementManagedDevices.Read.All.'
Add-EntraField $primary 'Verbose' 'Verbose logging' 'Check' -Tip 'Include per-device lookup details in the activity log.'
Add-EntraField $primary 'OutputPath' 'Entra import CSV output' 'Save' -Default (Join-Path $desktop 'PrimaryPCs.csv')

$lastUser = Add-EntraPage -Id 'LastUser' -Section 'User-Device Affinity' -Title 'Device Last User(s)' -Description 'Intune managed devices and selected logged-on user details' -RelativePath 'Get-DevicesLastUser.ps1' -Scopes @('DeviceManagementManagedDevices.Read.All','DeviceManagementManagedDevices.ReadWrite.All','DeviceManagementConfiguration.Read.All','DeviceManagementConfiguration.ReadWrite.All','User.Read.All')
Add-EntraField $lastUser 'Source' 'Device source' 'Choice' @('Choose a source','Device names','Entra group','CSV file','All devices')
Add-EntraField $lastUser 'UserMode' 'User selection' 'Choice' @('Latest logged-on user','Most logged-on user(s)','All logged-on users') -Tip 'Latest: newest logon. Most: count userId occurrences in usersLoggedOn and return all tied winners. All: each distinct logged-on user, newest first. One row per device/user; primary-user fallback when no logged-on userId is available.'
Add-EntraField $lastUser 'ExtensionAttributeNumbers' 'User extension attribute columns (optional)' 'CheckList' @(1..15 | ForEach-Object { "extensionAttribute$_" }) -Default 'Choose attributes' -Tip 'Select columns for the next run and CSV exports. Values come from each resolved user; unset attributes are blank.'
Add-EntraField $lastUser 'OutputPath' 'Report CSV output' 'Save' -Default (Join-Path $desktop "intune_device_last_logged_on_user_$(Get-Date -Format 'yyyy-MM-dd').csv")
Add-EntraField $lastUser 'DeviceName' 'Device names' 'Lines' -Tip 'One Intune managed device name per line.'
Add-EntraField $lastUser 'GroupNameOrId' 'Entra group name or ID' -Tip 'Must resolve to exactly one Entra group. Only direct device members are included.'
Add-EntraField $lastUser 'CsvPath' 'Device CSV' 'File' -Tip 'DeviceId, ManagedDeviceId or Id (Intune or Entra device ID); or DeviceName, ManagedDeviceName, Name or displayName. ID takes precedence. Entra group member exports work as-is.'

$attributes = Add-EntraPage -Id 'Attributes' -Section 'Users' -Title 'User Extension Attributes' -Description 'Cloud-managed user extension attributes' -RelativePath 'ExtensionAttributes\Set-ExtensionAttribute.ps1' -Scopes @('User.ReadWrite.All')
Add-EntraField $attributes 'Source' 'User source' 'Choice' @('UPN list','CSV file')
Add-EntraField $attributes 'AttributeNumber' 'Extension attribute' 'Choice' @('1','2','3','4','5','6','7','8','9','10','11','12','13','14','15')
Add-EntraField $attributes 'UserPrincipalName' 'User UPNs' 'Lines' -Tip 'One UPN per line. Synced attributes must be changed in on-premises AD.'
Add-EntraField $attributes 'CsvPath' 'User CSV' 'File' -Tip 'CSV with UserPrincipalName column.'
Add-EntraField $attributes 'Action' 'Attribute action' 'Choice' @('Set','Clear')
Add-EntraField $attributes 'Value' 'Attribute value'
Add-EntraField $attributes 'WhatIf' 'WhatIf preview (no attribute changes)' 'Check'

$cleanup = Add-EntraPage -Id 'DeviceCleanup' -Section 'Devices' -Title 'Device Cleanup' -Description 'Device inventory and selected cloud-device cleanup' -RelativePath 'Invoke-EntraDeviceCleanup.ps1' -Scopes @('Device.Read.All','DeviceManagementManagedDevices.Read.All','DeviceManagementServiceConfig.Read.All')
$cleanup.Run.Content = 'Load devices'
Add-EntraField $cleanup 'Source' 'Device source' 'Choice' @('CSV file','Entra lookup','Device names') -Tip 'CSV file: devices listed in a CSV. Entra lookup: inactive Windows devices. Device names: exact names typed below.'
Add-EntraField $cleanup 'CsvPath' 'Device CSV' 'File' -Tip 'Entra group/device export or CSV with DeviceName/displayName, deviceId and/or id/ObjectId.'
Add-EntraField $cleanup 'InactivityDays' 'Inactive for at least (days)' -Default '90' -Tip 'Lists Windows devices whose latest Entra sign-in or Intune sync is older than this, plus devices with no activity date.'
Add-EntraField $cleanup 'DeviceNames' 'Device names' 'Lines' -Tip 'One exact device name per line. Every Entra device with a matching name is listed.'
Add-EntraField $cleanup 'NameFilter' 'Device name contains'
Add-EntraField $cleanup 'JoinFilter' 'Join type' 'Choice' @('All','Microsoft Entra joined','Microsoft Entra registered','Microsoft Entra hybrid joined','Unknown')
Add-EntraField $cleanup 'ManagedFilter' 'Intune managed' 'Choice' @('All','Yes','No')
Add-EntraField $cleanup 'DateFrom' 'Activity from (UTC)' -Tip 'Inclusive date, yyyy-MM-dd.'
Add-EntraField $cleanup 'DateTo' 'Activity to (UTC)' -Tip 'Inclusive date, yyyy-MM-dd.'
Add-EntraField $cleanup 'DateMode' 'Activity dates' 'Choice' @('All dates','Known dates','Unknown dates')
Add-EntraField $cleanup 'Preview' 'Preview only (no changes)' 'Check'
Add-EntraField $cleanup 'Acknowledge' 'I acknowledge explicit deletion of remaining cloud Entra objects' 'Check'
$cleanup.Controls.Preview.IsChecked = $true
$cleanup.Table.AutoGenerateColumns = $false
$cleanup.Table.SelectionMode = 'Extended'
$cleanup.Table.SelectionUnit = 'FullRow'
foreach ($name in @('Device Name','Join Type','Enabled','Intune managed','Autopilot','Activity date','Device ID','Status')) {
    $column = [Windows.Controls.DataGridTextColumn]::new(); $column.Header = $name
    $binding = [Windows.Data.Binding]::new($name)
    if ($name -eq 'Activity date') { $binding.StringFormat = 'yyyy-MM-dd HH:mm:ss'; $column.Header = 'Activity date (UTC)' }
    $column.Binding = $binding; $column.SortMemberPath = $name; $column.MinWidth = $(if ($name -in @('Enabled','Autopilot')) { 80 } else { 110 }); $column.MaxWidth = 420
    $column.Width = [Windows.Controls.DataGridLength]::new($(switch ($name) { 'Device Name' { 180 } 'Join Type' { 210 } 'Enabled' { 85 } 'Autopilot' { 90 } 'Activity date' { 180 } 'Device ID' { 270 } 'Status' { 260 } default { 125 } }))
    $textStyle = [Windows.Style]::new([Windows.Controls.TextBlock])
    [void]$textStyle.Setters.Add([Windows.Setter]::new([Windows.FrameworkElement]::MarginProperty,[Windows.Thickness]::new(7,5,7,5)))
    [void]$textStyle.Setters.Add([Windows.Setter]::new([Windows.Controls.TextBlock]::TextTrimmingProperty,[Windows.TextTrimming]::CharacterEllipsis))
    [void]$textStyle.Setters.Add([Windows.Setter]::new([Windows.FrameworkElement]::ToolTipProperty,[Windows.Data.Binding]::new($name)))
    $column.ElementStyle = $textStyle
    [void]$cleanup.Table.Columns.Add($column)
}
$cleanup.DeleteSelected = New-EntraButton -Text 'Delete selected' -Tip 'Delete selected eligible cloud devices' -Glyph ([string][char]0xE74D)
$cleanup.DeleteSelected.IsEnabled = $false; $cleanup.DeleteSelected.Tag = $cleanup
[void]$cleanup.Actions.Children.Insert(1,$cleanup.DeleteSelected)
$cleanup.DisableSelected = New-EntraButton -Text 'Disable selected' -Tip 'Disable selected enabled cloud devices in Entra' -Glyph ([string][char]0xE733)
$cleanup.DisableSelected.IsEnabled = $false; $cleanup.DisableSelected.Tag = $cleanup
[void]$cleanup.Actions.Children.Insert(1,$cleanup.DisableSelected)
$cleanup.SelectAll = New-EntraButton -Tip 'Select all visible eligible devices' -Glyph ([string][char]0xE8B3)
$cleanup.ClearSelection = New-EntraButton -Tip 'Clear selection' -Glyph ([string][char]0xE8E6)
foreach ($key in @('SelectAll','ClearSelection')) {
    $cleanup[$key].Tag = $cleanup
    $cleanup[$key].IsEnabled = $false
}
[void]$cleanup.Actions.Children.Insert(1,$cleanup.SelectAll)
[void]$cleanup.Actions.Children.Insert(2,$cleanup.ClearSelection)
$cleanup.SelectAll.Add_Click({
    param($sender,$eventArgs)
    if ($script:Job) { return }
    $page = $sender.Tag
    $eligible = @(Get-EntraCleanupEligibleRows $page)
    $page.Table.SelectedItems.Clear()
    foreach ($row in $eligible) { [void]$page.Table.SelectedItems.Add($row) }
    Update-EntraCleanupSelection $page
})
$cleanup.ClearSelection.Add_Click({
    param($sender,$eventArgs)
    if ($script:Job) { return }
    $sender.Tag.Table.SelectedItems.Clear()
    Update-EntraCleanupSelection $sender.Tag
})
$cleanup.DeleteSelected.Add_Click({
    param($sender,$eventArgs)
    try { Start-EntraOperation -Page $sender.Tag -PreparedInvocation (Get-EntraCleanupDeletionInvocation $sender.Tag) }
    catch { Write-EntraActivity $sender.Tag $_.Exception.Message 'ERROR' }
})
$cleanup.DisableSelected.Add_Click({
    param($sender,$eventArgs)
    try { Start-EntraOperation -Page $sender.Tag -PreparedInvocation (Get-EntraCleanupDisableInvocation $sender.Tag) }
    catch { Write-EntraActivity $sender.Tag $_.Exception.Message 'ERROR' }
})
$cleanup.Table.Tag = $cleanup
$cleanup.Table.Add_SelectionChanged({ param($sender,$eventArgs) Update-EntraCleanupSelection $sender.Tag })
foreach ($key in @('NameFilter','JoinFilter','ManagedFilter','DateFrom','DateTo','DateMode')) {
    $control = $cleanup.Controls[$key]; $control.Tag = $cleanup
    if ($control -is [Windows.Controls.ComboBox]) {
        $control.Add_SelectionChanged({ param($sender,$eventArgs) Update-EntraResults $sender.Tag })
    } else { $control.Add_TextChanged({ param($sender,$eventArgs) Update-EntraResults $sender.Tag }) }
}

foreach ($page in $script:Pages) {
    if ($page.Id -eq 'PrimaryPC') {
        foreach ($key in @('GroupName','GroupId','CsvPath','UserPrincipalName')) {
            $page.Controls[$key].Tag = $page
            $page.Controls[$key].Add_TextChanged({ param($sender,$eventArgs) Update-EntraRunAvailability $sender.Tag })
        }
    }
    foreach ($key in @('Source','Action','DeviceSource','SignInFallback')) {
        if ($page.Controls.ContainsKey($key)) {
            $control = $page.Controls[$key]
            $control.Tag = $page
            $handler = { param($sender,$eventArgs) Update-EntraForm $sender.Tag }
            if ($control -is [Windows.Controls.CheckBox]) { $control.Add_Checked($handler); $control.Add_Unchecked($handler) }
            else { $control.Add_SelectionChanged($handler) }
        }
    }
    $page.Available = Test-Path -LiteralPath $page.Path -PathType Leaf
    Update-EntraRunAvailability $page
    if (-not $page.Available) { Write-EntraActivity $page "Missing script: $($page.Path)" 'ERROR' }
    Update-EntraForm $page; Update-EntraResults $page
}
$script:PagesHost.Add_SelectionChanged({
    $index = $script:PagesHost.SelectedIndex
    if ($index -ge 0 -and $index -lt $script:Pages.Count) {
        $page = $script:Pages[$index]
        $script:Window.FindName('PageTitle').Text = $page.Title
        $script:Window.FindName('PageDescription').Text = $page.Description
        $page.Navigation.IsChecked = $true
    }
})
$startNavigation = @($script:NavSections.Values | ForEach-Object { $_.Children } | Select-Object -First 1)[0]
$script:PagesHost.SelectedIndex = -1
$script:PagesHost.SelectedIndex = [int]$startNavigation.Tag
$startNavigation.IsChecked = $true
$script:Window.FindName('ConnectButton').Add_Click({ Start-EntraOperation -Page $script:Pages[$script:PagesHost.SelectedIndex] -Kind 'Connect' })
$script:Window.FindName('DisconnectButton').Add_Click({ Start-EntraOperation -Page $script:Pages[$script:PagesHost.SelectedIndex] -Kind 'Disconnect' })
$script:Window.FindName('ThemeButton').Add_Click({
    $colors = if ($script:IsDark) {
        @{ BgBrush='#f5f6f8'; SurfaceBrush='#ffffff'; InputBgBrush='#ffffff'; TextBrush='#20252d'; SubTextBrush='#596575'; BorderBrush2='#d5dae1'; SelectionBrush='#fce7ea'; SuccessBrush='#147d55'; WarningBrush='#926200'; ErrorBrush='#b42332' }
    } else {
        @{ BgBrush='#191c22'; SurfaceBrush='#252b34'; InputBgBrush='#20252d'; TextBrush='#e6e9ef'; SubTextBrush='#9ca8b7'; BorderBrush2='#39424e'; SelectionBrush='#49272e'; SuccessBrush='#36c18f'; WarningBrush='#fbbf24'; ErrorBrush='#ff8791' }
    }
    foreach ($key in $colors.Keys) { $script:Window.Resources[$key] = [Windows.Media.SolidColorBrush]::new([Windows.Media.ColorConverter]::ConvertFromString($colors[$key])) }
    $script:IsDark = -not $script:IsDark
})
$script:Timer = [Windows.Threading.DispatcherTimer]::new()
$script:Timer.Interval = [TimeSpan]::FromMilliseconds(200)
$script:Timer.Add_Tick({ Receive-EntraOperation })
$script:Window.Add_Closing({
    param($sender,$eventArgs)
    if ($script:Job) {
        $eventArgs.Cancel = $true
        if ([Windows.MessageBox]::Show($script:Window,'Stop the active operation and close? Completed changes cannot be undone.','Operation running','YesNo','Warning') -eq 'Yes') {
            $script:CloseRequested = $true
            Stop-EntraOperation
        }
    }
})
$script:Window.Add_Closed({
    $script:Timer.Stop()
    if ($script:Runspace) {
        $script:Runspace.Dispose()
        $script:Runspace = $null
    }
})
if ($ValidateOnly) {
    $script:Window.Measure([Windows.Size]::new(1280,900))
    $script:Window.Arrange([Windows.Rect]::new(0,0,1280,900))
    $script:Window.UpdateLayout()
    [pscustomobject]@{ Title=$script:Window.Title; PageCount=$script:Pages.Count; MissingScripts=@($script:Pages | Where-Object { -not $_.Available }).Count; Window=$script:Window; Pages=$script:Pages }
} else {
    Initialize-EntraConnection
    [void]$script:Window.ShowDialog()
}