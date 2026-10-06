<#
.SYNOPSIS
    Meeting Cleanup - window (dot-sourced by MeetingCleanup.psm1).

.DESCRIPTION
    WPF window. With .NET 9 or later (PowerShell 7.5 and later) it uses the Fluent theme of Windows 11:
    light or dark like Windows, rounded controls, with the accent colour of the report (#B11F4B). With
    PowerShell 7.4 (.NET 8) the same window uses the classic WPF controls with the colours of the report.

    Layout: a header; on the left the organizer, the period and the subject, where to search, the action
    and the connection; on the right the meetings found (a box to tick each one) with the copies of the
    meeting selected, and the progress; at the bottom the buttons.

    It runs exactly the same engine as the command line (Find-MclMeetings, Invoke-MclCleanup,
    Export-MclReport): the progress shows the lines of the console. The work happens on the window thread,
    kept responsive between the Graph calls (Invoke-MclGuiPump); Stop ends it at the next call.
    Search is read-only and writes a report; the action button removes or cancels the meetings ticked,
    after a confirmation that says exactly what will happen.

    Closing never needs PowerShell code: Close is the cancel button of the window (Esc too) and the title-bar
    button is native; the only Closing handler is attached while a run is in progress. A PowerShell event
    handler fails ("The pipeline has been stopped") once the command that opened the window is stopped, so
    the window must close without one. Ctrl+C is ignored in the console while it is open.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.2.0
#>

$script:Gui = $null
# Answers given in advance to the questions of the window (lab tests only, see Show-MclGuiQuestion).
$script:GuiAnswers = $null

function Get-MclGuiXaml {
    <# The window. Colours come from the Fluent theme resources (or the classic fallback of Set-MclGuiTheme). #>
    @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="1320" Height="900" MinWidth="1080" MinHeight="680" WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" UseLayoutRounding="True">
  <Window.Resources>
    <Style x:Key="MclCard" TargetType="Border">
      <Setter Property="Background" Value="{DynamicResource CardBackgroundFillColorDefaultBrush}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource CardStrokeColorDefaultBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="8"/>
      <Setter Property="Padding" Value="18,14,18,16"/>
      <Setter Property="Margin" Value="0,0,0,12"/>
    </Style>
    <Style x:Key="MclCardTitle" TargetType="TextBlock">
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,0,0,8"/>
      <Setter Property="Foreground" Value="{DynamicResource TextFillColorPrimaryBrush}"/>
    </Style>
    <Style x:Key="MclLabel" TargetType="TextBlock">
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Foreground" Value="{DynamicResource TextFillColorSecondaryBrush}"/>
      <Setter Property="Margin" Value="0,8,0,4"/>
    </Style>
    <Style x:Key="MclHint" TargetType="TextBlock">
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Foreground" Value="{DynamicResource TextFillColorSecondaryBrush}"/>
    </Style>
    <Style x:Key="MclIcon" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="Segoe Fluent Icons, Segoe MDL2 Assets"/>
    </Style>
  </Window.Resources>

  <Grid x:Name="Root" Background="{DynamicResource ApplicationBackgroundBrush}">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Grid x:Name="Header" Margin="24,18,24,14">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <Border Width="46" Height="46" CornerRadius="10" Background="{DynamicResource MclBrand}" VerticalAlignment="Center">
        <TextBlock Style="{StaticResource MclIcon}" Text="&#xE787;" FontSize="22" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
      </Border>
      <StackPanel Grid.Column="1" Margin="14,0,0,0" VerticalAlignment="Center">
        <TextBlock Text="EXCHANGE ONLINE CALENDAR CLEANUP" FontSize="11" FontWeight="SemiBold" Foreground="{DynamicResource MclBrandText}"/>
        <TextBlock Text="Meeting Cleanup" FontSize="24" FontWeight="SemiBold" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
        <TextBlock FontSize="13" Foreground="{DynamicResource TextFillColorSecondaryBrush}" TextTrimming="CharacterEllipsis"
                   Text="Find the meetings of organizers or rooms in every calendar, then remove them silently, cancel them, or transfer them to a new organizer."/>
      </StackPanel>
      <TextBlock x:Name="Version" Grid.Column="2" FontSize="12" Foreground="{DynamicResource TextFillColorSecondaryBrush}" VerticalAlignment="Top"/>
    </Grid>

    <Grid Grid.Row="1" Margin="24,0,24,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="410"/>
        <ColumnDefinition Width="16"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>
      <ScrollViewer x:Name="SettingsScroll" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Padding="0,0,4,0">
        <StackPanel x:Name="Inputs">
          <Border Style="{StaticResource MclCard}">
            <StackPanel>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="WhoTitle" Text="Organizers" Style="{StaticResource MclCardTitle}"/>
                <Button x:Name="LoadOrganizers" Grid.Column="1" Content="Load a list..." Padding="10,2" Margin="0,-4,0,6"/>
              </Grid>
              <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                <RadioButton x:Name="ModeOrganizers" GroupName="Mode" Content="Meetings of organizers" Margin="0,0,16,0"/>
                <RadioButton x:Name="ModeRooms" GroupName="Mode" Content="Every meeting of rooms"/>
              </StackPanel>
              <TextBox x:Name="Organizer" AcceptsReturn="True" TextWrapping="NoWrap" MinHeight="34" MaxHeight="96" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
              <TextBlock x:Name="OrganizerHint" Style="{StaticResource MclHint}" Margin="0,6,0,0" Text="One per line (or separated by ;): SMTP address (any alias), or the X500 address of a deleted mailbox. A list: text or CSV file."/>
            </StackPanel>
          </Border>

          <Border Style="{StaticResource MclCard}">
            <StackPanel>
              <TextBlock Text="Meetings" Style="{StaticResource MclCardTitle}" Margin="0,0,0,0"/>
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="12"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                  <TextBlock Text="From" Style="{StaticResource MclLabel}"/>
                  <DatePicker x:Name="StartDate" SelectedDateFormat="Short"/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock Text="To (included)" Style="{StaticResource MclLabel}"/>
                  <DatePicker x:Name="EndDate" SelectedDateFormat="Short"/>
                </StackPanel>
              </Grid>
              <TextBlock Text="Subject contains (empty = every meeting of the period)" Style="{StaticResource MclLabel}"/>
              <TextBox x:Name="Subject"/>
              <TextBlock x:Name="PeriodHint" Style="{StaticResource MclHint}" Margin="0,6,0,0" Text="A series is found when one of its occurrences falls in the period, and is handled as a whole."/>
            </StackPanel>
          </Border>

          <Border x:Name="SearchCard" Style="{StaticResource MclCard}">
            <StackPanel>
              <TextBlock Text="Search in" Style="{StaticResource MclCardTitle}"/>
              <CheckBox x:Name="ScopeOrganizer" Content="The organizer's calendar"/>
              <TextBlock Style="{StaticResource MclHint}" Margin="28,0,0,6" Text="When the mailbox still exists."/>
              <CheckBox x:Name="ScopeRooms" Content="Every room mailbox"/>
              <TextBlock Style="{StaticResource MclHint}" Margin="28,0,0,6" Text="Places API, plus the rooms of the configuration."/>
              <CheckBox x:Name="ScopeMailboxes" Content="The mailboxes of a file"/>
              <Grid Margin="28,4,0,6">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBox x:Name="MailboxFile"/>
                <Button x:Name="Browse" Grid.Column="1" Margin="8,0,0,0" Content="Browse..."/>
              </Grid>
              <CheckBox x:Name="ScopeAll" Content="Every mailbox of the tenant"/>
              <TextBlock Style="{StaticResource MclHint}" Margin="28,0,0,6" Text="Complete but long (about 3,000 mailboxes a minute): for a deleted organizer whose meetings have no room."/>
              <TextBlock Style="{StaticResource MclHint}" Margin="0,4,0,0" Text="Each meeting found is then looked up in the calendar of every internal attendee, room and member of an invited group."/>
            </StackPanel>
          </Border>

          <Border Style="{StaticResource MclCard}">
            <StackPanel>
              <TextBlock Text="Action on the meetings ticked" Style="{StaticResource MclCardTitle}"/>
              <RadioButton x:Name="ActionRemove" GroupName="Action" Content="Remove silently"/>
              <TextBlock Style="{StaticResource MclHint}" Margin="28,0,0,8" Text="The copies of the attendees and the rooms are removed, without any message. The meeting stays in the organizer's calendar when the mailbox exists: removing it there always sends a cancellation."/>
              <RadioButton x:Name="ActionCancel" GroupName="Action" Content="Cancel and clean"/>
              <TextBlock Style="{StaticResource MclHint}" Margin="28,0,0,6" Text="The organizer cancels the meeting (message to every attendee, rooms released), then the copies left are removed. A meeting without an organizer copy is removed silently."/>
              <TextBlock Text="Cancellation message" Style="{StaticResource MclLabel}" Margin="28,4,0,4"/>
              <TextBox x:Name="Comment" Margin="28,0,0,8" TextWrapping="Wrap" AcceptsReturn="True" MinHeight="56"/>
              <RadioButton x:Name="ActionTransfer" GroupName="Action" Content="Transfer to a new organizer"/>
              <TextBlock x:Name="TransferHint" Style="{StaticResource MclHint}" Margin="28,0,0,6" Text="Organizer mailbox present: Exchange Online moves the meeting (attendees updated silently). Mailbox gone: the meeting is re-created by the new organizer, who sends one invitation; the old copies are removed silently."/>
              <Grid Margin="28,4,0,0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="8"/>
                  <ColumnDefinition Width="112"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                  <TextBlock Text="New organizer" Style="{StaticResource MclLabel}" Margin="0,0,0,4"/>
                  <TextBox x:Name="NewOrganizer"/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                  <TextBlock Text="Method" Style="{StaticResource MclLabel}" Margin="0,0,0,4"/>
                  <ComboBox x:Name="TransferMethod"/>
                </StackPanel>
              </Grid>
            </StackPanel>
          </Border>

          <Expander x:Name="ConnectionExpander" Header="Connection (Microsoft Graph application)" Margin="0,0,0,12">
            <Border Style="{StaticResource MclCard}" Margin="0,8,0,0">
              <StackPanel>
                <TextBlock Text="Tenant ID or domain" Style="{StaticResource MclLabel}" Margin="0,0,0,4"/>
                <TextBox x:Name="TenantId"/>
                <TextBlock Text="Application (client) ID" Style="{StaticResource MclLabel}"/>
                <TextBox x:Name="AppId"/>
                <TextBlock Text="Sign-in of the application" Style="{StaticResource MclLabel}"/>
                <ComboBox x:Name="AuthMode"/>
                <StackPanel x:Name="ThumbPanel">
                  <TextBlock Text="Certificate thumbprint" Style="{StaticResource MclLabel}"/>
                  <TextBox x:Name="Thumbprint"/>
                </StackPanel>
                <StackPanel x:Name="SecretPanel">
                  <TextBlock Text="Client secret (empty = environment variable; never written)" Style="{StaticResource MclLabel}"/>
                  <PasswordBox x:Name="Secret"/>
                </StackPanel>
                <TextBlock x:Name="ConfigHint" Style="{StaticResource MclHint}" Margin="0,8,0,0"/>
              </StackPanel>
            </Border>
          </Expander>
        </StackPanel>
      </ScrollViewer>

      <Grid Grid.Column="2">
        <Grid.RowDefinitions>
          <RowDefinition Height="*" MinHeight="260"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="210" MinHeight="120"/>
        </Grid.RowDefinitions>
        <Border Style="{StaticResource MclCard}" Margin="0,0,0,6">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="150"/>
            </Grid.RowDefinitions>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock Text="Meetings" Style="{StaticResource MclCardTitle}"/>
              <TextBlock x:Name="Counts" Grid.Column="1" Margin="12,1,0,8" FontSize="12" VerticalAlignment="Top" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
              <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Top">
                <Border x:Name="StatusPill" CornerRadius="10" Padding="10,2" Margin="0,0,10,0" VerticalAlignment="Center">
                  <TextBlock x:Name="Status" FontSize="12" FontWeight="SemiBold"/>
                </Border>
                <Button x:Name="SelectAll" Content="Tick all" Padding="10,3" Margin="0,0,6,0"/>
                <Button x:Name="SelectNone" Content="Untick all" Padding="10,3"/>
              </StackPanel>
            </Grid>
            <DataGrid x:Name="Meetings" Grid.Row="1" AutoGenerateColumns="False" CanUserAddRows="False" CanUserDeleteRows="False" HeadersVisibility="Column"
                      SelectionMode="Single" GridLinesVisibility="None" RowHeaderWidth="0" BorderThickness="0" Background="Transparent" IsReadOnly="True">
              <DataGrid.Columns>
                <DataGridTemplateColumn Header="" Width="44">
                  <DataGridTemplateColumn.CellTemplate>
                    <DataTemplate>
                      <CheckBox IsChecked="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" IsEnabled="{Binding CanTick}" HorizontalAlignment="Center" MinWidth="0" Padding="0"/>
                    </DataTemplate>
                  </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn Header="Start" Binding="{Binding Start}" Width="112"/>
                <DataGridTextColumn Header="Subject" Binding="{Binding Subject}" Width="2*" MinWidth="140"/>
                <DataGridTextColumn Header="Organizer" Binding="{Binding Who}" Width="*" MinWidth="96" Visibility="Collapsed"/>
                <DataGridTextColumn Header="Kind" Binding="{Binding Kind}" Width="62"/>
                <DataGridTextColumn Header="Organizer copy" Binding="{Binding OrganizerCopy}" Width="128"/>
                <DataGridTextColumn Header="Copies" Binding="{Binding CopiesText}" Width="96"/>
                <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="110"/>
              </DataGrid.Columns>
            </DataGrid>
            <TextBlock x:Name="MeetingsEmpty" Grid.Row="1" Margin="4,48,4,0" TextWrapping="Wrap" HorizontalAlignment="Center" FontSize="13" Foreground="{DynamicResource TextFillColorTertiaryBrush}"
                       Text="Type the organizer, choose where to search, then Search. A search changes nothing."/>
            <TextBlock x:Name="CopiesTitle" Grid.Row="2" Margin="0,10,0,6" FontSize="12" FontWeight="SemiBold" TextTrimming="CharacterEllipsis" Foreground="{DynamicResource TextFillColorSecondaryBrush}" Text="Copies of the meeting selected"/>
            <DataGrid x:Name="Copies" Grid.Row="3" AutoGenerateColumns="False" IsReadOnly="True" HeadersVisibility="Column" GridLinesVisibility="None" RowHeaderWidth="0"
                      BorderThickness="0" Background="Transparent" FontSize="12">
              <DataGrid.Columns>
                <DataGridTextColumn Header="Mailbox" Binding="{Binding Mailbox}" Width="2*"/>
                <DataGridTextColumn Header="Role" Binding="{Binding Role}" Width="80"/>
                <DataGridTextColumn Header="Occurrence" Binding="{Binding Occurrence}" Width="112"/>
                <DataGridTextColumn Header="Found by" Binding="{Binding Via}" Width="*"/>
                <DataGridTextColumn Header="Result" Binding="{Binding Result}" Width="105"/>
                <DataGridTextColumn Header="Detail" Binding="{Binding Detail}" Width="2*"/>
              </DataGrid.Columns>
            </DataGrid>
          </Grid>
        </Border>
        <GridSplitter Grid.Row="1" Height="6" HorizontalAlignment="Stretch" Background="Transparent" ResizeDirection="Rows"/>
        <Border Grid.Row="2" Style="{StaticResource MclCard}" Margin="0,6,0,12" Padding="18,10,18,10">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
              </Grid.ColumnDefinitions>
              <TextBlock Text="Progress" Style="{StaticResource MclCardTitle}" Margin="0,0,12,6"/>
              <Grid Grid.Column="1" x:Name="ProgressPanel" Visibility="Collapsed" Margin="0,2,0,6">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="160"/>
                  <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <ProgressBar x:Name="ProgressBar" Height="6" Minimum="0" Maximum="1" VerticalAlignment="Center"/>
                <TextBlock x:Name="ProgressText" Grid.Column="1" Margin="10,0,0,0" FontSize="12" TextTrimming="CharacterEllipsis" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
              </Grid>
            </Grid>
            <ScrollViewer x:Name="LogScroll" Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
              <ItemsControl x:Name="Log" Margin="0,0,12,0">
                <ItemsControl.ItemTemplate>
                  <DataTemplate>
                    <Grid Margin="{Binding Margin}">
                      <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="22"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                      </Grid.ColumnDefinitions>
                      <TextBlock Text="{Binding Glyph}" Foreground="{Binding Brush}" FontFamily="Segoe Fluent Icons, Segoe MDL2 Assets" FontSize="12" Margin="0,3,0,0" VerticalAlignment="Top"/>
                      <TextBlock Grid.Column="1" Text="{Binding Text}" TextWrapping="Wrap" FontSize="{Binding Size}" FontWeight="{Binding Weight}" Foreground="{Binding TextBrush}"/>
                      <TextBlock Grid.Column="2" Text="{Binding Time}" FontSize="11" Margin="10,2,0,0" Foreground="{DynamicResource TextFillColorTertiaryBrush}"/>
                    </Grid>
                  </DataTemplate>
                </ItemsControl.ItemTemplate>
              </ItemsControl>
            </ScrollViewer>
          </Grid>
        </Border>
      </Grid>
    </Grid>

    <Border x:Name="Actions" Grid.Row="2" Padding="24,12" BorderThickness="0,1,0,0"
            BorderBrush="{DynamicResource DividerStrokeColorDefaultBrush}" Background="{DynamicResource LayerFillColorDefaultBrush}">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal">
          <Button x:Name="Search" MinWidth="130" Padding="16,6" Margin="0,0,8,0">
            <StackPanel Orientation="Horizontal">
              <TextBlock Style="{StaticResource MclIcon}" Text="&#xE721;" Margin="0,2,8,0"/>
              <TextBlock Text="Search"/>
            </StackPanel>
          </Button>
          <Button x:Name="Apply" MinWidth="210" Padding="16,6" Margin="0,0,8,0" IsEnabled="False">
            <StackPanel Orientation="Horizontal">
              <TextBlock x:Name="ApplyIcon" Style="{StaticResource MclIcon}" Text="&#xE74D;" Margin="0,2,8,0"/>
              <TextBlock x:Name="ApplyText" Text="Remove the meetings ticked"/>
            </StackPanel>
          </Button>
          <Button x:Name="Stop" Content="Stop" MinWidth="80" IsEnabled="False"/>
        </StackPanel>
        <TextBlock x:Name="Footer" Grid.Column="1" Margin="16,0" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" FontSize="12"
                   Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
        <StackPanel Grid.Column="2" Orientation="Horizontal">
          <Button x:Name="Restore" Margin="0,0,8,0" ToolTip="Put back the copies removed by a Remove run (Recoverable Items)">
            <StackPanel Orientation="Horizontal">
              <TextBlock Style="{StaticResource MclIcon}" Text="&#xE777;" Margin="0,2,8,0"/>
              <TextBlock Text="Restore..."/>
            </StackPanel>
          </Button>
          <Button x:Name="OpenReport" Content="Open the report" Margin="0,0,8,0" IsEnabled="False"/>
          <Button x:Name="OpenFolder" Content="Open the folder" Margin="0,0,8,0" IsEnabled="False"/>
          <Button x:Name="Close" Content="Close" MinWidth="90" IsCancel="True"/>
        </StackPanel>
      </Grid>
    </Border>
  </Grid>
</Window>
'@
}

function New-MclGuiBrush {
    param([Parameter(Mandatory = $true)][string]$Color)
    $brush = [Windows.Media.SolidColorBrush]::new([Windows.Media.ColorConverter]::ConvertFromString($Color))
    $brush.Freeze()
    return $brush
}

function Test-MclGuiDarkMode {
    <# Windows shows the applications in dark mode (AppsUseLightTheme = 0); light when the setting is missing. #>
    $personalize = Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -ErrorAction SilentlyContinue
    return [string](Get-MclProperty $personalize 'AppsUseLightTheme') -eq '0'
}

function Initialize-MclGuiTheme {
    <#
        Loads WPF and applies the theme to the application: Fluent (.NET 9+), light or dark as Windows
        (System), or Light / Dark for the documentation images. Returns Fluent and Dark.
    #>
    param([ValidateSet('System', 'Light', 'Dark')][string]$Theme = 'System')

    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
    $dark = $Theme -eq 'Dark'
    if ($Theme -eq 'System') { $dark = Test-MclGuiDarkMode }
    # One application per process: created once, never shut down by a closed window.
    $app = [Windows.Application]::Current
    if (-not $app) {
        $app = [Windows.Application]::new()
        $app.ShutdownMode = [Windows.ShutdownMode]::OnExplicitShutdown
    }
    $fluent = $null -ne [Windows.Application].GetProperty('ThemeMode')
    if ($fluent) {
        # ThemeMode is the Fluent theme of WPF (.NET 9 and later). Light or Dark, never System: without the
        # setting (Windows Server 2016) WPF would pick dark and the colours of the window would not match.
        $app.ThemeMode = [Windows.ThemeMode]::new($(if ($dark) { 'Dark' } else { 'Light' }))
    }
    [pscustomobject]@{ Fluent = $fluent; Dark = $dark; Application = $app }
}

function Set-MclGuiTheme {
    <#
        Colours of the window: the accent of the report on the Fluent accent resources, the status colours,
        and with the classic theme (.NET 8) the Fluent resources the window uses, with the report palette.
    #>
    param([Parameter(Mandatory = $true)][Windows.Window]$Window, [Parameter(Mandatory = $true)][pscustomobject]$Theme)

    $dark = $Theme.Dark
    $r = $Window.Resources
    $set = { param([string[]]$Keys, [string]$LightColor, [string]$DarkColor) $b = New-MclGuiBrush $(if ($dark) { $DarkColor } else { $LightColor }); foreach ($k in $Keys) { $r[$k] = $b } }
    if (-not $Theme.Fluent) {
        & $set 'ApplicationBackgroundBrush' '#F7F4EF' '#202020'
        & $set 'CardBackgroundFillColorDefaultBrush' '#FFFFFF' '#2B2B2B'
        & $set 'CardStrokeColorDefaultBrush', 'ControlStrokeColorDefaultBrush' '#DEDEDE' '#3D3D3D'
        & $set 'ControlFillColorDefaultBrush' '#FFFFFF' '#2D2D2D'
        & $set 'ControlFillColorSecondaryBrush' '#F5F5F5' '#323232'
        & $set 'DividerStrokeColorDefaultBrush' '#DEDEDE' '#3D3D3D'
        & $set 'LayerFillColorDefaultBrush' '#FCFBF8' '#262626'
        & $set 'TextFillColorPrimaryBrush' '#242424' '#FFFFFF'
        & $set 'TextFillColorSecondaryBrush' '#5C5C5C' '#C5C5C5'
        & $set 'TextFillColorTertiaryBrush' '#8A8A8A' '#9A9A9A'
    }
    # The accent of the report instead of the accent colour of Windows.
    & $set 'AccentFillColorDefaultBrush', 'AccentButtonBackground', 'AccentButtonBorderBrush' '#B11F4B' '#FD8EA1'
    & $set 'AccentFillColorSecondaryBrush', 'AccentButtonBackgroundPointerOver' '#E6B11F4B' '#E6FD8EA1'
    & $set 'AccentFillColorTertiaryBrush', 'AccentButtonBackgroundPressed' '#CCB11F4B' '#CCFD8EA1'
    & $set 'AccentTextFillColorPrimaryBrush' '#9A1A41' '#FD8EA1'
    # Boxes, radio buttons, progress bar, calendar of the date pickers: the same accent.
    & $set 'CheckBoxCheckBackgroundFillChecked', 'CheckBoxCheckBackgroundStrokeChecked', 'RadioButtonOuterEllipseCheckedFill', 'RadioButtonOuterEllipseCheckedStroke',
        'ProgressBarForeground', 'CalendarViewSelectedBackground', 'CalendarViewSelectedBorderBrush' '#B11F4B' '#FD8EA1'
    & $set 'CheckBoxCheckBackgroundFillCheckedPointerOver', 'CheckBoxCheckBackgroundStrokeCheckedPointerOver', 'RadioButtonOuterEllipseCheckedStrokePointerOver' '#E6B11F4B' '#E6FD8EA1'
    & $set 'CheckBoxCheckBackgroundFillCheckedPressed', 'CheckBoxCheckBackgroundStrokeCheckedPressed' '#CCB11F4B' '#CCFD8EA1'
    # Row selected in the lists: a soft accent, the text stays readable.
    & $set 'DataGridRowSelectedBackgroundThemeBrush' '#1FB11F4B' '#40FD8EA1'
    & $set 'DataGridRowSelectedForegroundThemeBrush' '#242424' '#FFFFFF'
    $r[[Windows.SystemColors]::HighlightBrushKey] = $r['DataGridRowSelectedBackgroundThemeBrush']
    $r[[Windows.SystemColors]::InactiveSelectionHighlightBrushKey] = $r['DataGridRowSelectedBackgroundThemeBrush']
    $r[[Windows.SystemColors]::HighlightTextBrushKey] = $r['DataGridRowSelectedForegroundThemeBrush']
    $r[[Windows.SystemColors]::InactiveSelectionHighlightTextBrushKey] = $r['DataGridRowSelectedForegroundThemeBrush']
    & $set 'MclBrand' '#B11F4B' '#B11F4B'
    & $set 'MclBrandText' '#B11F4B' '#FD8EA1'
    & $set 'MclAccentSoft' '#14B11F4B' '#33FD8EA1'
    & $set 'MclSuccess' '#16A34A' '#4ADE80'
    & $set 'MclCaution' '#D97706' '#FBBF24'
    & $set 'MclCritical' '#DC2626' '#F87171'
    & $set 'MclInfoBackground' '#F3F3F3' '#2E2E2E'
    & $set 'MclCautionBackground' '#FFF7E8' '#33FBBF24'
    & $set 'MclCriticalBackground' '#FDECEC' '#33F87171'
    & $set 'MclSuccessBackground' '#EAF7EE' '#334ADE80'
}

function Invoke-MclGuiPump {
    <# Lets the window repaint and handle clicks during a run (the WPF equivalent of DoEvents). #>
    $frame = [Windows.Threading.DispatcherFrame]::new()
    [void][Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::Background,
        [Windows.Threading.DispatcherOperationCallback] { param($f) $f.Continue = $false; $null }, $frame)
    [Windows.Threading.Dispatcher]::PushFrame($frame)
}

function New-MclForm {
    <#
    .SYNOPSIS
        Builds the window (without showing it). Used by Show-MclGui, the tests and the documentation tool.
    .PARAMETER Theme
        System (like Windows), Light or Dark (documentation images, tests).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Configuration, [ValidateSet('System', 'Light', 'Dark')][string]$Theme = 'System')

    $look = Initialize-MclGuiTheme -Theme $Theme
    $window = [Windows.Markup.XamlReader]::Parse((Get-MclGuiXaml))
    $window.Title = "Meeting Cleanup $($script:ToolVersion)"
    Set-MclGuiTheme -Window $window -Theme $look
    $controls = @{}
    foreach ($name in 'Root', 'Header', 'Version', 'SettingsScroll', 'Inputs', 'Organizer', 'LoadOrganizers', 'OrganizerHint', 'Restore', 'StartDate', 'EndDate', 'Subject', 'ScopeOrganizer', 'ScopeRooms',
        'ScopeMailboxes', 'MailboxFile', 'Browse', 'ScopeAll', 'ActionRemove', 'ActionCancel', 'Comment', 'ActionTransfer', 'TransferHint', 'NewOrganizer', 'TransferMethod', 'WhoTitle', 'ModeOrganizers', 'ModeRooms', 'PeriodHint', 'SearchCard', 'ConnectionExpander', 'TenantId', 'AppId', 'AuthMode', 'ThumbPanel',
        'Thumbprint', 'SecretPanel', 'Secret', 'ConfigHint', 'Counts', 'StatusPill', 'Status', 'SelectAll', 'SelectNone', 'Meetings', 'MeetingsEmpty', 'CopiesTitle', 'Copies',
        'ProgressPanel', 'ProgressBar', 'ProgressText', 'LogScroll', 'Log', 'Actions', 'Search', 'Apply', 'ApplyIcon', 'ApplyText', 'Stop', 'Footer', 'OpenReport', 'OpenFolder', 'Close') {
        $controls[$name] = $window.FindName($name)
    }
    foreach ($b in 'Search', 'Apply') {
        if ($look.Fluent) { $controls[$b].SetResourceReference([Windows.FrameworkElement]::StyleProperty, 'AccentButtonStyle') }
        else { $controls[$b].SetResourceReference([Windows.Controls.Control]::BackgroundProperty, 'AccentFillColorDefaultBrush'); $controls[$b].Foreground = [Windows.Media.Brushes]::White }
    }
    $controls.Version.Text = "v$($script:ToolVersion)  " + [char]0x00B7 + '  Nicolas Fabert'

    # Values of the configuration.
    $zone = Get-MclTimeZone $Configuration.TimeZone
    $today = [TimeZoneInfo]::ConvertTimeFromUtc([datetime]::UtcNow, $zone).Date
    $controls.StartDate.SelectedDate = $today.AddDays(-[int]$Configuration.PastDays)
    $controls.EndDate.SelectedDate = $today.AddDays([int]$Configuration.FutureDays)
    $scopes = @($Configuration.SearchIn)
    $controls.ScopeOrganizer.IsChecked = $scopes -contains 'Organizer'
    $controls.ScopeRooms.IsChecked = $scopes -contains 'Rooms'
    $controls.ScopeMailboxes.IsChecked = $scopes -contains 'Mailboxes'
    $controls.ScopeAll.IsChecked = $scopes -contains 'AllMailboxes'
    $controls.MailboxFile.Text = [string]$Configuration.MailboxFile
    $controls.ActionRemove.IsChecked = $true
    $controls.ModeOrganizers.IsChecked = $true
    $controls.Comment.Text = [string]$Configuration.CancelComment
    foreach ($m in 'Auto', 'Native', 'Recreate') { [void]$controls.TransferMethod.Items.Add($m) }
    $controls.TransferMethod.SelectedItem = [string]$Configuration.TransferMethod
    $controls.TenantId.Text = [string]$Configuration.TenantId
    $controls.AppId.Text = [string]$Configuration.AppId
    foreach ($m in 'Certificate', 'ClientSecret') { [void]$controls.AuthMode.Items.Add($m) }
    $controls.AuthMode.SelectedItem = [string]$Configuration.AuthMode
    $controls.Thumbprint.Text = [string]$Configuration.CertificateThumbprint
    $controls.ConfigHint.Text = "From $([string](Get-MclProperty $Configuration 'ConfigPath')). Changes here are for this window only."
    $controls.ConnectionExpander.IsExpanded = -not ($Configuration.TenantId -and $Configuration.AppId)

    $meetings = [Collections.ObjectModel.ObservableCollection[object]]::new()
    $copies = [Collections.ObjectModel.ObservableCollection[object]]::new()
    $items = [Collections.ObjectModel.ObservableCollection[object]]::new()
    $controls.Meetings.ItemsSource = $meetings
    $controls.Copies.ItemsSource = $copies
    $controls.Log.ItemsSource = $items
    $script:Gui = @{
        Form = $window; Controls = $controls; Configuration = $Configuration.Clone(); Settings = $null; Theme = $look
        Running = $false; Result = $null; Acted = $false; LastReport = $null; LastFolder = $null; LastAction = ''; RestoreSource = $null
        Rows = $meetings; CopyRows = $copies; Items = $items; Lines = [Collections.Generic.List[string]]::new()
        # Attached to Closing only while a run is in progress: closing then stops the run first.
        ClosingGuard = [ComponentModel.CancelEventHandler] {
            param($sender, $e)
            $e.Cancel = $true
            if ($script:Ui) { $script:Ui.Cancel = $true }
            Add-MclGuiLine 'Warn' 'A run is in progress: it stops at the next Graph call, then the window can be closed.'
        }
    }
    Set-MclGuiStatus 'Ready' 'Ready'
    $controls.Footer.Text = "Reports: $($Configuration.OutputPath)"

    $controls.Search.Add_Click({ Invoke-MclGuiSearch })
    $controls.Apply.Add_Click({ Invoke-MclGuiApply })
    $controls.Stop.Add_Click({
            if ($script:Ui) {
                $script:Ui.Cancel = $true
                Add-MclGuiLine 'Warn' $(if ($script:Ui.Hold) { 'Stop requested: the meetings being re-created are finished first (created, sent, old copies removed), then the run stops.' } else { 'Stop requested: the run stops at the next Graph call.' })
            }
        })
    $controls.SelectAll.Add_Click({ Set-MclGuiSelection $true })
    $controls.SelectNone.Add_Click({ Set-MclGuiSelection $false })
    $controls.Meetings.Add_SelectionChanged({ Update-MclGuiCopies })
    # A box ticked or unticked in the list: the action button counts again.
    $controls.Meetings.AddHandler([Windows.Controls.Primitives.ButtonBase]::ClickEvent, [Windows.RoutedEventHandler] { Update-MclGuiState })
    $controls.ActionRemove.Add_Checked({ Update-MclGuiState })
    $controls.ActionCancel.Add_Checked({ Update-MclGuiState })
    $controls.ActionTransfer.Add_Checked({ Update-MclGuiState })
    $controls.ModeOrganizers.Add_Checked({ Update-MclGuiState; Update-MclGuiOrganizerHint })
    $controls.ModeRooms.Add_Checked({ Update-MclGuiState; Update-MclGuiOrganizerHint })
    $controls.AuthMode.Add_SelectionChanged({ Update-MclGuiState })
    $controls.ScopeMailboxes.Add_Click({ Update-MclGuiState })
    $controls.Browse.Add_Click({
            $dialog = [Microsoft.Win32.OpenFileDialog]::new()
            $dialog.Filter = 'Addresses (*.txt;*.csv)|*.txt;*.csv|All files (*.*)|*.*'
            if ($dialog.ShowDialog($script:Gui.Form)) { $script:Gui.Controls.MailboxFile.Text = $dialog.FileName; $script:Gui.Controls.ScopeMailboxes.IsChecked = $true; Update-MclGuiState }
        })
    $controls.LoadOrganizers.Add_Click({ Import-MclGuiOrganizers })
    $controls.Organizer.Add_TextChanged({ Update-MclGuiOrganizerHint })
    $controls.Restore.Add_Click({ Invoke-MclGuiRestore })
    $controls.OpenReport.Add_Click({ if ($script:Gui.LastReport) { Start-Process -FilePath $script:Gui.LastReport } })
    $controls.OpenFolder.Add_Click({ if ($script:Gui.LastFolder) { Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$($script:Gui.LastFolder)`"" } })

    # The window fits the screen where it opens (small laptop screen at 150 %); the left column scrolls.
    $area = [Windows.SystemParameters]::WorkArea
    $window.Width = [Math]::Min($window.Width, $area.Width)
    $window.Height = [Math]::Min($window.Height, $area.Height)
    $window.MinWidth = [Math]::Min($window.MinWidth, $area.Width)
    $window.MinHeight = [Math]::Min($window.MinHeight, $area.Height)
    Update-MclGuiState
    [pscustomobject]@{ Form = $window; Controls = $controls; Lines = $script:Gui.Lines; Items = $items; Rows = $meetings }
}

function Update-MclGuiState {
    <# Enables what can be used now: the action button and its text, the message, the secret, the file. #>
    $g = $script:Gui
    if (-not $g) { return }
    $c = $g.Controls
    $rooms = [bool]$c.ModeRooms.IsChecked
    # Rooms: every meeting of the rooms given; a transfer moves the meetings of organizers only (also when the
    # meetings in the list come from a rooms search and the mode was switched back).
    $roomsResult = $g.Result -and [string](Get-MclProperty $g.Result.Request 'Mode') -eq 'Rooms'
    if (($rooms -or $roomsResult) -and $c.ActionTransfer.IsChecked) { $c.ActionRemove.IsChecked = $true }
    $c.ActionTransfer.IsEnabled = -not ($rooms -or $roomsResult)
    $c.SearchCard.IsEnabled = -not $rooms
    $c.WhoTitle.Text = if ($rooms) { 'Rooms' } else { 'Organizers' }
    $c.PeriodHint.Text = if ($rooms) { 'Every meeting of the rooms in the period, whatever its organizer. A series: only its occurrences in the period are acted on (the series goes on outside it).' } else { 'A series is found when one of its occurrences falls in the period, and is handled as a whole.' }
    $cancel = [bool]$c.ActionCancel.IsChecked
    $transfer = [bool]$c.ActionTransfer.IsChecked
    $c.Comment.IsEnabled = $cancel
    $c.NewOrganizer.IsEnabled = $transfer
    $c.TransferMethod.IsEnabled = $transfer
    $secret = [string]$c.AuthMode.SelectedItem -eq 'ClientSecret'
    $c.SecretPanel.Visibility = if ($secret) { 'Visible' } else { 'Collapsed' }
    $c.ThumbPanel.Visibility = if ($secret) { 'Collapsed' } else { 'Visible' }
    $c.MailboxFile.IsEnabled = [bool]$c.ScopeMailboxes.IsChecked
    $ticked = @($g.Rows | Where-Object Selected).Count
    $c.ApplyIcon.Text = [string][char]$(if ($cancel) { 0xE711 } elseif ($transfer) { 0xE748 } else { 0xE74D })
    $verb = if ($cancel) { 'Cancel' } elseif ($transfer) { 'Transfer' } else { 'Remove' }
    $c.ApplyText.Text = if ($ticked) { "$verb $ticked meeting$(if ($ticked -gt 1) { 's' })" } else { "$verb the meetings ticked" }
    $c.Apply.IsEnabled = -not $g.Running -and $null -ne $g.Result -and -not $g.Acted -and $ticked -gt 0
    if ($null -ne $g.Result) { $c.Counts.Text = '{0} found {1} {2} ticked' -f @($g.Rows).Count, [char]0x00B7, $ticked }
    else { $c.Counts.Text = '' }
}

function Set-MclGuiSelection {
    param([bool]$Value)
    $g = $script:Gui
    if (-not $g -or $g.Running -or $g.Acted) { return }
    foreach ($row in @($g.Rows)) { $row.Selected = $Value }
    $g.Controls.Meetings.Items.Refresh()
    Update-MclGuiState
}

function Update-MclGuiCopies {
    <# The copies of the meeting selected in the list. #>
    $g = $script:Gui
    if (-not $g) { return }
    $g.CopyRows.Clear()
    $row = $g.Controls.Meetings.SelectedItem
    if (-not $row) { $g.Controls.CopiesTitle.Text = 'Copies of the meeting selected'; return }
    foreach ($c in @($row.Meeting.Copies)) {
        $g.CopyRows.Add([pscustomobject]@{ Mailbox = $c.Mailbox; Role = $c.Role; Occurrence = [string](Get-MclProperty $c 'Occurrence'); Via = $c.Via; Result = $(if ($c.Result) { $c.Result } elseif ($c.EventId) { 'Found' } else { '' }); Detail = $c.Detail })
    }
    $g.Controls.CopiesTitle.Text = "Copies of '$($row.Subject)'  $([char]0x00B7)  organizer $($row.Meeting.Organizer)"
    Update-MclGuiState
}

function Add-MclGuiLine {
    <# One line of the progress: icon and colour of its status. 'Progress' updates the progress bar instead. #>
    param([string]$Status, [string]$Text)

    $g = $script:Gui
    if (-not $g) { return }
    if ($Status -eq 'Progress') {
        $fraction = 0.0
        $parts = $Text.Split('|', 2)
        if ($parts.Count -eq 2 -and [double]::TryParse($parts[0], [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$fraction)) { $Text = $parts[1] }
        $g.Controls.ProgressPanel.Visibility = 'Visible'
        $g.Controls.ProgressBar.Value = $fraction
        $g.Controls.ProgressText.Text = $Text
        Invoke-MclGuiPump
        return
    }
    $glyphs = @{ Step = 0xE76C; Ok = 0xE73E; Warn = 0xE7BA; Fail = 0xEA39; Info = 0xE946; Skip = 0xE72A }
    $colours = @{ Step = 'MclBrandText'; Ok = 'MclSuccess'; Warn = 'MclCaution'; Fail = 'MclCritical'; Info = 'TextFillColorSecondaryBrush'; Skip = 'TextFillColorTertiaryBrush' }
    $key = if ($glyphs.ContainsKey($Status)) { $Status } else { 'Info' }
    $step = $Status -eq 'Step'
    if ($step) { $g.Controls.ProgressPanel.Visibility = 'Collapsed' }
    $shown = if ($step) { $Text -replace '^\[(\d+/\d+)\]\s*', '$1   ' } else { $Text }
    $window = $g.Form
    $g.Items.Add([pscustomobject]@{
            Glyph     = [string][char]$glyphs[$key]
            Brush     = $window.TryFindResource($colours[$key])
            Text      = $shown
            TextBrush = $window.TryFindResource($(if ($key -in 'Info', 'Skip') { 'TextFillColorSecondaryBrush' } else { 'TextFillColorPrimaryBrush' }))
            Weight    = if ($step) { [Windows.FontWeights]::SemiBold } else { [Windows.FontWeights]::Normal }
            Size      = if ($step) { 13 } else { 12 }
            Margin    = if ($step) { [Windows.Thickness]::new(0, $(if ($g.Items.Count) { 10 } else { 0 }), 0, 3) } else { [Windows.Thickness]::new(0, 1, 0, 1) }
            Time      = (Get-Date).ToString('HH:mm:ss')
        })
    $g.Lines.Add("[$Status] $Text")
    $g.Controls.LogScroll.ScrollToEnd()
    Invoke-MclGuiPump
}

function Set-MclGuiStatus {
    <# The status pill of the meetings: Ready, Running, Completed, Warning or Failed. #>
    param([string]$Text, [string]$Status)

    $c = $script:Gui.Controls
    $c.Status.Text = $Text
    $pair = switch ($Status) {
        'Completed' { 'MclSuccess', 'MclSuccessBackground' }
        'Failed' { 'MclCritical', 'MclCriticalBackground' }
        'Warning' { 'MclCaution', 'MclCautionBackground' }
        'Running' { 'MclBrandText', 'MclAccentSoft' }
        default { 'TextFillColorSecondaryBrush', 'MclInfoBackground' }
    }
    $c.Status.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, $pair[0])
    $c.StatusPill.SetResourceReference([Windows.Controls.Border]::BackgroundProperty, $pair[1])
}

function Get-MclGuiSettings {
    <# The configuration with the connection fields of the window. #>
    $g = $script:Gui
    $c = $g.Controls
    $cfg = $g.Configuration.Clone()
    $cfg.TenantId = $c.TenantId.Text.Trim()
    $cfg.AppId = $c.AppId.Text.Trim()
    $cfg.AuthMode = [string]$c.AuthMode.SelectedItem
    $cfg.CertificateThumbprint = ($c.Thumbprint.Text -replace '\s', '').Trim()
    return $cfg
}

function Update-MclGuiRows {
    <# The list of the meetings from the result (after a search or an action). #>
    $g = $script:Gui
    $g.Rows.Clear()
    foreach ($m in @($g.Result.Meetings)) {
        # One copy per mailbox: an occurrence is counted once for its mailbox; the new organizer is not a copy.
        $real = @($m.Copies | Where-Object { $_.EventId -and $_.Role -in 'Organizer', 'Attendee', 'Room' } | Group-Object Mailbox | ForEach-Object { $_.Group[0] })
        $rooms = @($real | Where-Object Role -eq 'Room').Count
        $kind = if ((Get-MclProperty $m 'Scope') -eq 'Occurrences') { '{0} occ.' -f $m.Occurrences } else { $m.Kind }
        $g.Rows.Add([pscustomobject]@{
                Selected = [bool]$m.Selected; CanTick = -not $g.Acted; Start = $m.StartText; Subject = $m.Subject; Who = $(if ($m.OrganizerName) { $m.OrganizerName } else { $m.Organizer }); Kind = $kind; OrganizerCopy = $m.OrganizerCopy
                Copies = $real.Count; Rooms = $rooms; CopiesText = $(if ($rooms) { '{0} ({1} room{2})' -f $real.Count, $rooms, $(if ($rooms -gt 1) { 's' }) } else { [string]$real.Count }); Status = $m.Status; Meeting = $m
            })
    }
    # The Organizer column only when the meetings come from more than one organizer.
    $several = @($g.Result.Meetings | ForEach-Object { if ($_.OrganizerKey) { $_.OrganizerKey } else { $_.Organizer } } | Select-Object -Unique).Count -gt 1
    foreach ($column in $g.Controls.Meetings.Columns) { if ($column.Header -eq 'Organizer') { $column.Visibility = if ($several) { 'Visible' } else { 'Collapsed' } } }
    $g.Controls.MeetingsEmpty.Visibility = if ($g.Rows.Count) { 'Collapsed' } else { 'Visible' }
    if (-not $g.Rows.Count) { $g.Controls.MeetingsEmpty.Text = 'No meeting found: widen the period, check the address, or search in more mailboxes.' }
    if ($g.Rows.Count) { $g.Controls.Meetings.SelectedIndex = 0 }
    Update-MclGuiCopies
    Update-MclGuiState
}

function Start-MclGuiRun {
    param([string]$Text)
    $g = $script:Gui
    $c = $g.Controls
    $script:Ui = @{ Sink = { param($Status, $Text) Add-MclGuiLine $Status $Text }; Pump = { Invoke-MclGuiPump }; Cancel = $false }
    $g.Running = $true
    $g.Form.add_Closing($g.ClosingGuard)
    foreach ($b in 'Search', 'Apply', 'Restore', 'OpenReport', 'OpenFolder', 'Close', 'SelectAll', 'SelectNone', 'Inputs', 'Meetings') { $c[$b].IsEnabled = $false }
    $c.Stop.IsEnabled = $true
    Set-MclGuiStatus $Text 'Running'
}

function Stop-MclGuiRun {
    $g = $script:Gui
    $c = $g.Controls
    $g.Form.remove_Closing($g.ClosingGuard)
    $script:Ui = $null
    $g.Running = $false
    foreach ($b in 'Search', 'Restore', 'Close', 'SelectAll', 'SelectNone', 'Inputs', 'Meetings') { $c[$b].IsEnabled = $true }
    $c.Stop.IsEnabled = $false
    $c.ProgressPanel.Visibility = 'Collapsed'
    $c.OpenReport.IsEnabled = [bool]$g.LastReport
    $c.OpenFolder.IsEnabled = [bool]$g.LastFolder
    Update-MclGuiState
}

function Save-MclGuiReport {
    param([Parameter(Mandatory = $true)][hashtable]$Settings, [string]$Directory)
    $g = $script:Gui
    Write-MclNextStep 'Report' 'Report'
    $exportArgs = @{ Result = $g.Result; OutputPath = $Settings.OutputPath; Prefix = $Settings.ReportPrefix; Formats = $Settings.ReportFormats; Delimiter = $Settings.CsvDelimiter }
    if ($Directory) { $exportArgs.Directory = $Directory }
    $report = Export-MclReport @exportArgs
    $g.LastFolder = $report.Directory
    $g.LastReport = Get-MclProperty $report.Files 'Html'
    Write-MclItem Ok "Report: $($report.Directory)" -Icon File
    $g.Controls.Footer.Text = "Report: $($report.Directory)"
}

function Invoke-MclGuiSearch {
    $g = $script:Gui
    $c = $g.Controls
    $g.Items.Clear(); $g.Lines.Clear()
    $cfg = Get-MclGuiSettings
    $scopes = @(
        if ($c.ScopeOrganizer.IsChecked) { 'Organizer' }
        if ($c.ScopeRooms.IsChecked) { 'Rooms' }
        if ($c.ScopeMailboxes.IsChecked) { 'Mailboxes' }
        if ($c.ScopeAll.IsChecked) { 'AllMailboxes' }
    )
    $rooms = [bool]$c.ModeRooms.IsChecked
    $requestArgs = if ($rooms) { @{ Settings = $cfg; Room = @(Split-MclAddressList @($c.Organizer.Text)); Subject = $c.Subject.Text.Trim(); Action = 'Report' } }
        else { @{ Settings = $cfg; Organizer = @($c.Organizer.Text); Subject = $c.Subject.Text.Trim(); SearchIn = $scopes; Action = 'Report' } }
    if ($c.StartDate.SelectedDate) { $requestArgs.Start = [datetime]$c.StartDate.SelectedDate }
    if ($c.EndDate.SelectedDate) { $requestArgs.End = [datetime]$c.EndDate.SelectedDate }
    if ($c.ScopeMailboxes.IsChecked -and $c.MailboxFile.Text.Trim()) { $requestArgs.MailboxFile = $c.MailboxFile.Text.Trim() }
    $problems = [Collections.Generic.List[string]]::new()
    if (-not $scopes.Count -and -not $rooms) { $problems.Add('Tick at least one place to search.') }
    if ($rooms -and -not @(Split-MclAddressList @($c.Organizer.Text)).Count) { $problems.Add('Give the rooms: one address per line, or Load a list...') }
    $request = $null
    try {
        $request = New-MclRequest @requestArgs
        foreach ($p in (Test-MclRequest -Request $request).Problems) { if ($p -notlike 'Search in:*' -or $scopes.Count) { $problems.Add($p) } }
    }
    catch { $problems.Add($_.Exception.Message) }
    foreach ($p in (Test-MclConfiguration -Configuration $cfg -ForConnection).Problems) { $problems.Add("$p (Connection, at the bottom left)") }
    if ($problems.Count) {
        foreach ($p in $problems) { Add-MclGuiLine 'Fail' $p }
        Set-MclGuiStatus 'Fix the values on the left' 'Failed'
        return
    }

    $secret = $null
    Start-MclGuiRun 'Searching...'
    try {
        $g.Result = $null; $g.Acted = $false; $g.LastReport = $null; $g.LastFolder = $null
        $g.Rows.Clear(); $g.CopyRows.Clear()
        Write-MclLog 'STEP' "Window search: $(if ($rooms) { "rooms $($request.Room -join ', ')" } else { $request.Organizer -join ', ' }) from $($request.Start.ToString('o')) to $($request.End.ToString('o'))$(if (-not $rooms) { " in $($scopes -join ', ')" })"
        Initialize-MclSteps -Total 6
        Write-MclNextStep 'Microsoft Graph' 'Key'
        if ($cfg.AuthMode -eq 'ClientSecret' -and $c.Secret.SecurePassword.Length) { $secret = $c.Secret.SecurePassword.Copy() }
        $connection = Connect-MclGraph -Settings $cfg -Secret $secret -Action 'Report'
        Write-MclItem Ok ('Application {0} {1} tenant {2}' -f $(if ($connection.AppName) { $connection.AppName } else { $cfg.AppId }), [char]0x00B7, $connection.TenantGuid) -Icon Key
        if (-not $connection.CanWrite) { Write-MclItem Warn 'Calendars.Read only: the meetings can be listed, not removed or cancelled.' }
        $g.Result = Find-MclMeetings -Settings $cfg -Request $request
        $g.Settings = $cfg
        Save-MclGuiReport -Settings $cfg
        Update-MclGuiRows
        $n = $g.Result.Counts
        Set-MclGuiStatus ('{0} meeting(s) {1} {2} copies' -f $n.Meetings, [char]0x00B7, $n.Copies) $g.Result.Status
        if ($n.Meetings) { Add-MclGuiLine 'Info' 'Nothing has been changed. Untick the meetings to keep, choose the action on the left, then use the action button.' }
    }
    catch [OperationCanceledException] {
        Add-MclGuiLine 'Warn' 'Search stopped: nothing was changed.'
        Set-MclGuiStatus 'Stopped' 'Warning'
    }
    catch {
        Add-MclGuiLine 'Fail' $_.Exception.Message
        Set-MclGuiStatus 'Failed - see the progress' 'Failed'
    }
    finally {
        if ($secret) { $secret.Dispose() }
        Stop-MclGuiRun
    }
}

function Show-MclGuiQuestion {
    <# A question in a message box; returns the button clicked. $script:GuiAnswers (a queue filled by the lab tests) answers without showing the box. #>
    param([Parameter(Mandatory = $true)][string]$Text, [string]$Title = 'Meeting Cleanup', [string]$Buttons = 'YesNo', [string]$Image = 'Question')
    if ($null -ne $script:GuiAnswers -and $script:GuiAnswers.Count) {
        $answer = [Windows.MessageBoxResult]$script:GuiAnswers.Dequeue()
        Write-MclLog 'INFO' "$Title - answered $answer in advance: $($Text.Split("`n")[0])"
        return $answer
    }
    [Windows.MessageBox]::Show($script:Gui.Form, $Text, $Title, [Windows.MessageBoxButton]$Buttons, [Windows.MessageBoxImage]$Image, [Windows.MessageBoxResult]::No)
}

function Invoke-MclGuiApply {
    $g = $script:Gui
    $c = $g.Controls
    if ($null -eq $g.Result -or $g.Acted) { return }
    foreach ($row in @($g.Rows)) { $row.Meeting.Selected = [bool]$row.Selected }
    if ($c.ActionTransfer.IsChecked) { Invoke-MclGuiTransfer; return }
    $action = if ($c.ActionCancel.IsChecked) { 'Cancel' } else { 'Remove' }
    $comment = $c.Comment.Text.Trim()
    if ($action -eq 'Cancel' -and $comment -match '[<>]') { Add-MclGuiLine 'Fail' 'The cancellation message must be plain text (no < or >).'; return }
    $plan = Get-MclCleanupPlan -Result $g.Result -Action $action
    if (-not $plan.Meetings.Count) { return }
    $text = "$(if ($action -eq 'Cancel') { 'Cancel and clean' } else { 'Remove silently' }):`n`n - " + ($plan.Lines -join "`n - ")
    if ($action -eq 'Cancel') { $text += "`n`nMessage: $comment" }
    $text += "`n`nA backup is written first. The copies removed can be put back with Restore..., for the retention of deleted items (14 days by default). Continue?"
    $answer = Show-MclGuiQuestion -Text $text -Image Warning
    if ($answer -ne [Windows.MessageBoxResult]::Yes) { Add-MclGuiLine 'Info' 'Nothing was changed.'; return }

    $secret = $null
    $runPath = $null
    Start-MclGuiRun $(if ($action -eq 'Cancel') { 'Cancelling...' } else { 'Removing...' })
    try {
        $cfg = $g.Settings
        Write-MclLog 'INFO' "Confirmed in the window by $([Environment]::UserName): $($plan.Text)"
        Initialize-MclSteps -Total (2 + [int][bool]$cfg.Verify)
        if ($cfg.AuthMode -eq 'ClientSecret' -and $c.Secret.SecurePassword.Length) { $secret = $c.Secret.SecurePassword.Copy() }
        $null = Connect-MclGraph -Settings $cfg -Secret $secret -Action $action
        $g.Acted = $true
        $runPath = New-MclRunFolder -OutputPath $cfg.OutputPath -Prefix $cfg.ReportPrefix -Action $action
        $g.Result = Invoke-MclCleanup -Settings $cfg -Result $g.Result -Action $action -Comment $comment -BackupPath (Join-Path $runPath "$($cfg.ReportPrefix)-Backup.json")
        Save-MclGuiReport -Settings $cfg -Directory $runPath
        $g.LastAction = $action
        Update-MclGuiRows
        $n = $g.Result.Counts
        Set-MclGuiStatus ('{0} {1} {2} removed {1} {3} cancelled {1} {4} failed' -f $g.Result.Status, [char]0x00B7, $n.Removed, $n.Cancelled, $n.Failed) $g.Result.Status
        Add-MclGuiLine 'Info' $(if ($action -eq 'Remove') { 'Search again to see what is left. To undo it: Restore... (the copies come back, no message).' } else { 'Search again to see what is left, or to clean other meetings.' })
    }
    catch [OperationCanceledException] {
        Add-MclGuiLine 'Warn' 'Stopped: the copies already handled are in the report (they can be restored), the others were left as they were.'
        Set-MclGuiStatus 'Stopped' 'Warning'
        if ($null -ne $g.Result -and $runPath) {
            try {
                $g.Result.Status = 'Warning'; $g.Result.Error = 'Stopped by the user'
                Update-MclResultCounts $g.Result
                Save-MclGuiReport -Settings $g.Settings -Directory $runPath
                $g.LastAction = $action
                Update-MclGuiRows
            }
            catch { Write-MclLog 'WARN' "Report after a stop: $($_.Exception.Message)" }
        }
    }
    catch {
        Add-MclGuiLine 'Fail' $_.Exception.Message
        Set-MclGuiStatus 'Failed - see the progress' 'Failed'
    }
    finally {
        if ($secret) { $secret.Dispose() }
        Stop-MclGuiRun
    }
}

function Invoke-MclGuiTransfer {
    <# Transfer of the meetings ticked to the new organizer of the window: plan, confirmation, Exchange Online if needed. #>
    $g = $script:Gui
    $c = $g.Controls
    $address = $c.NewOrganizer.Text.Trim()
    if ($address -notmatch $script:SmtpPattern) { Add-MclGuiLine 'Fail' 'Transfer: type the address of the new organizer (a mailbox of the tenant).'; return }
    $method = [string]$c.TransferMethod.SelectedItem
    $cfg = $g.Settings
    $secret = $null
    $runPath = $null
    $plan = $null
    Start-MclGuiRun 'Transferring...'
    try {
        if ($cfg.AuthMode -eq 'ClientSecret' -and $c.Secret.SecurePassword.Length) { $secret = $c.Secret.SecurePassword.Copy() }
        Initialize-MclSteps -Total (3 + [int][bool]$cfg.Verify)
        $null = Connect-MclGraph -Settings $cfg -Secret $secret -Action 'Remove'
        $new = Resolve-MclNewOrganizer -Address $address
        $plan = Get-MclTransferPlan -Result $g.Result -NewOrganizer $new -Method $method -Comment ([string]$cfg.TransferComment)
        if (-not ($plan.Native.Count + $plan.Recreate.Count)) { foreach ($line in $plan.Lines) { Add-MclGuiLine 'Warn' $line }; Set-MclGuiStatus 'Nothing to transfer' 'Warning'; return }
        $text = "Transfer to $(if ($new.Name) { "$($new.Name) <$($new.Address)>" } else { $new.Address }):`n`n - " + ($plan.Lines -join "`n - ") + "`n`nA backup is written first. Continue?"
        if ((Show-MclGuiQuestion -Text $text -Title 'Meeting Cleanup - Transfer' -Image Warning) -ne [Windows.MessageBoxResult]::Yes) { Add-MclGuiLine 'Info' 'Nothing was changed.'; return }
        Write-MclLog 'INFO' "Transfer confirmed in the window by $([Environment]::UserName): $($plan.Text)"
        $g.Acted = $true
        $runPath = New-MclRunFolder -OutputPath $cfg.OutputPath -Prefix $cfg.ReportPrefix -Action 'Transfer'
        Write-MclNextStep 'Exchange Online PowerShell' 'Server'
        if ($plan.Native.Count) { Connect-MclExchange -Settings $cfg -Secret $secret -For Transfer } else { Write-MclItem Skip 'Not needed: every meeting is re-created with Microsoft Graph.' }
        try { $g.Result = Invoke-MclTransfer -Settings $cfg -Result $g.Result -Plan $plan -Comment ([string]$cfg.TransferComment) -BackupPath (Join-Path $runPath "$($cfg.ReportPrefix)-Backup.json") }
        finally { if ($plan.Native.Count) { Disconnect-MclExchange } }
        Save-MclGuiReport -Settings $cfg -Directory $runPath
        $g.LastAction = 'Transfer'
        Update-MclGuiRows
        $n = $g.Result.Counts
        Set-MclGuiStatus ('{0} {1} {2} transferred {1} {3} failed' -f $g.Result.Status, [char]0x00B7, $n.Transferred, @($g.Result.Meetings | Where-Object Status -eq 'Failed').Count) $g.Result.Status
    }
    catch [OperationCanceledException] {
        Add-MclGuiLine 'Warn' 'Stopped: the meetings already transferred are in the report.'
        Set-MclGuiStatus 'Stopped' 'Warning'
        if ($null -ne $g.Result -and $runPath) {
            try { $g.Result.Status = 'Warning'; $g.Result.Error = 'Stopped by the user'; Update-MclResultCounts $g.Result; Save-MclGuiReport -Settings $cfg -Directory $runPath; $g.LastAction = 'Transfer'; Update-MclGuiRows }
            catch { Write-MclLog 'WARN' "Report after a stop: $($_.Exception.Message)" }
        }
    }
    catch {
        Add-MclGuiLine 'Fail' $_.Exception.Message
        Set-MclGuiStatus 'Failed - see the progress' 'Failed'
    }
    finally {
        if ($secret) { $secret.Dispose() }
        Stop-MclGuiRun
    }
}

function Update-MclGuiOrganizerHint {
    <# Under the organizers: how many addresses are typed, and which are not valid. #>
    $g = $script:Gui
    if (-not $g) { return }
    $rooms = [bool]$g.Controls.ModeRooms.IsChecked
    $list = @(Split-MclAddressList @($g.Controls.Organizer.Text))
    if (-not $list.Count) {
        $g.Controls.OrganizerHint.Text = if ($rooms) { 'One room per line (or separated by ;): SMTP address of the room. A list: text or CSV file.' } else { 'One per line (or separated by ;): SMTP address (any alias), or the X500 address of a deleted mailbox. A list: text or CSV file.' }
        return
    }
    $bad = @($list | Where-Object { $_ -notmatch $script:SmtpPattern -and ($rooms -or $_ -notmatch $script:X500Pattern) })
    $word = if ($rooms) { 'room' } else { 'organizer' }
    $g.Controls.OrganizerHint.Text = '{0} {1}{2}{3}' -f $list.Count, $word, $(if ($list.Count -gt 1) { 's' } else { '' }), $(if ($bad.Count) { " $($script:Dot) not an address: $(($bad | Select-Object -First 3) -join ', ')" } else { '' })
}

function Import-MclGuiOrganizers {
    <# Load a list of organizers (text or CSV file) into the box. #>
    $g = $script:Gui
    $dialog = [Microsoft.Win32.OpenFileDialog]::new()
    $dialog.Filter = 'Organizers (*.txt;*.csv)|*.txt;*.csv|All files (*.*)|*.*'
    $dialog.Title = 'List of organizers: one address per line, or a CSV file (PrimarySmtpAddress, Mail, UserPrincipalName, LegacyExchangeDN...)'
    if (-not $dialog.ShowDialog($g.Form)) { return }
    try {
        $list = @(Read-MclAddressFile -Path $dialog.FileName -AllowX500)
        if (-not $list.Count) { Add-MclGuiLine 'Warn' "No SMTP or X500 address in $($dialog.FileName)."; return }
        $g.Controls.Organizer.Text = ($list -join [Environment]::NewLine)
        Add-MclGuiLine 'Info' "$($list.Count) organizer(s) loaded from $($dialog.FileName)."
    }
    catch { Add-MclGuiLine 'Fail' $_.Exception.Message }
}

function Select-MclGuiRestoreFolder {
    <# The folder of the run to restore: the Remove just done in the window, or a folder chosen under the reports. #>
    $g = $script:Gui
    # The run just done in the window, or the run of the last restore (to finish it after a stop).
    $offer = if ($g.LastAction -in 'Remove', 'Cancel', 'Transfer' -and $g.LastFolder) { $g.LastFolder } elseif ($g.LastAction -eq 'Restore' -and $g.RestoreSource) { $g.RestoreSource } else { $null }
    if ($offer) {
        $answer = Show-MclGuiQuestion -Text "Restore the copies removed by this run?`n`n$offer`n`nNo: choose the folder of another run." -Title 'Meeting Cleanup - Restore' -Buttons YesNoCancel
        if ($answer -eq [Windows.MessageBoxResult]::Yes) { return $offer }
        if ($answer -eq [Windows.MessageBoxResult]::Cancel) { return $null }
    }
    $initial = [string]$g.Configuration.OutputPath
    if ('Microsoft.Win32.OpenFolderDialog' -as [type]) {
        $dialog = [Microsoft.Win32.OpenFolderDialog]::new()
        $dialog.Title = 'Folder of the report of a Remove run (MeetingCleanup_Remove_...)'
        if ($initial -and (Test-Path -LiteralPath $initial)) { $dialog.InitialDirectory = $initial }
        if ($dialog.ShowDialog($g.Form)) { return $dialog.FolderName }
        return $null
    }
    $dialog = [Microsoft.Win32.OpenFileDialog]::new()
    $dialog.Filter = 'Report of a run (*-Summary.json)|*-Summary.json'
    if ($initial -and (Test-Path -LiteralPath $initial)) { $dialog.InitialDirectory = $initial }
    if ($dialog.ShowDialog($g.Form)) { return $dialog.FileName }
    return $null
}

function Invoke-MclGuiRestore {
    <# Restore of a Remove run: plan, confirmation, Exchange Online PowerShell, Recoverable Items, check, report. #>
    $g = $script:Gui
    $c = $g.Controls
    $folder = Select-MclGuiRestoreFolder
    if (-not $folder) { return }
    $g.Items.Clear(); $g.Lines.Clear()
    $cfg = if ($g.Settings) { $g.Settings } else { Get-MclGuiSettings }
    $problems = @((Test-MclConfiguration -Configuration $cfg -ForConnection).Problems)
    if ($problems.Count) { foreach ($p in $problems) { Add-MclGuiLine 'Fail' "$p (Connection, at the bottom left)" }; Set-MclGuiStatus 'Fix the connection' 'Failed'; return }
    try {
        $source = Import-MclRestoreSource -Path $folder
        $plan = Get-MclRestorePlan -Result $source
    }
    catch { Add-MclGuiLine 'Fail' $_.Exception.Message; Set-MclGuiStatus 'Not a report to restore' 'Failed'; return }
    if (-not $plan.Restore.Count) {
        foreach ($line in $plan.Lines) { Add-MclGuiLine 'Info' $line }
        Add-MclGuiLine 'Warn' 'Nothing to restore in this run.'
        return
    }
    $text = "Restore ($($source.SourceAction) run of $([IO.Path]::GetFileName((Split-Path $source.FromReport -Parent)))):`n`n - " + ($plan.Lines -join "`n - ")
    $text += "`n`nExchange Online PowerShell: $(if ($cfg.RestoreConnection -eq 'Interactive') { "as an administrator $($cfg.RestoreUser) (sign-in window)" } else { 'as the application (role Mailbox Import Export)' }). Continue?"
    $answer = Show-MclGuiQuestion -Text $text -Title 'Meeting Cleanup - Restore'
    if ($answer -ne [Windows.MessageBoxResult]::Yes) { Add-MclGuiLine 'Info' 'Nothing was changed.'; return }

    $secret = $null
    Start-MclGuiRun 'Restoring...'
    try {
        Write-MclLog 'INFO' "Restore confirmed in the window by $([Environment]::UserName): $folder - $($plan.Text)"
        Initialize-MclSteps -Total 5
        Write-MclNextStep 'Microsoft Graph' 'Key'
        if ($cfg.AuthMode -eq 'ClientSecret' -and $c.Secret.SecurePassword.Length) { $secret = $c.Secret.SecurePassword.Copy() }
        $connection = Connect-MclGraph -Settings $cfg -Secret $secret -Action 'Remove'
        if ($source.Tenant -and $source.Tenant -ne $connection.TenantGuid) { throw "The report belongs to tenant $($source.Tenant), the application signs in to $($connection.TenantGuid)." }
        Write-MclNextStep 'Exchange Online PowerShell' 'Server'
        Connect-MclExchange -Settings $cfg -Secret $secret
        # From here the meetings shown are those of the restore: never ticked again for an action, whatever happens.
        $g.Acted = $true
        $g.Settings = $cfg
        $g.LastAction = 'Restore'
        $g.RestoreSource = $folder
        $g.Result = $source
        try { $g.Result = Invoke-MclRestore -Settings $cfg -Result $source }
        finally { Disconnect-MclExchange }
        Save-MclGuiReport -Settings $cfg
        Update-MclGuiRows
        $n = $g.Result.Counts
        Set-MclGuiStatus ('{0} {1} {2} restored {1} {3} not found {1} {4} failed' -f $g.Result.Status, [char]0x00B7, $n.Restored, $n.NotFound, $n.Failed) $g.Result.Status
    }
    catch [OperationCanceledException] {
        Add-MclGuiLine 'Warn' 'Stopped: the copies already restored are back; the others are still in Recoverable Items. Restore... again finishes them.'
        Set-MclGuiStatus 'Stopped' 'Warning'
        if ($g.Result -and $g.Result.Action -eq 'Restore') {
            try { $g.Result.Status = 'Warning'; $g.Result.Error = 'Stopped by the user'; Update-MclResultCounts $g.Result; Save-MclGuiReport -Settings $cfg; Update-MclGuiRows }
            catch { Write-MclLog 'WARN' "Report after a stop: $($_.Exception.Message)" }
        }
    }
    catch {
        Add-MclGuiLine 'Fail' $_.Exception.Message
        Set-MclGuiStatus 'Failed - see the progress' 'Failed'
        if ($g.Result -and $g.Result.Action -eq 'Restore') { try { Update-MclGuiRows } catch { Write-MclLog 'WARN' "Rows after a failure: $($_.Exception.Message)" } }
    }
    finally {
        if ($secret) { $secret.Dispose() }
        Stop-MclGuiRun
    }
}

function Show-MclGui {
    <#
    .SYNOPSIS
        Opens the window. Default configuration: config\MeetingCleanup.config.psd1 of the tool folder.
    #>
    [CmdletBinding()]
    param([hashtable]$Configuration)

    if (-not $Configuration) { $Configuration = Import-MclConfiguration }
    $window = New-MclForm -Configuration $Configuration
    # Ctrl+C in the console would stop the command that owns the window: the window then could not run any
    # of its PowerShell handlers. Ctrl+C is ignored while the window is open.
    $previousCtrlC = $null
    try { if (-not [Console]::IsInputRedirected) { $previousCtrlC = [Console]::TreatControlCAsInput; [Console]::TreatControlCAsInput = $true } } catch { $previousCtrlC = $null }
    $previousQuiet = $script:Quiet
    try {
        # The console stays quiet: the window shows the progress (the log file still gets every line).
        $script:Quiet = $true
        [void]$window.Form.ShowDialog()
    }
    finally {
        $script:Quiet = $previousQuiet
        if ($null -ne $previousCtrlC) { try { [Console]::TreatControlCAsInput = $previousCtrlC } catch { } }
        $script:Gui = $null
    }
}
