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
    Export-MclReport): the progress shows the lines of the console. The work runs in a runspace of its own (the
    module loaded there when the window opens): the window always answers, its lines come through a queue read
    every 100 ms (Start-MclGuiWork, Step-MclGuiWork); Stop ends the run at the next call. The progress bar
    (step, part done, time left) and the taskbar button follow the run (Set-MclGuiProgress). The lists are
    ListView rows of compiled objects (src\MeetingCleanup.Native.cs), filled at once and sortable.
    Search is read-only and writes a report; the action button removes or cancels the meetings ticked,
    after a confirmation that says exactly what will happen.

    Closing never needs PowerShell code: Close is the cancel button of the window (Esc too) and the title-bar
    button is native; the only Closing handler is attached while a run is in progress. A PowerShell event
    handler fails ("The pipeline has been stopped") once the command that opened the window is stopped, so
    the window must close without one. Ctrl+C is ignored in the console while it is open.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.2.2
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
            <!-- ListView/GridView rather than DataGrid: less layout per row (measured), so the list scrolls with
                 thousands of meetings. Rows are compiled objects (MeetingCleanupNative.MeetingRow). -->
            <ListView x:Name="Meetings" Grid.Row="1" SelectionMode="Single" BorderThickness="0" Background="Transparent"
                      VirtualizingPanel.IsVirtualizing="True" VirtualizingPanel.VirtualizationMode="Recycling" VirtualizingPanel.ScrollUnit="Item"
                      ScrollViewer.HorizontalScrollBarVisibility="Disabled">
              <ListView.ItemContainerStyle>
                <Style TargetType="ListViewItem" BasedOn="{StaticResource {x:Static GridView.GridViewItemContainerStyleKey}}">
                  <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                  <Setter Property="Padding" Value="0,1"/>
                  <Setter Property="MinHeight" Value="30"/>
                </Style>
              </ListView.ItemContainerStyle>
              <ListView.View>
                <GridView AllowsColumnReorder="False">
                  <GridViewColumn Width="44">
                    <GridViewColumn.CellTemplate>
                      <DataTemplate>
                        <CheckBox IsChecked="{Binding Selected, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" IsEnabled="{Binding CanTick}" HorizontalAlignment="Center" MinWidth="0" Padding="0"/>
                      </DataTemplate>
                    </GridViewColumn.CellTemplate>
                  </GridViewColumn>
                  <GridViewColumn Header="Start" DisplayMemberBinding="{Binding Start}" Width="128"/>
                  <GridViewColumn Header="Subject" DisplayMemberBinding="{Binding Subject}" Width="240"/>
                  <GridViewColumn Header="Organizer" DisplayMemberBinding="{Binding Who}" Width="0"/>
                  <GridViewColumn Header="Kind" DisplayMemberBinding="{Binding Kind}" Width="60"/>
                  <GridViewColumn Header="Organizer copy" DisplayMemberBinding="{Binding OrganizerCopy}" Width="118"/>
                  <GridViewColumn Header="Copies" DisplayMemberBinding="{Binding CopiesText}" Width="90"/>
                  <GridViewColumn Header="Status" DisplayMemberBinding="{Binding Status}" Width="100"/>
                </GridView>
              </ListView.View>
            </ListView>
            <TextBlock x:Name="MeetingsEmpty" Grid.Row="1" Margin="4,48,4,0" TextWrapping="Wrap" HorizontalAlignment="Center" FontSize="13" Foreground="{DynamicResource TextFillColorTertiaryBrush}"
                       Text="Type the organizer, choose where to search, then Search. A search changes nothing."/>
            <TextBlock x:Name="CopiesTitle" Grid.Row="2" Margin="0,10,0,6" FontSize="12" FontWeight="SemiBold" TextTrimming="CharacterEllipsis" Foreground="{DynamicResource TextFillColorSecondaryBrush}" Text="Copies of the meeting selected"/>
            <ListView x:Name="Copies" Grid.Row="3" SelectionMode="Single" BorderThickness="0" Background="Transparent" FontSize="12"
                      VirtualizingPanel.IsVirtualizing="True" VirtualizingPanel.VirtualizationMode="Recycling" VirtualizingPanel.ScrollUnit="Item"
                      ScrollViewer.HorizontalScrollBarVisibility="Disabled">
              <ListView.ItemContainerStyle>
                <Style TargetType="ListViewItem" BasedOn="{StaticResource {x:Static GridView.GridViewItemContainerStyleKey}}">
                  <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                  <Setter Property="Padding" Value="0"/>
                  <Setter Property="MinHeight" Value="26"/>
                </Style>
              </ListView.ItemContainerStyle>
              <ListView.View>
                <GridView AllowsColumnReorder="False">
                  <GridViewColumn Header="Mailbox" DisplayMemberBinding="{Binding Mailbox}" Width="220"/>
                  <GridViewColumn Header="Role" DisplayMemberBinding="{Binding Role}" Width="84"/>
                  <GridViewColumn Header="Occurrence" DisplayMemberBinding="{Binding Occurrence}" Width="116"/>
                  <GridViewColumn Header="Found by" DisplayMemberBinding="{Binding Via}" Width="116"/>
                  <GridViewColumn Header="Result" DisplayMemberBinding="{Binding Result}" Width="106"/>
                  <GridViewColumn Header="Detail" DisplayMemberBinding="{Binding Detail}" Width="200"/>
                </GridView>
              </ListView.View>
            </ListView>
          </Grid>
        </Border>
        <GridSplitter Grid.Row="1" Height="6" HorizontalAlignment="Stretch" Background="Transparent" ResizeDirection="Rows"/>
        <Border Grid.Row="2" Style="{StaticResource MclCard}" Margin="0,6,0,12" Padding="18,10,18,10">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock Text="Progress" Style="{StaticResource MclCardTitle}" Margin="0,0,14,6"/>
              <!-- During a run: the step and what it counts, then the part done and the time left (Set-MclGuiProgress). -->
              <TextBlock x:Name="ProgressText" Grid.Column="1" Margin="0,2,12,6" FontSize="12" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
              <TextBlock x:Name="ProgressInfo" Grid.Column="2" Margin="0,2,0,6" FontSize="12" FontWeight="SemiBold" VerticalAlignment="Center" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
            </Grid>
            <ProgressBar x:Name="ProgressBar" Grid.Row="1" Height="4" Minimum="0" Maximum="1" Margin="0,0,0,8" Visibility="Collapsed"/>
            <ScrollViewer x:Name="LogScroll" Grid.Row="2" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
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
    # The mark of the row selected in the lists (Fluent): the accent of the report.
    & $set 'ListViewItemPillFillBrush' '#B11F4B' '#FD8EA1'
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
        'ProgressBar', 'ProgressText', 'ProgressInfo', 'LogScroll', 'Log', 'Actions', 'Search', 'Apply', 'ApplyIcon', 'ApplyText', 'Stop', 'Footer', 'OpenReport', 'OpenFolder', 'Close') {
        $controls[$name] = $window.FindName($name)
    }
    foreach ($b in 'Search', 'Apply') {
        if ($look.Fluent) { $controls[$b].SetResourceReference([Windows.FrameworkElement]::StyleProperty, 'AccentButtonStyle') }
        else { $controls[$b].SetResourceReference([Windows.Controls.Control]::BackgroundProperty, 'AccentFillColorDefaultBrush'); $controls[$b].Foreground = [Windows.Media.Brushes]::White }
    }
    $controls.Version.Text = "v$($script:ToolVersion)  " + [char]0x00B7 + '  Nicolas Fabert'
    # The button of the window in the taskbar shows the progress of a run too.
    $window.TaskbarItemInfo = [Windows.Shell.TaskbarItemInfo]::new()

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

    # Lists replaced in one go (one refresh), rows compiled (MeetingCleanupNative.MeetingRow / CopyRow).
    $meetings = [MeetingCleanupNative.BulkCollection]::new()
    $copies = [MeetingCleanupNative.BulkCollection]::new()
    $items = [Collections.ObjectModel.ObservableCollection[object]]::new()
    $controls.Meetings.ItemsSource = $meetings
    $controls.Copies.ItemsSource = $copies
    $controls.Log.ItemsSource = $items
    # The channel of the background run (Start-MclGuiWork): its lines, Stop (Cancel), Hold, the log file. Every key
    # the engine reads is there (a synchronized hashtable throws on a missing key under Set-StrictMode).
    $shared = [hashtable]::Synchronized(@{ Cancel = $false; Hold = $false; Queue = [Collections.Concurrent.ConcurrentQueue[string[]]]::new(); Log = $null; Sink = $null; Pump = $null })
    $timer = [Windows.Threading.DispatcherTimer]::new([Windows.Threading.DispatcherPriority]::Background)
    $timer.Interval = [TimeSpan]::FromMilliseconds(100)
    $timer.Add_Tick({ Step-MclGuiWork })
    $script:Gui = @{
        Form = $window; Controls = $controls; Configuration = $Configuration.Clone(); Settings = $null; Theme = $look
        Running = $false; Result = $null; Acted = $false; LastReport = $null; LastFolder = $null; LastAction = ''; RestoreSource = $null
        Rows = $meetings; CopyRows = $copies; Items = $items; Lines = [Collections.Generic.List[string]]::new()
        Shared = $shared; Timer = $timer; Job = $null; Runspace = $null
        # The progress of the run in course (Set-MclGuiProgress): its step, its start, the part done (-1: none yet).
        Progress = @{ Active = $false; Step = ''; Started = [datetime]::UtcNow; Fraction = -1.0; Stopping = $false }
        # Attached to Closing only while a run is in progress: closing then stops the run first.
        ClosingGuard = [ComponentModel.CancelEventHandler] {
            param($sender, $e)
            $e.Cancel = $true
            if ($script:Gui) { $script:Gui.Shared.Cancel = $true; Set-MclGuiProgress -Stopping }
            Add-MclGuiLine 'Warn' 'A run is in progress: it stops at the next Graph call, then the window can be closed.'
        }
    }
    Set-MclGuiStatus 'Ready' 'Ready'
    $controls.Footer.Text = "Reports: $($Configuration.OutputPath)"

    $controls.Search.Add_Click({ Invoke-MclGuiSearch })
    $controls.Apply.Add_Click({ Invoke-MclGuiApply })
    $controls.Stop.Add_Click({
            $g = $script:Gui
            if ($g -and $g.Running) {
                $g.Shared.Cancel = $true
                Set-MclGuiProgress -Stopping
                Add-MclGuiLine 'Warn' $(if ($g.Shared.Hold) { 'Stop requested: the meetings being re-created are finished first (created, sent, old copies removed), then the run stops.' } else { 'Stop requested: the run stops at the next Graph call.' })
            }
        })
    $controls.SelectAll.Add_Click({ Set-MclGuiSelection $true })
    $controls.SelectNone.Add_Click({ Set-MclGuiSelection $false })
    $controls.Meetings.Add_SelectionChanged({ Update-MclGuiCopies })
    # A box ticked or unticked in the list: the action button counts again. A column header: the list is sorted.
    $controls.Meetings.AddHandler([Windows.Controls.Primitives.ButtonBase]::ClickEvent, [Windows.RoutedEventHandler] { param($sender, $e) if ($e.OriginalSource -is [Windows.Controls.GridViewColumnHeader]) { Set-MclGuiSort $sender $e.OriginalSource } else { Update-MclGuiState } })
    $controls.Copies.AddHandler([Windows.Controls.Primitives.ButtonBase]::ClickEvent, [Windows.RoutedEventHandler] { param($sender, $e) if ($e.OriginalSource -is [Windows.Controls.GridViewColumnHeader]) { Set-MclGuiSort $sender $e.OriginalSource } })
    # GridView has no proportional width: the subject (the mailbox and the detail of a copy) take the width left.
    $controls.Meetings.Add_SizeChanged({ Update-MclGuiColumns })
    $controls.Copies.Add_SizeChanged({ Update-MclGuiColumns })
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
    $ticked = [MeetingCleanupNative.GuiRows]::CountSelected($g.Rows)
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
    # Compiled rows notify their box: no redraw of the whole list.
    [MeetingCleanupNative.GuiRows]::SetSelected($g.Rows, $Value)
    Update-MclGuiState
}

function Update-MclGuiCopies {
    <# The copies of the meeting selected in the list. #>
    $g = $script:Gui
    if (-not $g) { return }
    $row = $g.Controls.Meetings.SelectedItem
    if (-not $row) { $g.CopyRows.ReplaceAll($null); $g.Controls.CopiesTitle.Text = 'Copies of the meeting selected'; return }
    $g.CopyRows.ReplaceAll([MeetingCleanupNative.GuiRows]::ForCopies($row.Meeting.Copies))
    $g.Controls.CopiesTitle.Text = "Copies of '$($row.Subject)'  $([char]0x00B7)  organizer $($row.Meeting.Organizer)"
    Update-MclGuiState
}

function Set-MclGuiSort {
    <# A column header clicked: the list sorted by that column, ascending then descending. #>
    param($List, $Header)
    $column = $Header.Column
    if (-not $column -or -not $column.DisplayMemberBinding) { return }
    $path = $column.DisplayMemberBinding.Path.Path
    if ($path -eq 'CopiesText') { $path = 'Copies' }
    $view = [Windows.Data.CollectionViewSource]::GetDefaultView($List.ItemsSource)
    $direction = [ComponentModel.ListSortDirection]::Ascending
    if ($view.SortDescriptions.Count -and $view.SortDescriptions[0].PropertyName -eq $path -and $view.SortDescriptions[0].Direction -eq $direction) { $direction = [ComponentModel.ListSortDirection]::Descending }
    $view.SortDescriptions.Clear()
    $view.SortDescriptions.Add([ComponentModel.SortDescription]::new($path, $direction))
}

function Update-MclGuiColumns {
    <# GridView has no proportional width: the subject takes the width left (the mailbox and the detail for the copies). #>
    $g = $script:Gui
    if (-not $g) { return }
    $list = $g.Controls.Meetings
    if ($list.ActualWidth -gt 0) {
        $fixed = 0.0; $subject = $null; $organizer = $null
        foreach ($col in $list.View.Columns) {
            if ($col.Header -eq 'Subject') { $subject = $col } elseif ($col.Header -eq 'Organizer') { $organizer = $col } elseif (-not [double]::IsNaN($col.Width)) { $fixed += $col.Width }
        }
        $free = $list.ActualWidth - $fixed - 34
        # The Organizer column (meetings of several organizers) takes 40 % of the width left.
        if ($organizer -and $organizer.Width -gt 0) { $organizer.Width = [Math]::Max(90, [Math]::Floor($free * 0.4)); $free -= $organizer.Width }
        if ($subject) { $subject.Width = [Math]::Max(120, $free) }
    }
    $list = $g.Controls.Copies
    if ($list.ActualWidth -gt 0) {
        $fixed = 0.0; $wide = [Collections.Generic.List[object]]::new()
        foreach ($col in $list.View.Columns) { if ($col.Header -in 'Mailbox', 'Detail') { $wide.Add($col) } elseif (-not [double]::IsNaN($col.Width)) { $fixed += $col.Width } }
        # The mailbox 60 % of the width left, the detail 40 %.
        $free = [Math]::Max(240, $list.ActualWidth - $fixed - 34)
        foreach ($col in $wide) { $col.Width = [Math]::Floor($free * $(if ($col.Header -eq 'Mailbox') { 0.6 } else { 0.4 })) }
    }
}
function Add-MclGuiLine {
    <# One line of the progress: icon and colour of its status. 'Progress' updates the progress bar instead. #>
    param([string]$Status, [string]$Text, [switch]$NoScroll)

    $g = $script:Gui
    if (-not $g) { return }
    if ($Status -eq 'Progress') {
        # fraction|text|time left (Write-MclProgress), cut at the first and the last '|'; the time left may be absent.
        $fraction = 0.0; $label = $Text; $left = ''
        $first = $Text.IndexOf('|'); $last = $Text.LastIndexOf('|')
        if ($first -gt 0 -and [double]::TryParse($Text.Substring(0, $first), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$fraction)) {
            if ($last -gt $first) { $label = $Text.Substring($first + 1, $last - $first - 1); $left = $Text.Substring($last + 1) }
            else { $label = $Text.Substring($first + 1) }
        }
        Set-MclGuiProgress -Fraction $fraction -Text $label -Left $left
        return
    }
    $glyphs = @{ Step = 0xE76C; Ok = 0xE73E; Warn = 0xE7BA; Fail = 0xEA39; Info = 0xE946; Skip = 0xE72A }
    $colours = @{ Step = 'MclBrandText'; Ok = 'MclSuccess'; Warn = 'MclCaution'; Fail = 'MclCritical'; Info = 'TextFillColorSecondaryBrush'; Skip = 'TextFillColorTertiaryBrush' }
    $key = if ($glyphs.ContainsKey($Status)) { $Status } else { 'Info' }
    $step = $Status -eq 'Step'
    if ($step) { Set-MclGuiProgress -Step $Text }
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
    if (-not $NoScroll) { $g.Controls.LogScroll.ScrollToEnd() }
}

function Set-MclGuiProgress {
    <#
        The progress bar of the window and its button in the taskbar, during a run:
          -Start <text>  the run begins: the bar moves (nothing counted yet), the time since the start on the right;
          -Step <text>   a step begins ('[3/6] Title'): the same, with the step;
          -Fraction      the part done, its text (1,240/1,858 mailboxes searched) and the time left (Write-MclProgress);
          -Tick          the timer of the run (Step-MclGuiWork): the time since the start while nothing is counted;
          -Stopping      Stop requested: the taskbar button turns yellow, 'Stopping...';
          -Waiting       a question to the administrator (the plan of a transfer or a restore): the bar stops;
          -Done          the run is over: bar and texts hidden, taskbar button back to normal.
        The window reads the lines of a run every 100 ms and keeps the last part done only: ten updates a second at
        most, whatever the volume.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Tick')]
    param(
        [Parameter(ParameterSetName = 'Start', Mandatory = $true)][string]$Start,
        [Parameter(ParameterSetName = 'Step', Mandatory = $true)][string]$Step,
        [Parameter(ParameterSetName = 'Fraction', Mandatory = $true)][double]$Fraction,
        [Parameter(ParameterSetName = 'Fraction')][AllowEmptyString()][string]$Text,
        [Parameter(ParameterSetName = 'Fraction')][AllowEmptyString()][string]$Left,
        [Parameter(ParameterSetName = 'Tick')][switch]$Tick,
        [Parameter(ParameterSetName = 'Stopping', Mandatory = $true)][switch]$Stopping,
        [Parameter(ParameterSetName = 'Waiting', Mandatory = $true)][switch]$Waiting,
        [Parameter(ParameterSetName = 'Done', Mandatory = $true)][switch]$Done
    )

    $g = $script:Gui
    if (-not $g) { return }
    $c = $g.Controls; $p = $g.Progress; $task = $g.Form.TaskbarItemInfo
    $dot = [char]0x00B7
    switch ($PSCmdlet.ParameterSetName) {
        'Start' {
            $g.Progress = $p = @{ Active = $true; Step = $Start; Started = [datetime]::UtcNow; Fraction = -1.0; Stopping = $false }
            $c.ProgressBar.Visibility = 'Visible'
        }
        'Done' {
            $p.Active = $false
            $c.ProgressBar.IsIndeterminate = $false
            $c.ProgressBar.Visibility = 'Collapsed'
            $c.ProgressText.Text = ''; $c.ProgressInfo.Text = ''
            if ($task) { $task.ProgressState = 'None' }
            return
        }
    }
    if (-not $p.Active) { return }
    switch ($PSCmdlet.ParameterSetName) {
        'Step' { $p.Step = 'Step ' + ($Step -replace '^\[(\d+/\d+)\]\s*', ('$1 ' + $dot + ' ')); $p.Fraction = -1.0; $c.ProgressBar.Visibility = 'Visible' }
        'Stopping' { $p.Stopping = $true }
        'Waiting' {
            # -2: nothing moves until the next -Start or -Step.
            $p.Fraction = -2.0
            $c.ProgressBar.IsIndeterminate = $false
            $c.ProgressBar.Visibility = 'Collapsed'
            $c.ProgressText.Text = 'Waiting for your answer'; $c.ProgressInfo.Text = ''
            if ($task) { $task.ProgressState = 'None' }
            return
        }
        'Fraction' {
            $p.Fraction = [Math]::Min(1.0, [Math]::Max(0.0, $Fraction))
            $c.ProgressBar.IsIndeterminate = $false
            $c.ProgressBar.Value = $p.Fraction
            $c.ProgressText.Text = if ($Text) { "$($p.Step)  $dot  $Text" } else { $p.Step }
            $percent = [string]::Format([Globalization.CultureInfo]::InvariantCulture, '{0:0} %', [Math]::Floor($p.Fraction * 100))
            $c.ProgressInfo.Text = if ($p.Stopping) { 'Stopping...' } elseif ($Left) { "$percent  $dot  $Left" } else { $percent }
            if ($task) { $task.ProgressState = $(if ($p.Stopping) { 'Paused' } else { 'Normal' }); $task.ProgressValue = $p.Fraction }
            return
        }
    }
    if ($p.Fraction -ge 0 -or $p.Fraction -le -2) {
        # A part is known (the bar stays where it is), or a question is asked; Stop turns the taskbar button yellow.
        if ($p.Stopping -and $p.Fraction -ge 0) { $c.ProgressInfo.Text = 'Stopping...'; if ($task) { $task.ProgressState = 'Paused' } }
        return
    }
    # Nothing counted yet in this step: the bar moves, the time since the start of the run.
    if (-not $c.ProgressBar.IsIndeterminate) { $c.ProgressBar.IsIndeterminate = $true }
    if ($c.ProgressText.Text -ne $p.Step) { $c.ProgressText.Text = $p.Step }
    $info = if ($p.Stopping) { 'Stopping...' } else {
        $t = [datetime]::UtcNow - $p.Started
        if ($t.TotalHours -ge 1) { '{0}:{1:00}:{2:00} elapsed' -f [int][Math]::Floor($t.TotalHours), $t.Minutes, $t.Seconds } else { '{0}:{1:00} elapsed' -f $t.Minutes, $t.Seconds }
    }
    if ($c.ProgressInfo.Text -ne $info) { $c.ProgressInfo.Text = $info }
    if ($task) {
        $state = if ($p.Stopping) { 'Paused' } else { 'Indeterminate' }
        if ([string]$task.ProgressState -ne $state) { $task.ProgressState = $state; if ($p.Stopping) { $task.ProgressValue = 1 } }
    }
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
    <# The list of the meetings from the result (after a search or an action), replaced in one go. #>
    $g = $script:Gui
    $meetings = if ($null -ne $g.Result) { $g.Result.Meetings } else { $null }
    [Windows.Data.CollectionViewSource]::GetDefaultView($g.Rows).SortDescriptions.Clear()
    $g.Rows.ReplaceAll([MeetingCleanupNative.GuiRows]::ForMeetings($meetings, [bool]$g.Acted))
    # The Organizer column only when the meetings come from more than one organizer.
    $several = [MeetingCleanupNative.GuiRows]::OrganizerCount($meetings) -gt 1
    foreach ($column in $g.Controls.Meetings.View.Columns) { if ($column.Header -eq 'Organizer') { $column.Width = if ($several) { 150 } else { 0 } } }
    Update-MclGuiColumns
    $g.Controls.MeetingsEmpty.Visibility = if ($g.Rows.Count) { 'Collapsed' } else { 'Visible' }
    if (-not $g.Rows.Count) { $g.Controls.MeetingsEmpty.Text = 'No meeting found: widen the period, check the address, or search in more mailboxes.' }
    if ($g.Rows.Count) { $g.Controls.Meetings.SelectedIndex = 0 }
    Update-MclGuiCopies
    Update-MclGuiState
}
function Start-MclGuiRun {
    <# A run starts: buttons disabled, Stop enabled, a closing of the window stops the run first. #>
    param([string]$Text)
    $g = $script:Gui
    $c = $g.Controls
    $g.Shared.Cancel = $false
    $g.Shared.Hold = $false
    $g.Shared.Log = $script:LogWriter
    $g.Running = $true
    $g.Form.add_Closing($g.ClosingGuard)
    foreach ($b in 'Search', 'Apply', 'Restore', 'OpenReport', 'OpenFolder', 'Close', 'SelectAll', 'SelectNone', 'Inputs', 'Meetings') { $c[$b].IsEnabled = $false }
    $c.Stop.IsEnabled = $true
    Set-MclGuiStatus $Text 'Running'
    Set-MclGuiProgress -Start $Text
}

function Stop-MclGuiRun {
    $g = $script:Gui
    $c = $g.Controls
    $g.Form.remove_Closing($g.ClosingGuard)
    $g.Running = $false
    foreach ($b in 'Search', 'Restore', 'Close', 'SelectAll', 'SelectNone', 'Inputs', 'Meetings') { $c[$b].IsEnabled = $true }
    $c.Stop.IsEnabled = $false
    Set-MclGuiProgress -Done
    $c.OpenReport.IsEnabled = [bool]$g.LastReport
    $c.OpenFolder.IsEnabled = [bool]$g.LastFolder
    Update-MclGuiState
}

function Set-MclGuiReport {
    <# The report of the run just done: Open the report / Open the folder, and the footer. #>
    param($Report)
    if (-not $Report) { return }
    $g = $script:Gui
    $g.LastFolder = $Report.Directory
    $g.LastReport = $Report.Html
    $g.Controls.Footer.Text = "Report: $($Report.Directory)"
}

#region Background work ---------------------------------------------------------------------------------------
# A search or an action of the window runs in a runspace of its own, with the module loaded there: the window
# keeps answering whatever the volume (thousands of copies to read, compare, write in the report). Its lines go
# through a queue (Shared.Queue) that the window reads every 100 ms; Stop and Hold go through Shared too.

$script:GuiInline = $false
$script:GuiWorkScript = @'
param($Kind, $Arguments, $Shared)
& (Get-Module MeetingCleanup) { param($Kind, $Arguments, $Shared) Invoke-MclGuiWork -Kind $Kind -Arguments $Arguments -Shared $Shared } $Kind $Arguments $Shared
'@

function Save-MclRunReport {
    <# The report of a run of the window (Export-MclReport); returns its folder and its HTML file. #>
    param([Parameter(Mandatory = $true)][hashtable]$Settings, [Parameter(Mandatory = $true)][pscustomobject]$Result, [string]$Directory)
    Write-MclNextStep 'Report' 'Report'
    $exportArgs = @{ Result = $Result; OutputPath = $Settings.OutputPath; Prefix = $Settings.ReportPrefix; Formats = $Settings.ReportFormats; Delimiter = $Settings.CsvDelimiter }
    if ($Directory) { $exportArgs.Directory = $Directory }
    $report = Export-MclReport @exportArgs
    Write-MclItem Ok "Report: $($report.Directory)" -Icon File
    return @{ Directory = $report.Directory; Html = Get-MclProperty $report.Files 'Html' }
}

function Invoke-MclGuiWork {
    <#
    .SYNOPSIS
        One piece of work of the window, in its background runspace (or inline): Search, Cleanup, TransferPlan,
        Transfer, RestorePlan, Restore. Never throws: returns @{ Ok; Cancelled; Started; Error; Result; Report; ... }.
    .NOTES
        Started: the action has begun (Graph connected, something may have changed) - the window then never offers
        the meetings for an action again.
    #>
    param([Parameter(Mandatory = $true)][string]$Kind, [hashtable]$Arguments = @{}, [Parameter(Mandatory = $true)][hashtable]$Shared)
    $script:Ui = $Shared
    $script:Quiet = $true
    if ($Shared.Log -and -not [object]::ReferenceEquals($script:LogWriter, $Shared.Log)) { $script:LogWriter = $Shared.Log }
    $a = $Arguments
    $out = @{ Ok = $false; Cancelled = $false; Started = $false; Error = ''; Result = $null; Report = $null }
    $runPath = $null
    try {
        switch ($Kind) {
            'Search' {
                Initialize-MclSteps -Total 6
                Write-MclNextStep 'Microsoft Graph' 'Key'
                $connection = Connect-MclGraph -Settings $a.Settings -Secret $a.Secret -Action 'Report'
                Write-MclItem Ok ('Application {0} {1} tenant {2}' -f $(if ($connection.AppName) { $connection.AppName } else { $a.Settings.AppId }), [char]0x00B7, $connection.TenantGuid) -Icon Key
                if (-not $connection.CanWrite) { Write-MclItem Warn 'Calendars.Read only: the meetings can be listed, not removed or cancelled.' }
                $out.Result = Find-MclMeetings -Settings $a.Settings -Request $a.Request
                $out.Report = Save-MclRunReport -Settings $a.Settings -Result $out.Result
            }
            'Cleanup' {
                Initialize-MclSteps -Total (2 + [int][bool]$a.Settings.Verify)
                $null = Connect-MclGraph -Settings $a.Settings -Secret $a.Secret -Action $a.Action
                $out.Started = $true
                $out.Result = $a.Result
                $runPath = New-MclRunFolder -OutputPath $a.Settings.OutputPath -Prefix $a.Settings.ReportPrefix -Action $a.Action
                $out.Result = Invoke-MclCleanup -Settings $a.Settings -Result $a.Result -Action $a.Action -Comment $a.Comment -BackupPath (Join-Path $runPath "$($a.Settings.ReportPrefix)-Backup.json")
                $out.Report = Save-MclRunReport -Settings $a.Settings -Result $out.Result -Directory $runPath
            }
            'TransferPlan' {
                Initialize-MclSteps -Total (3 + [int][bool]$a.Settings.Verify)
                $null = Connect-MclGraph -Settings $a.Settings -Secret $a.Secret -Action 'Remove'
                $out.New = Resolve-MclNewOrganizer -Address $a.Address
                $out.Plan = Get-MclTransferPlan -Result $a.Result -NewOrganizer $out.New -Method $a.Method -Comment $a.Comment
            }
            'Transfer' {
                $out.Started = $true
                $out.Result = $a.Result
                $runPath = New-MclRunFolder -OutputPath $a.Settings.OutputPath -Prefix $a.Settings.ReportPrefix -Action 'Transfer'
                Write-MclNextStep 'Exchange Online PowerShell' 'Server'
                if ($a.Plan.Native.Count) { Connect-MclExchange -Settings $a.Settings -Secret $a.Secret -For Transfer } else { Write-MclItem Skip 'Not needed: every meeting is re-created with Microsoft Graph.' }
                try { $out.Result = Invoke-MclTransfer -Settings $a.Settings -Result $a.Result -Plan $a.Plan -Comment $a.Comment -BackupPath (Join-Path $runPath "$($a.Settings.ReportPrefix)-Backup.json") }
                finally { if ($a.Plan.Native.Count) { Disconnect-MclExchange } }
                $out.Report = Save-MclRunReport -Settings $a.Settings -Result $out.Result -Directory $runPath
            }
            'RestorePlan' {
                $out.Source = Import-MclRestoreSource -Path $a.Path
                $out.Plan = Get-MclRestorePlan -Result $out.Source
            }
            'Restore' {
                Initialize-MclSteps -Total 5
                Write-MclNextStep 'Microsoft Graph' 'Key'
                $connection = Connect-MclGraph -Settings $a.Settings -Secret $a.Secret -Action 'Remove'
                if ($a.Source.Tenant -and $a.Source.Tenant -ne $connection.TenantGuid) { throw "The report belongs to tenant $($a.Source.Tenant), the application signs in to $($connection.TenantGuid)." }
                Write-MclNextStep 'Exchange Online PowerShell' 'Server'
                Connect-MclExchange -Settings $a.Settings -Secret $a.Secret
                $out.Started = $true
                $out.Result = $a.Source
                try { $out.Result = Invoke-MclRestore -Settings $a.Settings -Result $a.Source }
                finally { Disconnect-MclExchange }
                $out.Report = Save-MclRunReport -Settings $a.Settings -Result $out.Result
            }
            default { throw "Unknown work of the window: $Kind" }
        }
        $out.Ok = $true
    }
    catch [OperationCanceledException] {
        $out.Cancelled = $true
        # Stopped during an action: what was done is in the report (it can be restored, or finished by a new run).
        if ($out.Started -and $out.Result) {
            try {
                $out.Result.Status = 'Warning'; $out.Result.Error = 'Stopped by the user'
                Update-MclResultCounts $out.Result
                $out.Report = Save-MclRunReport -Settings $a.Settings -Result $out.Result -Directory $runPath
            }
            catch { Write-MclLog 'WARN' "Report after a stop: $($_.Exception.Message)" }
        }
    }
    catch {
        $out.Error = $_.Exception.Message
        Write-MclLog 'ERROR' "Window ($Kind): $($_.Exception.Message)"
    }
    finally { $script:Ui = $null }
    return $out
}

function Open-MclGuiRunspace {
    <# The background runspace of the window, opened in the background (module loaded there) when the window opens. #>
    $g = $script:Gui
    if ($g.Runspace) { return }
    $iss = [Management.Automation.Runspaces.InitialSessionState]::CreateDefault2()
    $iss.ImportPSModule([string[]]@(Join-Path $script:ToolRoot 'MeetingCleanup.psd1'))
    $runspace = [Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
    # STA: an interactive sign-in of Exchange Online (restore as an administrator) needs it.
    $runspace.ApartmentState = [Threading.ApartmentState]::STA
    $runspace.ThreadOptions = [Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $runspace.OpenAsync()
    $g.Runspace = $runspace
}

function Get-MclGuiRunspace {
    <# The background runspace, once open and free (the module is loaded while it is still Busy after Opened). #>
    $g = $script:Gui
    if (-not $g.Runspace) { Open-MclGuiRunspace }
    $until = [datetime]::UtcNow.AddSeconds(90)
    while (($g.Runspace.RunspaceStateInfo.State -in 'BeforeOpen', 'Opening' -or ($g.Runspace.RunspaceStateInfo.State -eq 'Opened' -and $g.Runspace.RunspaceAvailability -ne 'Available')) -and [datetime]::UtcNow -lt $until) {
        Invoke-MclGuiPump; Start-Sleep -Milliseconds 30
    }
    if ($g.Runspace.RunspaceStateInfo.State -ne 'Opened') { throw "The engine of the window could not start: $($g.Runspace.RunspaceStateInfo.Reason)" }
    if ($g.Runspace.RunspaceAvailability -ne 'Available') { throw 'The engine of the window is still busy: try again in a moment.' }
    return $g.Runspace
}

function Start-MclGuiWork {
    <#
        Runs one piece of work (Invoke-MclGuiWork) in the background runspace of the window and returns at once;
        OnDone runs on the window thread with the outcome and the Context. Inline ($script:GuiInline: the tests and
        the documentation tool, whose simulated tenant lives in this runspace): the same work, on this thread.
    #>
    param([Parameter(Mandatory = $true)][string]$Kind, [hashtable]$Arguments = @{}, [Parameter(Mandatory = $true)][scriptblock]$OnDone, [hashtable]$Context = @{})
    $g = $script:Gui
    if ($script:GuiInline) {
        $outcome = Invoke-MclGuiWork -Kind $Kind -Arguments $Arguments -Shared $g.Shared
        Receive-MclGuiMessages
        & $OnDone $outcome $Context
        return
    }
    $ps = [PowerShell]::Create()
    $ps.Runspace = Get-MclGuiRunspace
    [void]$ps.AddScript($script:GuiWorkScript).AddArgument($Kind).AddArgument($Arguments).AddArgument($g.Shared)
    $g.Job = @{ PowerShell = $ps; Handle = $ps.BeginInvoke(); OnDone = $OnDone; Context = $Context; Kind = $Kind }
    $g.Timer.Start()
}

function Receive-MclGuiMessages {
    <# The lines of the background run since the last look: added to the progress; the bar shows the last state. #>
    $g = $script:Gui
    $item = $null; $progress = $null; $added = $false
    while ($g.Shared.Queue.TryDequeue([ref]$item)) {
        if ($item[0] -eq 'Progress') { $progress = $item[1]; continue }
        if ($item[0] -eq 'Step') { $progress = $null }
        Add-MclGuiLine $item[0] $item[1] -NoScroll
        $added = $true
    }
    if ($progress) { Add-MclGuiLine 'Progress' $progress }
    if ($added) { $g.Controls.LogScroll.ScrollToEnd() }
}

function Step-MclGuiWork {
    <# Every 100 ms while a background run is in progress: its lines, then its end (the outcome to OnDone). #>
    $g = $script:Gui
    if (-not $g) { return }
    Receive-MclGuiMessages
    Set-MclGuiProgress -Tick
    $job = $g.Job
    if (-not $job -or -not $job.Handle.IsCompleted) { return }
    $g.Job = $null
    $g.Timer.Stop()
    $outcome = $null
    try {
        $output = $job.PowerShell.EndInvoke($job.Handle)
        if ($output.Count) { $outcome = $output[$output.Count - 1]; if ($null -ne $outcome) { $outcome = $outcome.psobject.BaseObject } }
        if ($outcome -isnot [hashtable]) {
            $err = @($job.PowerShell.Streams.Error) | Select-Object -First 1
            $outcome = @{ Ok = $false; Cancelled = $false; Started = $false; Error = $(if ($err) { [string]$err } else { 'The background run ended without a result.' }); Result = $null; Report = $null }
        }
    }
    catch {
        $inner = if ($_.Exception.InnerException) { $_.Exception.InnerException.Message } else { $_.Exception.Message }
        $outcome = @{ Ok = $false; Cancelled = $false; Started = $false; Error = $inner; Result = $null; Report = $null }
    }
    finally { $job.PowerShell.Dispose() }
    Receive-MclGuiMessages
    try { & $job.OnDone $outcome $job.Context }
    catch {
        Add-MclGuiLine 'Fail' $_.Exception.Message
        Set-MclGuiStatus 'Failed - see the progress' 'Failed'
        if ($g.Running -and -not $g.Job) { Stop-MclGuiRun }
    }
}

function Wait-MclGuiWork {
    <# Lab tests and tools: waits (the window answering) until the run of the window is over, with what follows it. #>
    param([int]$TimeoutSeconds = 3600)
    $until = [datetime]::UtcNow.AddSeconds($TimeoutSeconds)
    while ($script:Gui -and ($script:Gui.Job -or $script:Gui.Running) -and [datetime]::UtcNow -lt $until) { Invoke-MclGuiPump; Start-Sleep -Milliseconds 40 }
}
#endregion

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
    if ($cfg.AuthMode -eq 'ClientSecret' -and $c.Secret.SecurePassword.Length) { $secret = $c.Secret.SecurePassword.Copy() }
    Start-MclGuiRun 'Searching...'
    $g.Result = $null; $g.Acted = $false; $g.LastReport = $null; $g.LastFolder = $null
    $g.Rows.ReplaceAll($null); $g.CopyRows.ReplaceAll($null)
    Write-MclLog 'STEP' "Window search: $(if ($rooms) { "rooms $($request.Room -join ', ')" } else { $request.Organizer -join ', ' }) from $($request.Start.ToString('o')) to $($request.End.ToString('o'))$(if (-not $rooms) { " in $($scopes -join ', ')" })"
    try {
        Start-MclGuiWork -Kind 'Search' -Arguments @{ Settings = $cfg; Request = $request; Secret = $secret } -Context @{ Settings = $cfg; Secret = $secret } -OnDone {
            param($outcome, $context)
            $g = $script:Gui
            try {
                if ($outcome.Ok) {
                    $g.Result = $outcome.Result
                    $g.Settings = $context.Settings
                    Set-MclGuiReport $outcome.Report
                    Update-MclGuiRows
                    $n = $g.Result.Counts
                    Set-MclGuiStatus ('{0} meeting(s) {1} {2} copies' -f $n.Meetings, [char]0x00B7, $n.Copies) $g.Result.Status
                    if ($n.Meetings) { Add-MclGuiLine 'Info' 'Nothing has been changed. Untick the meetings to keep, choose the action on the left, then use the action button.' }
                }
                elseif ($outcome.Cancelled) { Add-MclGuiLine 'Warn' 'Search stopped: nothing was changed.'; Set-MclGuiStatus 'Stopped' 'Warning' }
                else { Add-MclGuiLine 'Fail' $outcome.Error; Set-MclGuiStatus 'Failed - see the progress' 'Failed' }
            }
            finally {
                if ($context.Secret) { $context.Secret.Dispose() }
                Stop-MclGuiRun
            }
        }
    }
    catch {
        Add-MclGuiLine 'Fail' $_.Exception.Message
        Set-MclGuiStatus 'Failed - see the progress' 'Failed'
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

function Complete-MclGuiAction {
    <#
        End of an action of the window (Remove, Cancel, Transfer, Restore): the result, its report and the list,
        the status; after a stop, what was done (in the report).
    #>
    param([hashtable]$Outcome, [string]$Action, [string]$StoppedText)
    $g = $script:Gui
    if ($Outcome.Started) { $g.Acted = $true }
    if ($Outcome.Result) { $g.Result = $Outcome.Result }
    Set-MclGuiReport $Outcome.Report
    if ($Outcome.Ok -or ($Outcome.Cancelled -and $Outcome.Report)) { $g.LastAction = $Action }
    if ($Outcome.Started -and $g.Result) { try { Update-MclGuiRows } catch { Write-MclLog 'WARN' "Rows after the run: $($_.Exception.Message)" } }
    if ($Outcome.Cancelled) { Add-MclGuiLine 'Warn' $StoppedText; Set-MclGuiStatus 'Stopped' 'Warning'; return $false }
    if (-not $Outcome.Ok) { Add-MclGuiLine 'Fail' $Outcome.Error; Set-MclGuiStatus 'Failed - see the progress' 'Failed'; return $false }
    return $true
}

function Invoke-MclGuiApply {
    $g = $script:Gui
    $c = $g.Controls
    if ($null -eq $g.Result -or $g.Acted) { return }
    [MeetingCleanupNative.GuiRows]::ApplySelection($g.Rows)
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

    $cfg = $g.Settings
    $secret = $null
    if ($cfg.AuthMode -eq 'ClientSecret' -and $c.Secret.SecurePassword.Length) { $secret = $c.Secret.SecurePassword.Copy() }
    Start-MclGuiRun $(if ($action -eq 'Cancel') { 'Cancelling...' } else { 'Removing...' })
    Write-MclLog 'INFO' "Confirmed in the window by $([Environment]::UserName): $($plan.Text)"
    try {
        Start-MclGuiWork -Kind 'Cleanup' -Arguments @{ Settings = $cfg; Secret = $secret; Result = $g.Result; Action = $action; Comment = $comment } -Context @{ Action = $action; Secret = $secret } -OnDone {
            param($outcome, $context)
            $g = $script:Gui
            try {
                if (Complete-MclGuiAction $outcome $context.Action 'Stopped: the copies already handled are in the report (they can be restored), the others were left as they were.') {
                    $n = $g.Result.Counts
                    Set-MclGuiStatus ('{0} {1} {2} removed {1} {3} cancelled {1} {4} failed' -f $g.Result.Status, [char]0x00B7, $n.Removed, $n.Cancelled, $n.Failed) $g.Result.Status
                    Add-MclGuiLine 'Info' $(if ($context.Action -eq 'Remove') { 'Search again to see what is left. To undo it: Restore... (the copies come back, no message).' } else { 'Search again to see what is left, or to clean other meetings.' })
                }
            }
            finally {
                if ($context.Secret) { $context.Secret.Dispose() }
                Stop-MclGuiRun
            }
        }
    }
    catch {
        Add-MclGuiLine 'Fail' $_.Exception.Message
        Set-MclGuiStatus 'Failed - see the progress' 'Failed'
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
    $cfg = $g.Settings
    $secret = $null
    if ($cfg.AuthMode -eq 'ClientSecret' -and $c.Secret.SecurePassword.Length) { $secret = $c.Secret.SecurePassword.Copy() }
    $context = @{ Settings = $cfg; Secret = $secret }
    Start-MclGuiRun 'Transferring...'
    try {
        # 1. the plan (Graph: the new organizer, the meetings), 2. the confirmation, 3. the transfer
        Start-MclGuiWork -Kind 'TransferPlan' -Arguments @{ Settings = $cfg; Secret = $secret; Result = $g.Result; Address = $address; Method = [string]$c.TransferMethod.SelectedItem; Comment = [string]$cfg.TransferComment } -Context $context -OnDone {
            param($outcome, $context)
            $g = $script:Gui
            $next = $false
            try {
                if ($outcome.Cancelled) { Add-MclGuiLine 'Warn' 'Stopped: nothing was changed.'; Set-MclGuiStatus 'Stopped' 'Warning'; return }
                if (-not $outcome.Ok) { Add-MclGuiLine 'Fail' $outcome.Error; Set-MclGuiStatus 'Failed - see the progress' 'Failed'; return }
                $plan = $outcome.Plan; $new = $outcome.New
                if (-not ($plan.Native.Count + $plan.Recreate.Count)) { foreach ($line in $plan.Lines) { Add-MclGuiLine 'Warn' $line }; Set-MclGuiStatus 'Nothing to transfer' 'Warning'; return }
                $text = "Transfer to $(if ($new.Name) { "$($new.Name) <$($new.Address)>" } else { $new.Address }):`n`n - " + ($plan.Lines -join "`n - ") + "`n`nA backup is written first. Continue?"
                Set-MclGuiProgress -Waiting
                if ((Show-MclGuiQuestion -Text $text -Title 'Meeting Cleanup - Transfer' -Image Warning) -ne [Windows.MessageBoxResult]::Yes) { Add-MclGuiLine 'Info' 'Nothing was changed.'; return }
                Write-MclLog 'INFO' "Transfer confirmed in the window by $([Environment]::UserName): $($plan.Text)"
                Set-MclGuiProgress -Start 'Transferring...'
                $next = $true
                Start-MclGuiWork -Kind 'Transfer' -Arguments @{ Settings = $context.Settings; Secret = $context.Secret; Result = $g.Result; Plan = $plan; Comment = [string]$context.Settings.TransferComment } -Context $context -OnDone {
                    param($outcome, $context)
                    $g = $script:Gui
                    try {
                        if (Complete-MclGuiAction $outcome 'Transfer' 'Stopped: the meetings already transferred are in the report.') {
                            $n = $g.Result.Counts
                            $failed = 0; foreach ($m in $g.Result.Meetings) { if ($m.Status -eq 'Failed') { $failed++ } }
                            Set-MclGuiStatus ('{0} {1} {2} transferred {1} {3} failed' -f $g.Result.Status, [char]0x00B7, $n.Transferred, $failed) $g.Result.Status
                        }
                    }
                    finally {
                        if ($context.Secret) { $context.Secret.Dispose() }
                        Stop-MclGuiRun
                    }
                }
            }
            catch { $next = $false; Add-MclGuiLine 'Fail' $_.Exception.Message; Set-MclGuiStatus 'Failed - see the progress' 'Failed' }
            finally {
                if (-not $next) {
                    if ($context.Secret) { $context.Secret.Dispose() }
                    Stop-MclGuiRun
                }
            }
        }
    }
    catch {
        Add-MclGuiLine 'Fail' $_.Exception.Message
        Set-MclGuiStatus 'Failed - see the progress' 'Failed'
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
    <# Restore of a Remove run: the report read, the plan, the confirmation, Exchange Online PowerShell, Recoverable Items, the check, the report. #>
    $g = $script:Gui
    $c = $g.Controls
    $folder = Select-MclGuiRestoreFolder
    if (-not $folder) { return }
    $g.Items.Clear(); $g.Lines.Clear()
    $cfg = if ($g.Settings) { $g.Settings } else { Get-MclGuiSettings }
    $problems = @((Test-MclConfiguration -Configuration $cfg -ForConnection).Problems)
    if ($problems.Count) { foreach ($p in $problems) { Add-MclGuiLine 'Fail' "$p (Connection, at the bottom left)" }; Set-MclGuiStatus 'Fix the connection' 'Failed'; return }
    $secret = $null
    if ($cfg.AuthMode -eq 'ClientSecret' -and $c.Secret.SecurePassword.Length) { $secret = $c.Secret.SecurePassword.Copy() }
    $context = @{ Settings = $cfg; Secret = $secret; Folder = $folder }
    Start-MclGuiRun 'Reading the report...'
    try {
        # 1. the report read and the plan, 2. the confirmation, 3. the restore
        Start-MclGuiWork -Kind 'RestorePlan' -Arguments @{ Path = $folder } -Context $context -OnDone {
            param($outcome, $context)
            $g = $script:Gui
            $next = $false
            try {
                if ($outcome.Cancelled) { Add-MclGuiLine 'Warn' 'Stopped: nothing was changed.'; Set-MclGuiStatus 'Stopped' 'Warning'; return }
                if (-not $outcome.Ok) { Add-MclGuiLine 'Fail' $outcome.Error; Set-MclGuiStatus 'Not a report to restore' 'Failed'; return }
                $source = $outcome.Source; $plan = $outcome.Plan; $cfg = $context.Settings
                if (-not $plan.Restore.Count) {
                    foreach ($line in $plan.Lines) { Add-MclGuiLine 'Info' $line }
                    Add-MclGuiLine 'Warn' 'Nothing to restore in this run.'
                    Set-MclGuiStatus 'Nothing to restore' 'Warning'
                    return
                }
                $text = "Restore ($($source.SourceAction) run of $([IO.Path]::GetFileName((Split-Path $source.FromReport -Parent)))):`n`n - " + ($plan.Lines -join "`n - ")
                $text += "`n`nExchange Online PowerShell: $(if ($cfg.RestoreConnection -eq 'Interactive') { "as an administrator $($cfg.RestoreUser) (sign-in window)" } else { 'as the application (role Mailbox Import Export)' }). Continue?"
                Set-MclGuiProgress -Waiting
                if ((Show-MclGuiQuestion -Text $text -Title 'Meeting Cleanup - Restore') -ne [Windows.MessageBoxResult]::Yes) { Add-MclGuiLine 'Info' 'Nothing was changed.'; Set-MclGuiStatus 'Ready' 'Ready'; return }
                Write-MclLog 'INFO' "Restore confirmed in the window by $([Environment]::UserName): $($context.Folder) - $($plan.Text)"
                Set-MclGuiStatus 'Restoring...' 'Running'
                Set-MclGuiProgress -Start 'Restoring...'
                $next = $true
                Start-MclGuiWork -Kind 'Restore' -Arguments @{ Settings = $cfg; Secret = $context.Secret; Source = $source } -Context $context -OnDone {
                    param($outcome, $context)
                    $g = $script:Gui
                    try {
                        # From the start of the restore, the meetings shown are those of the restore: never ticked
                        # again for an action; Restore... offers this run again (a stopped restore is finished so).
                        if ($outcome.Started) { $g.Settings = $context.Settings; $g.LastAction = 'Restore'; $g.RestoreSource = $context.Folder }
                        if (Complete-MclGuiAction $outcome 'Restore' 'Stopped: the copies already restored are back; the others are still in Recoverable Items. Restore... again finishes them.') {
                            $n = $g.Result.Counts
                            Set-MclGuiStatus ('{0} {1} {2} restored {1} {3} not found {1} {4} failed' -f $g.Result.Status, [char]0x00B7, $n.Restored, $n.NotFound, $n.Failed) $g.Result.Status
                        }
                    }
                    finally {
                        if ($context.Secret) { $context.Secret.Dispose() }
                        Stop-MclGuiRun
                    }
                }
            }
            catch { $next = $false; Add-MclGuiLine 'Fail' $_.Exception.Message; Set-MclGuiStatus 'Failed - see the progress' 'Failed' }
            finally {
                if (-not $next) {
                    if ($context.Secret) { $context.Secret.Dispose() }
                    Stop-MclGuiRun
                }
            }
        }
    }
    catch {
        Add-MclGuiLine 'Fail' $_.Exception.Message
        Set-MclGuiStatus 'Failed - see the progress' 'Failed'
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
    # The engine of the window starts loading now, in the background: ready by the first search.
    Open-MclGuiRunspace
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
        if ($script:Gui) {
            $script:Gui.Timer.Stop()
            if ($script:Gui.Runspace) { try { $script:Gui.Runspace.Dispose() } catch { Write-MclLog 'WARN' "Engine of the window: $($_.Exception.Message)" } }
        }
        $script:Gui = $null
    }
}