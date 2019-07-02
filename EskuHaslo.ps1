$ErrorActionPreference= 'silentlycontinue'
Add-Type -Path "C:\Eskulap\orant\BIN\OdtPrinting\Oracle.ManagedDataAccess.dll"
Add-Type -Path "C:\bg\MySQL.Data.dll"
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing

$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$AssemblyLocation = Join-Path -Path $ScriptPath -ChildPath .\themes
foreach ($Assembly in (Get-ChildItem $AssemblyLocation -Filter *.dll)) {
    [System.Reflection.Assembly]::LoadFrom($Assembly.fullName) | out-null
}

[xml]$xaml = @"
<Controls:MetroWindow
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
        Title="Oracle - Zmiana hasła" 
        Height="500" 
        Width="400"
        ResizeMode="NoResize"
        BorderThickness="0" 
        GlowBrush="Blue"
        WindowStartupLocation="CenterScreen">

        <Window.Resources>
            <ResourceDictionary>
                <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Colors.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/Cobalt.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/BaseLight.xaml" />
				<ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.AnimatedTabControl.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/BaseLight.xaml" />
                </ResourceDictionary.MergedDictionaries>
            </ResourceDictionary>
        </Window.Resources>

        <Controls:MetroWindow.LeftWindowCommands>
            <Controls:WindowCommands>
                <Button>
                    <StackPanel Orientation="Horizontal">
                        <Image Name="Image2" HorizontalAlignment="Left" Height="20" Margin="0" VerticalAlignment="Top" Width="20"/>
                    </StackPanel>
                </Button>
            </Controls:WindowCommands>
        </Controls:MetroWindow.LeftWindowCommands>	

    <Grid>
        <Image Name="Image1" HorizontalAlignment="center" Margin="2" VerticalAlignment="Top" Width="120"/>
        <Rectangle HorizontalAlignment="Stretch" Fill="Blue" Height="2" Margin="0,180,0,0" VerticalAlignment="Top"/>

        <Button Name="b2" Content="Odblokuj" HorizontalAlignment="Left" Margin="60,186,0,0" VerticalAlignment="Top" Width="112"/>
        <Button Name="b1" Content="Hasło" HorizontalAlignment="Left" Margin="240,186,0,0" VerticalAlignment="Top" Width="112"/>

        <TextBox Name="TextBox2" HorizontalAlignment="center" HorizontalContentAlignment="Center" VerticalContentAlignment="Center" Margin="0,220,0,0" VerticalAlignment="Top" Width="380" Height="200"/>

        <StatusBar Name="statusBar1" Height="30" HorizontalAlignment="Stretch" VerticalAlignment="Bottom">
            <StatusBarItem HorizontalAlignment="Left">
                <Label Content="Przygotował Michał Zbyl" FontSize="11"/>
            </StatusBarItem>
            <StatusBarItem HorizontalAlignment="Right">
                <Label Content="Ver. 3.6.0"/>
            </StatusBarItem>
        </StatusBar>
    </Grid>
</Controls:MetroWindow>
"@

$Reader = (New-Object System.Xml.XmlNodeReader $xaml) 
$Form = [Windows.Markup.XamlReader]::Load($reader) 

$xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]")  | ForEach-Object { 
  New-Variable  -Name $_.Name -Value $Form.FindName($_.Name) -Force 
}

$Image1.Source = "$PSScriptRoot\Themes\kaf2.png"
$Image2.Source = "$PSScriptRoot\Themes\k.png"
$b1 = $Form.findname("b1")
$b2 = $Form.findname("b2")

$userr = $env:UserName

$ora_server = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String(""))
$o = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String(""))
$p = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String(""))
$ora_user = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String(""))
$ora_pass = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String(""))
$ora_sid = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String(""))

$userrr = $userr.toupper()
$h = "P@ssw0rd"

function mail ($Subject, $Body) {
    $EmailFrom = ""
    $EmailTo = ""

    $Subject = $Subject
    $Body = $Body 

    $SMTPServer = ""

    $msg = new-object Net.Mail.MailMessage
    $msg.From = $EmailFrom
    $msg.to.Add($EmailTo)
    $msg.Subject = $Subject

    $msg.Body = $Body

    $SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 587) 
    $SMTPClient.EnableSsl = $true
    $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($p, $o)

    $SMTPClient.Send($msg)
}

function toend {
    $TextBox2.Focus();
    $TextBox2.CaretIndex = Status.Text.Length;
    $TextBox2.ScrollToEnd();
}

$TextBox2.AppendText("Próba zmiany Hasła zostanie przeprowadzona dla:`r`n")
$TextBox2.AppendText("$userrr`r`n")
$TextBox2.AppendText("`r`n")
$TextBox2.AppendText("Hasło ustawione zostanie na $h`r`n")
$TextBox2.AppendText("Neleży się zalogować wpisująć hasło $h`r`n")
$TextBox2.AppendText("System poprosi o zmianę. Stare hasło to także $h`r`n")
$h | clip

# Hasło
$b1.Add_Click({
    if ($userrr) {
        $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$ora_server)(PORT=1521)) (CONNECT_DATA=(SID=$ora_sid)));User Id=$ora_user;Password=$ora_pass;")
        
        $querycheck = "Select * FROM RI_PRACOWNICY where PRAC_USERNAME LIKE '$userrr'"
        $connection.open()

        $command3=$connection.CreateCommand()
        $command3.CommandText=$querycheck
        $wynik = $command3.ExecuteReader()

        $TextBox2.AppendText("Trwa sprawdzanie. Czekaj...`r`n")
        $TextBox2.AppendText("`r`n")

        if ($wynik.HasRows) {
            $query = 'alter user '+$userrr+' identified by "'+$h+'" account unlock'
            $query2 = "update RI_PRACOWNICY SET PRAC_PASS_CHANGE_DATE = (sysdate)-32 WHERE PRAC_USERNAME = '$userrr'"

            $command=$connection.CreateCommand()
            $command2=$connection.CreateCommand()
            $command.CommandText=$query
            $command2.CommandText=$query2
            $command.ExecuteReader()
            $command2.ExecuteReader()

            $d = Get-Date
            $Subject = "GUI - Hasło"
            $Body = "$userr - użył funkcji zmiany hasła o $d"
            mail $Subject $Body

            $hasko = "Hasło dla $userrr zostało zmienione na $h"
        } else {
            $hasko = "Brak takiego użytkownika: $userrr"
        }

        $TextBox2.AppendText("$hasko")
        $TextBox2.AppendText("`r`n")

        $okOnly = [MahApps.Metro.Controls.Dialogs.MessageDialogStyle]::Affirmative
        $result = [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowModalMessageExternal($Form,"Uwaga","$hasko",$okOnly)
        $dialgResult.Text = $result
        If ($result -eq "Affirmative"){ 
            $dialgResult.Foreground = "Green"
        }
        else{
            $dialgResult.Foreground = "Red"
        }
        
        $connection.close() 
    } else {
        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("Login pusty`r`n")
        $TextBox2.AppendText("`r`n")
    }

    toend
})

$b2.Add_Click({
    if ($userrr) {

        $connection = New-Object Oracle.ManagedDataAccess.Client.OracleConnection("Data Source=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=$ora_server)(PORT=1521)) (CONNECT_DATA=(SID=$ora_sid)));User Id=$ora_user;Password=$ora_pass;")
        
        $querycheck = "Select * FROM RI_PRACOWNICY where PRAC_USERNAME LIKE '$userrr'"
        $connection.open()

        $command3=$connection.CreateCommand()
        $command3.CommandText=$querycheck
        $wynik = $command3.ExecuteReader()

        $TextBox2.AppendText("Trwa sprawdzanie. Czekaj...`r`n")

        if ($wynik.HasRows) {
            $query = 'alter user '+$userrr+' account unlock'

            $command=$connection.CreateCommand()
            $command.CommandText=$query
            $command.ExecuteReader()

            $hasko = "$userrr zostało odblokowane"
    
            $s = Get-Date
            $Subject = "GUI - Odblokowanie"
            $Body = "$userr - użył funkcji odblokuj o $s"
            mail $Subject $Body

        } else {
            $hasko = "Brak takiego użytkownika: $userrr"
        }

        $TextBox2.AppendText("$hasko")
        $TextBox2.AppendText("`r`n")

        $okOnly = [MahApps.Metro.Controls.Dialogs.MessageDialogStyle]::Affirmative
        $result = [MahApps.Metro.Controls.Dialogs.DialogManager]::ShowModalMessageExternal($Form,"Uwaga","$hasko",$okOnly)
        $dialgResult.Text = $result
        If ($result -eq "Affirmative"){ 
            $dialgResult.Foreground = "Green"
        }
        else{
            $dialgResult.Foreground = "Red"
        }
        
        $connection.close()

        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("$hasko")
        $TextBox2.AppendText("`r`n")
    } else {
        $TextBox2.AppendText("`r`n")
        $TextBox2.AppendText("Login pusty`r`n")
        $TextBox2.AppendText("`r`n")
    }
    
    toend
 })

[void]$Form.ShowDialog()