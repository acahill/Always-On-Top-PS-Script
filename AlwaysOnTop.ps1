#Requires -Version 2.0 
param(
    [switch]$multi = $false
)

$signature = @"
	[DllImport("user32.dll")]
	public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

	public static IntPtr FindWindow(string windowName) {
		return FindWindow(null,windowName);
	}

	[DllImport("user32.dll")]
	public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X,int Y, int cx, int cy, uint uFlags);

	[DllImport("user32.dll")]
	public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

	static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
	static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);

	const UInt32 SWP_NOSIZE = 0x0001;
	const UInt32 SWP_NOMOVE = 0x0002;

    const int GWL_EXSTYLE = -20;
    const int WS_EX_TOPMOST = 0x0008;

	const UInt32 TOPMOST_FLAGS = SWP_NOMOVE | SWP_NOSIZE;

	public static void MakeTopMost (IntPtr fHandle) {
		SetWindowPos(fHandle, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS);
	}

	public static void MakeNormal (IntPtr fHandle) {
		SetWindowPos(fHandle, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS);
	}

    public static bool IsWindowTopMost(IntPtr hWnd) {
        int exStyle = GetWindowLong(hWnd, GWL_EXSTYLE);
        return (exStyle & WS_EX_TOPMOST) == WS_EX_TOPMOST;
    }
"@

$app = Add-Type -MemberDefinition $signature -Name Win32Window -Namespace ScriptFanatic.WinAPI -ReferencedAssemblies System.Windows.Forms -Using System.Windows.Forms -PassThru

function Set-TopMost {
	param(		
		[Parameter(Position=0,ValueFromPipelineByPropertyName=$true)][Alias('MainWindowHandle')]$hWnd=0,
		[Parameter()][switch]$Disable
	)
	
	if($hWnd -ne 0) {
		if($Disable) {
			Write-Host "Set process handle :$hWnd to NORMAL state"
			$null = $app::MakeNormal($hWnd)
			return
		}
		
		Write-Host "Set process handle :$hWnd to TOPMOST state"
		$null = $app::MakeTopMost($hWnd)
	}
	else {
		Write-Verbose "$hWnd is 0"
	}
}

function Is-TopMost {
    param(		
		[Parameter(Position=0,ValueFromPipelineByPropertyName=$true)][Alias('MainWindowHandle')]$hWnd=0
	)
    if($hWnd -ne 0) {
        return $app::IsWindowTopMost($hWnd)
    }
    return 0
}

function getOpenApplications() {
    return Get-Process | Where-Object {$_.MainWindowTitle -ne ""} | Select-Object Id,MainWindowHandle,MainWindowTitle,Name,ProcessName,@{Name="TopMost"; Expression={Is-TopMost($_.MainWindowHandle)}}
}

function Get-WindowByTitle($WindowTitle="*") {
	Write-Verbose "WindowTitle is: $WindowTitle"
	
	if($WindowTitle -eq "*") {
		Write-Verbose "WindowTitle is *, print all windows title"
		Get-Process | Where-Object {$_.MainWindowTitle} | Select-Object Id,Name,MainWindowHandle,MainWindowTitle
	}
	else {
        $WindowTitle = $WindowTitle.replace(']','`]').replace('[','`[').replace('*','`*')
		Write-Verbose "WindowTitle is $WindowTitle"
		Get-Process | Where-Object {$_.MainWindowTitle -like "*$WindowTitle*"} | Select-Object Id,Name,MainWindowHandle,MainWindowTitle
	}
}


function forceApplicationOnTop($chosenApplication) {
    Write-Host "Chosen Window: "$chosenApplication

    if (!$multi) {
        Write-Host "Disabling other topmost windows..."
        $openApplications = getOpenApplications

        $openApplications | ForEach-Object {
            Get-WindowByTitle $_.MainWindowTitle | Set-TopMost -Disable
        }
    }
    Get-WindowByTitle $chosenApplication | Set-TopMost 

}

function createDropdownBox($openApplications) {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

    $width = 650

    $objForm = New-Object System.Windows.Forms.Form 
    $objForm.Text = "Always On Top"
    $objForm.Size = New-Object System.Drawing.Size($width,400) 
    $objForm.StartPosition = "CenterScreen"

    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") {$x=$objListBox.SelectedItem;$objForm.Close()}})
    $objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") {$objForm.Close()}})

    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = New-Object System.Drawing.Size(300,330)
    $OKButton.Size = New-Object System.Drawing.Size(75,23)
    $OKButton.Text = "OK"
    $OKButton.Add_Click({$x=$objListBox.SelectedItem;$objForm.Close()})
    $objForm.Controls.Add($OKButton)

    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = New-Object System.Drawing.Size(400,330)
    $CancelButton.Size = New-Object System.Drawing.Size(75,23)
    $CancelButton.Text = "Cancel"
    $CancelButton.Add_Click({$objForm.Close()})
    $objForm.Controls.Add($CancelButton)

    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(10,20) 
    $objLabel.Size = New-Object System.Drawing.Size(($width-35), 20)
    $objLabel.Text = "Select a window to keep on top:"
    $objForm.Controls.Add($objLabel) 

    $objListBox = New-Object System.Windows.Forms.ListView
    $objListBox.Location = New-Object System.Drawing.Size(10,40)
    $objListBox.Size = New-Object System.Drawing.Size(($width-35),20)
    $objListBox.Height = 280
    $objListBox.View = [System.Windows.Forms.View]::Details
    $objListBox.CheckBoxes = 1

    [void] $objListBox.Columns.Add("     Title", -2)
    [void] $objListBox.Columns.Add("Process Name", -2)

    $openApplications | ForEach-Object {
        $item = New-Object System.Windows.Forms.ListViewItem($_.MainWindowTitle)
        [void] $item.SubItems.Add($_.ProcessName)
        $item.Checked = $_.TopMost
        if ($_.TopMost) {
            $item.Font = New-Object System.Drawing.Font($objListBox.Font, [System.Drawing.FontStyle]::Bold)
        }
        [void] $objListBox.Items.Add($item)
    }

    $objForm.Controls.Add($objListBox) 

    $objForm.Topmost = $True

    $objForm.Add_Shown({$objForm.Activate()})
    [void] $objForm.ShowDialog()

    if ($objListBox.SelectedItems) {
        $x=$objListBox.SelectedItems[0].SubItems[0].Text
    }
    return $x
}

############ Script starts here ###################

getOpenApplications | Format-Table
Write-Host "Multiple windows topmost = $multi"
$openApplications = getOpenApplications
$chosenApplication = createDropdownBox($openApplications)
Write-Host "Selected Window = $chosenApplication"

if ($chosenApplication) {
    forceApplicationOnTop($chosenApplication)
    getOpenApplications | Format-Table
}
