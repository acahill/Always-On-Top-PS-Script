#Requires -Version 2.0

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
	const UInt32 SWP_ASYNCWINDOWPOS = 0x4000;
	const UInt32 SWP_NOACTIVATE = 0x0010;

    const int GWL_EXSTYLE = -20;
    const int WS_EX_TOPMOST = 0x0008;

	const UInt32 TOPMOST_FLAGS = SWP_NOMOVE | SWP_NOSIZE | SWP_ASYNCWINDOWPOS | SWP_NOACTIVATE;

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
$objListBox = New-Object System.Windows.Forms.ListView

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


function forceApplicationsOnTop() {
    $chosenApplications = $objListBox.CheckedItems
    Write-Host "Chosen Windows: $chosenApplications"
    Write-Host "Disabling other topmost windows..."
    $openApplications = getOpenApplications

    $openApplications | ForEach-Object {
        Get-WindowByTitle $_.MainWindowTitle | Set-TopMost -Disable
    }

    $chosenApplications | ForEach-Object {
        Get-WindowByTitle $_.Text | Set-TopMost
    }
    populateListBox
}


function populateListBox() {
    Write-Host "Refreshing window list"
    $openApplications = getOpenApplications
    $objListBox.Items.Clear()
    $openApplications | ForEach-Object {
        $item = New-Object System.Windows.Forms.ListViewItem($_.MainWindowTitle)
        [void] $item.SubItems.Add($_.ProcessName)
        $item.Checked = $_.TopMost
        if ($_.TopMost) {
            $item.Font = New-Object System.Drawing.Font($objListBox.Font, [System.Drawing.FontStyle]::Bold)
        }
        [void] $objListBox.Items.Add($item)
    }
}


function createForm($openApplications) {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

    $width = 650

    $objForm = New-Object System.Windows.Forms.Form
    $objForm.Text = "Always On Top"
    $objForm.Size = New-Object System.Drawing.Size($width,400) 
    $objForm.StartPosition = "CenterScreen"
    $objForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedSingle

    $objForm.KeyPreview = $True
    $objForm.Add_KeyDown({ if ($_.KeyCode -eq "Enter") { forceApplicationsOnTop } })
    $objForm.Add_KeyDown({ if ($_.KeyCode -eq "Escape") { $objForm.Close() } })

    $ApplyButton = New-Object System.Windows.Forms.Button
    $ApplyButton.Location = New-Object System.Drawing.Size(175,330)
    $ApplyButton.Size = New-Object System.Drawing.Size(75,23)
    $ApplyButton.Text = "Apply"
    $ApplyButton.Add_Click({ forceApplicationsOnTop })
    $objForm.Controls.Add($ApplyButton)

    $RefreshButton = New-Object System.Windows.Forms.Button
    $RefreshButton.Location = New-Object System.Drawing.Size(275,330)
    $RefreshButton.Size = New-Object System.Drawing.Size(75,23)
    $RefreshButton.Text = "Refresh"
    $RefreshButton.Add_Click({ populateListBox })
    $objForm.Controls.Add($RefreshButton)

    $CloseButton = New-Object System.Windows.Forms.Button
    $CloseButton.Location = New-Object System.Drawing.Size(375,330)
    $CloseButton.Size = New-Object System.Drawing.Size(75,23)
    $CloseButton.Text = "Close"
    $CloseButton.Add_Click({ $objForm.Close() })
    $objForm.Controls.Add($CloseButton)

    $objLabel = New-Object System.Windows.Forms.Label
    $objLabel.Location = New-Object System.Drawing.Size(10,20) 
    $objLabel.Size = New-Object System.Drawing.Size(($width - 35), 20)
    $objLabel.Text = "Select a window to keep on top:"
    $objForm.Controls.Add($objLabel) 

    $objListBox.Location = New-Object System.Drawing.Size(10,40)
    $objListBox.Size = New-Object System.Drawing.Size(($width - 27),20)
    $objListBox.Height = 280
    $objListBox.View = [System.Windows.Forms.View]::Details
    $objListBox.CheckBoxes = 1
    $objListBox.MultiSelect = 0
    $objListBox.FullRowSelect = 1
    $objListBox.HideSelection = 1

    $objListBox.Add_ItemSelectionChanged({
        param($listview, $e)
        if ($e.IsSelected) {
            $e.Item.Selected = 0
            $e.Item.Checked = !$e.Item.Checked
        }
    })

    [void] $objListBox.Columns.Add("     Title",   [int]($width - ($width / 4)))
    [void] $objListBox.Columns.Add("Process Name", [int]($width / 4 - 40))

    populateListBox

    $objForm.Controls.Add($objListBox) 

    $objForm.Add_Shown({ $objForm.Activate() })
    [void] $objForm.ShowDialog()
}


############ Script starts here ###################

getOpenApplications | Format-Table
$openApplications = getOpenApplications
$chosenApplication = createForm($openApplications)
