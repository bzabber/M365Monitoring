foreach ($file in (Get-ChildItem -Path "$PSScriptRoot\internal\functions"))
{
	. $file.FullName
}

foreach ($file in (Get-ChildItem -Path "$PSScriptRoot\functions"))
{
	. $file.FullName
}