<#
.SYNOPSIS
   <A brief description of the script>
.DESCRIPTION
   <A detailed description of the script>
.PARAMETER <paramName>
   <Description of script parameter>
.EXAMPLE
   <An example of using the script>
#>

Add-Type -AssemblyName System.Windows.Forms
$Form = New-Object System.Windows.Forms.Form
$Form.Text = "Sample Form"
$Label = New-Object System.Windows.Forms.Label
$Label.Text = "This form is very simple"
$Label.AutoSize = $true
$Form.Controls.Add($Label)
$Form.ShowDialog()