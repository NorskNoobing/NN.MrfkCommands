Register-ArgumentCompleter -CommandName New-ShippingLabel -ParameterName location -ScriptBlock {((Get-SnipeLocations).name | Sort-Object).ForEach({"`'$_`'"})}