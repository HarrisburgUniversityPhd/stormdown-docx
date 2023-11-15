write-host -NoNewline "Loading StormDown tools ... "

function write-template {
	[CmdletBinding(DefaultParameterSetName = 'name')]
    Param
    (
         [Parameter(Mandatory = $true, Position = 0, ParameterSetName = 'name')]
         [string] $name
    )	
	if (Test-Path -Path $name -PathType Container) {
		write-host "Writing: ./$name >>> $name.dotx"
		$7z = 'C:/Program Files/7-Zip/7z.exe'
		. $7z a -tzip -mx1 -r "$($name).dotx" "./$name/*.*"
	} else {
		write-host "Could not find folder: $name"
	}
}

function write-folder {
	[CmdletBinding(DefaultParameterSetName = 'name')]
    Param
    (
         [Parameter(Mandatory = $true, Position = 0, ParameterSetName = 'name')]
         [string] $name
    )
	if (Test-Path -Path "$name.dotx" -PathType Leaf) {
		write-host "Writing: template $name.dotx >>> folder ./$name"
		$7z = 'C:/Program Files/7-Zip/7z.exe'
		rm $name -r -force
		. $7z x -y $("-o$name") "$($name).dotx"
		dir -r "$($name)/*.xml" | % { ([xml](gc -LiteralPath $_ -encoding utf8)).Save($_) }
		dir -r "$($name)/*.rels" | % { ([xml](gc -LiteralPath $_ -encoding utf8)).Save($_) }
		dir -r "$($name)/*.xml" | % {
			$xml = [xml](gc -LiteralPath $_ -encoding utf8)
			$nodes = $xml.SelectNodes('//*')
			$attributes = @('w:rsidR', 'w:rsidRDefault', 'w:rsidP', 'w:rsidRPr', 'w:rsidSect', 'w:rsidTr')
			foreach($attribute in $attributes) { foreach($node in $nodes) { $node.RemoveAttribute($attribute) }}
			$xml.Save($_)
		}
	} else {
		write-host "Could not find file: $name.dotm"
	}
}

write-host "Done"
