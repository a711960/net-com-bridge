#########################################################################################################################################
#    "C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe" -noexit -File "%1" 
#    "C:\WINDOWS\system32\WindowsPowerShell\v1.0\powershell.exe" -command "Set-ExecutionPolicy RemoteSigned"
#	http://shfb.codeplex.com/
#	http://sandcastle.codeplex.com/
#########################################################################################################################################

cls
mode con cols=120 lines=60

###################### Save allPath ###################################

$Project_name = "SeleniumNetComBridge"
$Project_dir = [System.IO.Directory]::GetCurrentDirectory() + "\"
$CurrentVersion_path = $Project_dir + "Source\Properties\AssemblyInfo.cs"
$csproj_path = $Project_dir + "Source\NetComBridge.csproj"
$iss_path = $Project_dir + "Package.iss"
$shfbproj_path = $Project_dir + "Documentation.shfbproj"

set-alias cmd-7zip "C:\Program Files\7-Zip\7z.exe"
set-alias cmd-iscc "C:\Program Files\Inno Setup 5\ISCC.exe"
set-alias cmd-msbuild "C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe"

function getVersion($version){
    $new_version =""
    while ($new_version -eq ""){
        $input = read-host "   Digit to increment [w.x.y.z] or version [0.0.0.0] or skip [s] ? "
        if ($input -match "s|z|y|x|w") {
            $version_digit = $version -split "\."
            switch ($input){
                "s" { $new_version = $version }
                "z" { $new_version = $version_digit[0] + "." + $version_digit[1] + "." + $version_digit[2] + "." + [string]([int]$version_digit[3]+1) }
                "y" { $new_version = $version_digit[0] + "." + $version_digit[1] + "." + [string]([int]$version_digit[2]+1) + ".0" }
                "x" { $new_version = $version_digit[0] + "." + [string]([int]$version_digit[1]+1) + ".0.0" }
                "w" { $new_version = [string]([int]$version_digit[0]+1) + ".0.0.0" }
            }
        }else{
            if ($input -match "\d+\.\d+\.\d+\.\d+") { $new_version = $input }
        }
    }
    return $new_version
}

function getSha1($filepath){
    [String]$sha1 =""
    $sha1o = New-Object System.Security.Cryptography.SHA1CryptoServiceProvider
    $file = [System.IO.File]::Open($filepath, "open", "read")
    $sha1o.ComputeHash($file)| %{ $sha1 += $_.ToString("x2") }
    $file.Close()
    return $sha1
}

function pause(){
	$ret = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") 
}

function writeErr($texte){
    write-host($texte) -ForegroundColor red; break;
    write-host ""
    read-host "Press enter to quit " | out-null
}


#Check folders and files
(gi Env:PATH).value.split(";")| ForEach {if(!(test-path $_)){write-host("  Error ENV:PATH : Folder " + $_ + """ not found" ) -ForegroundColor Red;}}
Get-Variable -name *_dir | ForEach {if(!(test-path $_.Value)){write-host("  Error : Folder " + $_.Name + "=""" + $_.Value + """ not found") -ForegroundColor Red;}}
Get-Variable -name *_path | ForEach {if(!(test-path $_.Value)){write-host("  Error : File " + $_.Name + "=""" + $_.Value + """ not found") -ForegroundColor Red;}}
Get-Alias -name cmd-* | ForEach {if(!(test-path $_.Definition)){write-host("  Error : Program " + $_.Name + "=""" + $_.Definition + """ not found") -ForegroundColor Red;}}

#Get the version from update.rdf and subversion URL
$CurrentVersion = (([regex]::matches((get-content $CurrentVersion_path), "AssemblyVersion\(""([\.\d]+)\""\)"))[0]).Groups[1].Value;

#Get last compilation date
$LastCompil_date = if ((test-path($CurrentVersion_path)) -eq 1){(get-item( $CurrentVersion_path )).LastWriteTime}

#Write information on the screen

write-host " __________________________________________________________________________________________________"
write-host ""
write-host " Project name    : "  $Project_name
write-host " Directory       : "  $Project_dir
write-host " Current Version : "  $CurrentVersion
write-host " Last compile    : "  $LastCompil_date
write-host " __________________________________________________________________________________________________"
write-host ""

[String]$NewVersion = getVersion($CurrentVersion)
write-host "   New version : " $NewVersion
write-host ""
write-host "   ** Update the version in AssemblyInfo.cs ..."
		(get-content $CurrentVersion_path ) | %{$_ -replace "Version\(""[^\)]+""\)", "Version(""$NewVersion"")" } | Set-Content -path $CurrentVersion_path

write-host ""
write-host "   ** Msbuild compile sources ..."
	cmd-msbuild /v:quiet /p:Configuration=Release /p:TargetFrameworkVersion=v3.5 /p:SignAssembly=true $csproj_path
    if($LASTEXITCODE -eq 1) { writeErr("  Source compilation failed ! "); pause ; break; }

write-host ""
write-host "   ** Api documentation creation ..."
	$ask_compil=(read-host "   Create the .chm help file ? [y/n] ")
	if ($ask_compil -eq "y"){
		cmd-msbuild /v:quiet /p:CleanIntermediates=True /p:Configuration=Release $shfbproj_path
		if($LASTEXITCODE -eq 1) { writeErr("  Api documentation creation failed ! ") ; pause ; break; }
	}
write-host ""
write-host "   ** InnoSetup create the paclage ..."
	cmd-iscc /q $iss_path
    if($LASTEXITCODE -eq 1) { writeErr("  Package creation failed ! "); pause ; break; }
    
write-host ""
write-host "   ** Package installaton ..."
    start-process -wait ($Project_dir + "NetComBridgeSetup-" + [regex]::matches($NewVersion,"^\d+\.\d+\.\d+")[0].Value  + ".exe")

write-host ""
write-host ""
read-host "Press enter to quit " | out-null
exit