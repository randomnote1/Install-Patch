#Requires –Version 5

function Install-Patch
{
    #Requires –Version 5
	
	Param
    (
        ## A PsSession to the computer
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.Runspaces.PSSession]$PsSession,

        ## The path to the patch
        [Parameter(Mandatory=$true)]
        [ValidateScript({Test-Path -Path $_})]
        [System.IO.FileInfo]$Patch
    )
    
    function Invoke-PostInstallActions
    {
        Param
        (
            $ReturnValue,
            $ComputerName,
            $Patch
        )

        switch ( $ReturnValue )
        {
            ## Patch installed fine
            0
            {
                ## Give the user some feedback
                Write-Host ( "`t " + $Patch + " Complete" ) -ForegroundColor Green
            }

            ## Patch requires a reboot
            3010
            {
                ## Give the user some feedback
                Write-Host ( "`t " + $Patch + " Complete.  Rebooting " + $ComputerName + "..." ) -ForegroundColor Green

                ## Reboot the computer
                Restart-Computer -ComputerName $ComputerName -Wait -For Wmi -Force
            }

            ## Patch is already installed and requires a reboot
            2359302
            {
                ## Give the user some feedback
                Write-Host ( "`t " + $Patch + " Complete.  Rebooting " + $ComputerName + "..." ) -ForegroundColor Green

                ## Reboot the computer
                Restart-Computer -ComputerName $ComputerName -Wait -For Wmi -Force
            }

            ## The patch is not applicable to the image
            -2146498530
            {
                ## Give the user some feedback
                Write-Host ( "`t " + $Patch + " not applicable." ) -ForegroundColor Green
            }

            ## Any other error
            default { throw ( "Error installing the patch " + $Patch + ".  Return Code: " + $ReturnValue ) }
        }
    }

	Write-Host ( "`t Installing " + $Patch + " on " + $PsSession.ComputerName + "..." ) -ForegroundColor Green

    ## Get   the destination directory path
    $destPath = Invoke-Command -Session $PsSession -ScriptBlock {$env:TEMP}

	# Copy the update file
	Copy-Item -Path $Patch.FullName -Destination $destPath -ToSession $PsSession

    # Create a string to use as the log file name
    $logFileName = $Patch.Name + "_" + ( Get-Date -Format "yyyyMMdd_hhmmss" )

    # Build the command and arguments
    switch ( $Patch.Extension )
    {
        ".exe" 
        {
            $command = $destPath + "\" + $Patch.Name
            $arguments = "/q /norestart /log '" + $destPath + "\" + "Log_" + $logFileName + ".htm'"

            ## Create a script block to execute via WMI
			$wmiCommand = {
				Param
				(
					$destPath,
					$logFileName,
					$patchName,
					$command,
					$commandArguments
				)

				## Set our error action preference
                $ErrorActionPreference = "Stop"

                ## Trap our errors and put them in a log file
				trap { $_ | fl -Force | Out-String | Out-File -FilePath ( Join-Path -Path $destPath -ChildPath "Error_$logFileName.txt" ) }

                ## Get the patch path
				$patch = Join-Path -Path $destPath -ChildPath $patchName

                ## Create the process
                $processInfo = New-Object System.Diagnostics.ProcessStartInfo $command
                $processInfo.Arguments = $commandArguments
                $processInfo.UseShellExecute = $false

                ## Start the process
                $setup = [System.Diagnostics.Process]::Start($processInfo)

                ## wait for the installation to complete
                $setup.WaitForExit()

                ## Return the setup object
				$setup | Export-Clixml -Path ( Join-Path -Path $destPath -ChildPath "ProcessDetails_$logFileName.xml" )
			}

            ## Encode the command
			$wmiCommandString = $wmiCommand.ToString()
            $wmiCommandStringBytes = [Text.Encoding]::Unicode.GetBytes($wmiCommandString)
            $wmiCommandStringEncoded = [Convert]::ToBase64String($wmiCommandStringBytes)

            ## Get the PowerShell path
            $psPath = Invoke-Command -Session $PsSession -ScriptBlock { ( Get-Command powershell.exe ).Definition } -ArgumentList $destPath,$logFileName,$Patch.Name,$command,$arguments

            ## Generate the final command to be passed into the process creation statement
            $command = $psPath + ' -NoProfile -EncodedCommand "' + $wmiCommandStringEncoded + '"'

            ## Install the patch
            $process = Invoke-WmiMethod -ComputerName $PsSession.ComputerName -Class "Win32_Process" -Name Create($command)

            ## While the process is still running
            while ( Get-Process -ComputerName $PsSession.ComputerName -Id $process.ProcessId -ErrorAction SilentlyContinue )
            {
                ## Wait for a bit
                Start-Sleep -Seconds 30

                ## Give the user some feedback
                Write-Host "." -NoNewline -ForegroundColor Green
            }

            ## Import the process details object
			$setup = Invoke-Command -Session $PsSession -ScriptBlock { param($destPath,$logFileName) Import-Clixml -Path ( Join-Path -Path $destPath -ChildPath "ProcessDetails_$logFileName.xml" ) -ErrorAction SilentlyContinue } -ArgumentList $destPath,$logFileName

            ## Clean up after ourselves
			Invoke-Command -Session $PsSession -ScriptBlock { param($destPath,$logFileName) Remove-Item -Path ( Join-Path -Path $destPath -ChildPath "ProcessDetails_$logFileName.xml" ) -Force -ErrorAction SilentlyContinue } -ArgumentList $destPath,$logFileName

            ## Check the return value and do stuff
            Invoke-PostInstallActions -ReturnValue $setup.ExitCode -ComputerName $PsSession.ComputerName -Patch $Patch.Name
        }
        ".msu"
        {
            ## Extract the MSU File
            $retVal = Invoke-Command -Session $PsSession `
                                     -ArgumentList $destPath,$Patch.Name,$logFileName `
                                     -ScriptBlock {
                                        Param($destPath,$fileName,$logFileName)

                                        ## Define the arguments
                                        $arguments = @()
										$arguments += Join-Path -Path $destPath -ChildPath $fileName
										 # Specify the path to extract the MSU
										$arguments += '/extract:' + ( Join-Path -Path $destPath -ChildPath ( Get-Item -Path ( Join-Path -Path $destPath -ChildPath $fileName ) ).BaseName )
                                        $arguments += "/log:" + $destPath + "\Extract_Log_" + $logFileName + ".evtx"

                                        ## Start the process
                                        $proc = Start-Process -FilePath ( $env:SystemRoot + "\System32\wusa.exe" ) `
                                                              -ArgumentList $arguments `
                                                              -PassThru `
                                                              -Wait
                                        
                                        ## Return the exit code
                                        $proc.ExitCode
                                     }

            ## Install the package
            if ( $retVal -eq 0 )
            {
                ## Get the patch name
				$patchName = Invoke-Command -Session $PsSession -ScriptBlock { param($destPath,$fileName) ( Get-Item -Path ( Join-Path -Path $destPath -ChildPath $fileName ) ).BaseName } -ArgumentList $destPath,$Patch.Name
                
                ## Determine if there is a package install order file
				if ( Invoke-Command -Session $PsSession -ScriptBlock { param($destPath,$patchName) Test-Path -Path ( Join-Path -Path $destPath -ChildPath "$patchName\PkgInstallOrder.txt" ) } -ArgumentList $destPath,$patchName )
                {
                    ## This is an ini file, so open it up
					$installOrder = ConvertFrom-Ini -FileContent ( Invoke-Command -Session $PsSession -ScriptBlock { param($destPath,$patchName) Get-Content -Path ( Join-Path -Path $destPath -ChildPath "$patchName\PkgInstallOrder.txt" ) } -ArgumentList $destPath,$patchName )

                    ## get the cab files. Sort by the name column because this specifies the install order and hash tables don't always return in the same order
                    [array]$cabs += $installOrder['MSUInstallOrder'].GetEnumerator() | sort Name | % { $_.Value }
                }
                else
                {
                    # Find all the CAB files that have a KB number in them
					[array]$cabs += Invoke-Command -Session $PsSession -ScriptBlock { param($destPath,$fileName) Get-ChildItem -Path ( Join-Path -Path $destPath -ChildPath $fileName ) -Filter '*.cab' | ? { $_.Name -match 'KB[0-9]{7}' } | select -ExpandProperty Name } -ArgumentList $destPath,$Patch.BaseName
                }
                
                foreach ( $cab in $cabs )
                {
                    ## Set the log file name
                    $lfn = $logFileName -replace $Patch.Name,$cab
                    
                    ## Install the patch
                    $retVal = Invoke-Command -Session $PsSession `
                                             -ArgumentList $destPath,$Patch.Name,$lfn,$cab `
                                             -ScriptBlock {
                                                Param($destPath,$fileName,$logFileName,$cab)

                                                ## Get the patch name
                                                $patchName = ( Get-Item -Path ( $destPath + "\" + $fileName ) ).BaseName 

                                                ## Define the arguments
                                                $arguments = @()
                                                $arguments += "/online"
                                                $arguments += "/add-package"
                                                $arguments += "/PackagePath:" + $destPath + "\" + $patchName + "\" + $cab
                                                $arguments += "/LogPath:" + $destPath + "\Install_Log_" + $logFileName + ".log"
                                                $arguments += "/LogLevel:3"
                                                $arguments += "/quiet"
                                                $arguments += "/NoRestart"

                                                ## Start the process
                                                $proc = Start-Process -FilePath ( $env:SystemRoot + "\System32\dism.exe" ) `
                                                                      -ArgumentList $arguments `
                                                                      -PassThru `
                                                                      -Wait

                                                ## Return the exit code
                                                $proc.ExitCode
                                             }
                    
                    ## Check the return value and do stuff
                    Invoke-PostInstallActions -ReturnValue $retVal -ComputerName $PsSession.ComputerName -Patch $cab
                }
            }
            else { throw ( "Failed to extract the '" + $Patch.Name + "' patch.  Error code: " + $retVal ) }
        }
    }
}