<#
    Script Name  : Softwire SIM (PoC) 
    
    Description  : A customised T_O_O_L (Toolkit for Operational Optimisation and Learning) to:
                     - Interact with simulated doors manually.
                     - Simulate a busy site automatically for stress testing / training / demos. 
                      
                   This is just a PoC to prove logic, final build will be wrapped in a nice interactive GUI (at least... that's the plan).

    @Author      : James Savage.

    Last Updated : 27-03-2026 (update in the below "LandingPage" function if required).

    Version      : BETA 1.0 (update in the below "LandingPage" function if required).

    Note 1       : If you can't run PS scripts (it keeps closing) open a PS console as admin and run (accept Yes to All): 
                        
                   Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
#>

# Clear terminal
cls


###############################################################################################################################################################################################################
#                                                [1/4] IMPORT THE SOFTWIRE MODULE
#                                         --------------------------------------------
###############################################################################################################################################################################################################

# Update variable if yours is in a different location
$softwireModuleLocation = "C:\ProgramData\SoftwirePSM\Softwire.psd1"
 
try {
    if (Test-Path $softwireModuleLocation) {
        Import-Module $softwireModuleLocation -Force -ErrorAction Stop
        Write-Host "Softwire PS module imported..." -ForegroundColor Green
        Write-Host ""
    }
    else {
        throw "Module path does not exist: $softwireModuleLocation" 
    }
} 
catch {
    Write-Host "Cannot import the Softwire PowerShell module!" -ForegroundColor Red
    Write-Host ""
    Write-Host "Reason: $($_.Exception.Message)" -ForegroundColor Darkyellow
    Write-Host ""
    Write-Host "Fix:" -ForegroundColor Green
    Write-Host "  - Ensure SoftwireAPI.exe is located in: C:\ProgramData\SoftwirePowershell" -ForegroundColor Green
    Write-Host "  - Ensure Softwire.psd1 and Softwire.psm1 are located in: C:\ProgramData\SoftwirePSM" -ForegroundColor Green
    Write-Host "  - if you need help, email: jsavage@genetec.com" -ForegroundColor Green
    return
}


###############################################################################################################################################################################################################
#                                              [2/4] STATIC & DECLARED VARIABLES
#                                         --------------------------------------------
###############################################################################################################################################################################################################

# Variables used for Softwire commands that use -Session
$sclEndpoint = "127.0.0.1" # FIXED
$sclUsername = "admin"     # FIXED
$sclPassword = $null       # Variable declared - updates in 'Connect-Softwire' function
$sclSession  = $null       # Variable declared - updates in 'Connect-Softwire' function


###############################################################################################################################################################################################################
#                                                       [3/4] FUNCTIONS
#                                         --------------------------------------------
###############################################################################################################################################################################################################

#------------------------------------------------------------------------------------------------------
#--> [a] Main Menu Functions
#------------------------------------------------------------------------------------------------------
function Menu-ManualDoorUsage {
    <#
    Purpose:
        - Let user select a door
        - Display that door's hardware
        - Let user simualte a read, operate an input, or see status of the door
    #>

    while ($true) {

        #------------------------------------------------------------------------------------------------------------------
        # 1. List doors and ask user which one they want to work with
        #------------------------------------------------------------------------------------------------------------------
        $doors = Get-Doors

        Write-TitleToConsole -Title "Manual door usage menu"

        Write-Host "There are $($doors.Count) doors configured in your system:`n"

        for ($i = 0; $i -lt $doors.Count; $i++) {
            Write-Host "[$i] $($doors[$i].Name)"
        }

        Write-Host ""
        Write-Host "[C] Clear the console" -ForegroundColor DarkYellow
        Write-Host "[R] Refresh the console" -ForegroundColor DarkYellow
        Write-Host "[Q] Return to main menu" -ForegroundColor DarkYellow
        Write-Host "`nEnter the number of the door you want to use: " -ForegroundColor Cyan -NoNewline
        $selection = (Read-Host).Trim().ToUpperInvariant()

        if ($selection -eq 'C') {
            ClearConsole-Pretty -Message "The console will clear in"; continue
        }

        if ($selection -eq 'R') {
            cls; continue
        }

        if ($selection -eq 'Q') {
            ClearConsole-Pretty -Message "The console will clear in"; return
        }

        if ($selection -notmatch '^\d+$' -or [int]$selection -lt 0 -or [int]$selection -ge $doors.Count) {
            Write-Host "Invalid selection." -ForegroundColor Red; Start-Sleep 1; continue
        }

        ClearConsole-Pretty -Message "The console will clear in";

        $selectedDoor = $doors[[int]$selection]
        $selectedDoorID = $selectedDoor.Id     

        #------------------------------------------------------------------------------------------------------------------
        # 2. What does the user want to do with this selected door?
        #------------------------------------------------------------------------------------------------------------------
        :DoorActionMenu while ($true) {
            
            # Refresh the door incase we end up here with an "old snapshot" of the door
            $selectedDoor = Refresh-CurrentDoor -Door $selectedDoor

            #--------------------------------------------------------------------------------------------------------------
            # 3. List readers, inputs, and outputs that belong to the door
            #--------------------------------------------------------------------------------------------------------------
            Write-TitleToConsole -Title "You selected: $($selectedDoor.Name)"
            
            $hardware = Get-DoorHardware -Door $selectedDoor

            $readers = @($hardware | Where-Object RoleType -eq 'ReaderAuth' | Sort-Object Side)
            $inputs  = @($hardware | Where-Object RoleType -in @('OpenSensor','REX','ManualStation'))
            $outputs = @($hardware | Where-Object RoleType -eq 'Strike')

            Show-DoorHardwareSection -Title "Readers found:" -Items $readers
            Show-DoorHardwareSection -Title "Inputs found:"  -Items $inputs
            Show-DoorHardwareSection -Title "Outputs found:" -Items $outputs

            Write-Host "`nThere are 3 options available for interacting with this door:`n"
            Write-Host "[1] Simulate a read"
            Write-Host "[2] Operate an input"
            Write-Host "[3] Show the status of the door's readers, inputs, and outputs"
            Write-Host ""
            Write-Host "[C] Clear the console" -ForegroundColor DarkYellow
            Write-Host "[R] Refresh the console" -ForegroundColor DarkYellow
            Write-Host "[Q] Return to door selection" -ForegroundColor DarkYellow

            Write-Host "`nChoose what you want to do with this door: " -ForegroundColor Cyan -NoNewline
            $choice = (Read-Host).Trim().ToUpperInvariant()

            if ($choice -eq 'C') {
                ClearConsole-Pretty -Message "The console will clear in"
                continue
            }

            if ($choice -eq 'R') {
                $selectedDoor = Refresh-CurrentDoor -Door $selectedDoor
                cls
                continue
            }

            if ($choice -eq 'Q') {
                ClearConsole-Pretty -Message "The console will clear in"
                break
            }

            if ($choice -notmatch '^[1-3]$') {
                Write-Host "Invalid selection. Please try again." -ForegroundColor Red
                Start-Sleep 1
                continue
            }

            switch ($choice) {
                "1" { ClearConsole-Pretty -Message "The console will clear in"; $selectedDoor = Refresh-CurrentDoor -Door $selectedDoor; Manual-InvokeDoorRead -Door $selectedDoor -Readers $readers -Inputs $inputs -Outputs $outputs }
                "2" { ClearConsole-Pretty -Message "The console will clear in"; $selectedDoor = Refresh-CurrentDoor -Door $selectedDoor; Manual-InvokeDoorInput -Door $selectedDoor -Inputs $inputs }
                "3" { ClearConsole-Pretty -Message "The console will clear in";
                    while ($true) {

                        Write-TitleToConsole -Title "Show status of devices for door: $($selectedDoor.Name)"

                        # Refresh the HW again
                        $selectedDoor = Refresh-CurrentDoor -Door $selectedDoor
                        $hardware = Get-DoorHardware -Door $selectedDoor
                        $readers = @($hardware | Where-Object RoleType -eq 'ReaderAuth' | Sort-Object Side)
                        $inputs  = @($hardware | Where-Object RoleType -in @('OpenSensor','REX','ManualStation'))
                        $outputs = @($hardware | Where-Object RoleType -eq 'Strike')

                        Show-DoorStatus -Door $selectedDoor -Readers $readers -Inputs $inputs -Outputs $outputs

                        $toRefreshOrNotToRefresh = (Read-Host "`nNote: This information will not refresh live, press R to refresh or ENTER to continue").Trim().ToUpperInvariant()

                        if ([string]::IsNullOrWhiteSpace($toRefreshOrNotToRefresh)) {
                            ClearConsole-Pretty -Message "The console will clear in"; break
                        }

                        if ($toRefreshOrNotToRefresh -eq 'R') {
                            cls; continue
                        }

                        Write-Host "Invalid selection. Press R to refresh or ENTER to continue." -ForegroundColor Red
                        Start-Sleep 1
                    }
                }
            }
        }
    }
} 

function Menu-SimulateDoorUsage {
    <#
    Purpose:
        - Let user run a simulation of a site
        - Can be used to stress test or for demos/training
    #>
    
    #--------------------------------------------------------------------------------------------------------
    #--> 1. Write title to console and show info
    #--------------------------------------------------------------------------------------------------------
    Write-TitleToConsole -Title "Simulated door usage menu"

    Simulation-LandingPage

    #--------------------------------------------------------------------------------------------------------
    #--> 2. Gather variables
    #--------------------------------------------------------------------------------------------------------
    $simulationVariables = Simulation-GatherVariables

    # # The returned value is a PSCustomObject containing the simulation settings
    $minWaitTime = $simulationVariables.LowTimer     #--> The minimum time to wait before an event is triggered | INTEGER
    $maxWaitTime = $simulationVariables.HighTimer    #--> The maximum time to wait before an event is triggered | INTEGER
    $numOfEvents = $simulationVariables.NumOfEvents  #--> The number of events that should be generated         | INTEGER
    $eventNormal = $simulationVariables.NormalEvents #--> The percentage of "normal" events                     | INTEGER
    $eventForced = $simulationVariables.ForcedEvents #--> The percentage of "door forced" events                | INTEGER
    $eventHeld   = $simulationVariables.HeldEvents   #--> The percentage of "door held" events                  | INTEGER
    $globalPin   = $simulationVariables.PinToUse     #--> The global PIN to use (if readers are card+PIN)       | STRING
    
    #--------------------------------------------------------------------------------------------------------
    #--> 3. Catch point - can quit here
    #--------------------------------------------------------------------------------------------------------
    :catchPoint while ($true) {
        
        # Final Note!
        Write-Host "Feel free to minimise this window and configure/add/edit/delete doors/door hardware/cardholders/credentials/access rules etc. while this simulation is running" -ForegroundColor DarkGray
        Write-Host "Remember for Card + PIN readers cardholders will all need a PIN of: $($globalPin) (ignore any leading 0, it's just used for the simulator)" -ForegroundColor DarkGray
        Write-Host ""

        # Once the user is happy with all variables, press enter to run the simulation
        $response = (Read-Host "Are you ready... press ENTER to start the simulation or Q to quit").Trim().ToUpperInvariant()

        if ([string]::IsNullOrWhiteSpace($response)) {
            ClearConsole-Pretty -Message "Simulation starting"
            break catchPoint
        }

        if ($response -eq "Q") {
            ClearConsole-Pretty -Message "Returning to the previous menu"
            return
        }

        Write-Host "Invalid choice. Press ENTER to continue or type Q to quit." -ForegroundColor Red
        Write-Host ""
        Start-Sleep 1
        continue catchPoint   

    }

    #--------------------------------------------------------------------------------------------------------
    #--> 4. Simulation logic
    #--------------------------------------------------------------------------------------------------------
    # Declare variable for tracking when to stop the while loop
    $generatedEvents      = 0

    # Declare variables 
    $selectedNormalCount  = 0
    $selectedForcedCount  = 0
    $selectedHeldCount    = 0
    $executedNormalCount  = 0
    $executedForcedCount  = 0
    $executedHeldCount    = 0
    $failedAttempts       = 0

    # Simulation logic while loop
    while ($generatedEvents -lt $numOfEvents) {

        #----------------------------------------------------------------------------------------------------
        #--> A. Write the iteration
        #----------------------------------------------------------------------------------------------------
        Write-TitleToConsole -Title "Event $($generatedEvents + 1) of $($numOfEvents)"

        #----------------------------------------------------------------------------------------------------
        #--> B. Wait a random amount of time before firing the event
        #----------------------------------------------------------------------------------------------------
        $waitTime = Get-Random -Minimum $minWaitTime -Maximum ($maxWaitTime + 1)
        Write-Host "Wait time until next event     : " -NoNewline -ForegroundColor DarkYellow
        Write-Host "$($waitTime) seconds"
        Start-Sleep $waitTime

        #----------------------------------------------------------------------------------------------------
        #--> C. Choose the type of event to attempt
        #----------------------------------------------------------------------------------------------------
        $eventRoll = Get-Random -Minimum 1 -Maximum 101

        if ($eventRoll -le $eventNormal) {
            $eventType = "Normal"
            $selectedNormalCount++
        }
        elseif ($eventRoll -le ($eventNormal + $eventForced)) {
            $eventType = "Forced"
            $selectedForcedCount++
        }
        else {
            $eventType = "Held"
            $selectedHeldCount++
        }

        Write-Host "Event type randomly selected   : " -NoNewline -ForegroundColor DarkYellow
        Write-Host "$($eventType)"

        #----------------------------------------------------------------------------------------------------
        #--> D. Randomly select a suitable door suitable for the event chosen in C above ("Forced" ends here)
        #----------------------------------------------------------------------------------------------------

        # DOOR FORCED  - Ends here 
        if ($eventType -eq "Forced") {
            
            # Find out if there are any suitable doors 
            #    - must be configured to generate door forced events and have a door sensor
            #    - must not be in maintenance mode/unlocked
            $doorResult = Simulation-GetSuitableDoor -EventType $eventType

            # If no suitable doors we print reason and start while loop again (not increasing $generatedEvents as this failed)
            if (-not $doorResult.Success) {
                Write-Host "Door selection failed          : " -NoNewline -ForegroundColor DarkYellow
                Write-Host "$($doorResult.Reason)" -ForegroundColor Red
                Write-Host ""
                Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
                Start-Sleep 2
                $failedAttempts++
                continue
            }

            # If we got a door its now as an object in this variable
            $selectedDoor = $doorResult.Door
            $openSensorDevice = $doorResult.Sensor
            
            Write-Host "Door selected                  : " -NoNewline -ForegroundColor DarkYellow
            Write-Host "$($selectedDoor.Name)"

            # if suitable door: open it, wait, close it, $generatedEvents++; continue
            Write-Host "Action                         : " -NoNewline -ForegroundColor DarkYellow

            # Does the door still exist?
            try {
                $refreshedDoor = Get-SWDoors -Session $sclSession -ErrorAction Stop | Where-Object Id -eq $selectedDoor.Id | Select-Object -First 1

                if (-not $refreshedDoor) {
                    throw "Door not found..."
                }
            }
            catch {
                Write-Host "The selected door no longer exists or could not be refreshed." -ForegroundColor Red
                Write-Host ""
                Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
                Start-Sleep 2
                $failedAttempts++
                continue
            }

            # Does the input still exist?
            try {
                $openSensorRole = $refreshedDoor.Roles |
                    Where-Object {
                        $_.Type.PSObject.Properties.Name -contains 'OpenSensor' -and
                        $_.Type.OpenSensor.Device -eq $doorResult.Sensor
                    } |
                    Select-Object -First 1

                if (-not $openSensorRole) {
                    throw "Matching sensor role not found..."
                }

                $openSensorDevice = $openSensorRole.Type.OpenSensor.Device

                if (-not $openSensorDevice) {
                    throw "Sensor device not found..."
                }

                if ($openSensorDevice -ne $doorResult.Sensor) {
                    throw "Sensor device no longer matches originally selected hardware..."
                }
            }
            catch {
                Write-Host "The selected door sensor input no longer exists, could not be refreshed, or no longer matches the originally selected input." -ForegroundColor Red
                Write-Host ""
                Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
                Start-Sleep 2
                $failedAttempts++
                continue
            }

            # Now execute the command
            try {

                Write-Host "Forcing the door..." -NoNewLine 
                Set-SWInputState -Session $sclSession -InputPointer $openSensorDevice -State Active -ErrorAction Stop | Out-Null
                Start-Sleep -Milliseconds 100
                
                Write-Host " closing the door"
                Set-SWInputState -Session $sclSession -InputPointer $openSensorDevice -State Inactive -ErrorAction Stop | Out-Null
                Start-Sleep -Milliseconds 100

                Write-Host "Result                         : " -NoNewline -ForegroundColor DarkYellow
                Write-Host "Door forced successfully" -ForegroundColor Green

                # If everything has ran successfully then increment the counters and start the next iteration
                $executedForcedCount++
                $generatedEvents++
                continue

            }
            catch {
                
                Write-Host "None" -ForegroundColor Red
                Write-Host "Failed to force door           : " -NoNewline -ForegroundColor DarkYellow
                Write-Host $_.Exception.Message -ForegroundColor Red
                Write-Host ""
                Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
                Start-Sleep 2
                $failedAttempts++
                continue

            }

        }

        # DOOR HELD    - Carries over to next step (E)
        if ($eventType -eq "Held") {

            # Find out if there are any suitable doors 
            #    - must be configured to generate door held events and have a door sensor
            #    - must not be in maintenance mode/unlocked
            $doorResult = Simulation-GetSuitableDoor -EventType $eventType

            # If no suitable doors we print reason and start while loop again (not increasing $generatedEvents as this failed)
            if (-not $doorResult.Success) {
                Write-Host "Door selection failed          : " -NoNewline -ForegroundColor DarkYellow
                Write-Host "$($doorResult.Reason)" -ForegroundColor Red
                Write-Host ""
                Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
                Start-Sleep 2
                $failedAttempts++
                continue
            }

            # If we got a door its now as an object in this variable, we also get the below returned
            <#
                Success    = $true
                Reason     = $null
                Door       = $selectedDoor
                Sensor     = $openSensorDevice
                EventType  = 'Held'
                Method     = $method              <- this will be 'Reader' or 'REX'
                Reader     = $reader
                REX        = $rex
                ReaderMode = $readerMode          <- this will be 'CardAndPin' or 'CardOnly'
                Side       = $side
            #>
            $selectedDoor = $doorResult.Door
            
            # Write out what door and device we're using
            Write-Host "Door selected                  : " -NoNewline -ForegroundColor DarkYellow
            Write-Host "$($selectedDoor.Name)"

            if ($doorResult.Method -eq 'Reader') {

                $modeText = if ($doorResult.ReaderMode -eq 'CardAndPin') {
                    'card and PIN'
                }
                else {
                    'card only'
                }

                $message = "Using $($doorResult.Side) reader ($($modeText))"
            }
            else {
                if ($doorResult.Side -ne 'None') {
                    $message = "Using $($doorResult.Side) REX"
                }
                else {
                    $message = "Using REX"
                }
            }

            Write-Host "Method selected                : " -NoNewline -ForegroundColor DarkYellow
            Write-Host $message

        }
        
        # NORMAL EVENT - Carries over to next step (E)
        if ($eventType -eq "Normal") {

            # Find out if there are any suitable doors 
            #    - must not be in maintenance mode/unlocked and have at least one reader or REX
            $doorResult = Simulation-GetSuitableDoor -EventType $eventType

            # If no suitable doors we print reason and start while loop again (not increasing $generatedEvents as this failed)
            if (-not $doorResult.Success) {
                Write-Host "Door selection failed          : " -NoNewline -ForegroundColor DarkYellow
                Write-Host "$($doorResult.Reason)" -ForegroundColor Red
                Write-Host ""
                Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
                Start-Sleep 2
                $failedAttempts++
                continue
            }

            # If we got a door its now as an object in this variable, we also get the below returned
            <#
                Success    = $true
                Reason     = $null
                Door       = $selectedDoor
                Sensor     = $openSensorDevice
                EventType  = 'Normal'
                Method     = $method              <- this will be 'Reader' or 'REX'
                Reader     = $reader
                REX        = $rex
                ReaderMode = $readerMode          <- this will be 'CardAndPin' or 'CardOnly'
                Side       = $side
            #>
            $selectedDoor = $doorResult.Door
            
            # Write out what door and device we're using
            Write-Host "Door selected                  : " -NoNewline -ForegroundColor DarkYellow
            Write-Host "$($selectedDoor.Name)"

            if ($doorResult.Method -eq 'Reader') {

                $modeText = if ($doorResult.ReaderMode -eq 'CardAndPin') {
                    'card and PIN'
                }
                else {
                    'card only'
                }

                $message = "Using $($doorResult.Side) reader ($($modeText))"
            }
            else {
                if ($doorResult.Side -ne 'None') {
                    $message = "Using $($doorResult.Side) REX"
                }
                else {
                    $message = "Using REX"
                }
            }

            Write-Host "Method selected                : " -NoNewline -ForegroundColor DarkYellow
            Write-Host $message

        }

        #----------------------------------------------------------------------------------------------------
        #--> E. Randomly select a suitable cardholder (only relevant for "Held" or "Normal" if using reader)
        #----------------------------------------------------------------------------------------------------
        if ($doorResult.Method -eq 'Reader'){

            # Get a suitable cardholder
            $cardholderResult = Simulation-GetSuitableCardholder -ReaderMode $doorResult.ReaderMode

            # If no suitable cardholders we print reason and start while loop again (not increasing $generatedEvents as this failed)
            if (-not $cardholderResult.Success) {
                Write-Host "Cardholder selection failed    : " -NoNewline -ForegroundColor DarkYellow
                Write-Host "$($cardholderResult.Reason)" -ForegroundColor Red
                Write-Host ""
                Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
                Start-Sleep 2
                $failedAttempts++
                continue
            }

            # If we got a cardholder its now as an object in this $cardholderResult variable, we also get the below returned
            <#
                Success    = $true
                Reason     = $null
                Cardholder = $selectedCardholder
                CardOnly   = $true
            #>
            
            $selectedCardholder = $cardholderResult.Cardholder

            # Now $selectedCardholder contains the below ($selectedCardholder.Table.Columns.ColumnName)
            <#
                CardholderName
                CredentialName
                RawCredential
                BitCount
                GroupName
                HasPin
            #>
           
            # Write out what door and device we're using
            Write-Host "Cardholder selected            : " -NoNewline -ForegroundColor DarkYellow
            Write-Host "$($selectedCardholder.CardholderName)"
            
        }

        #----------------------------------------------------------------------------------------------------
        #--> F. Final check on door
        #----------------------------------------------------------------------------------------------------

        <# TESTING - at this point i have the below variables
  
            #Write-Host "doorResult.Door                    : $($doorResult.Door)"
            Write-Host "doorResult.Sensor                  : $($doorResult.Sensor)"
            Write-Host "doorResult.Method                  : $($doorResult.Method)"
            Write-Host "doorResult.Reader                  : $($doorResult.Reader)"
            Write-Host "doorResult.REX                     : $($doorResult.REX)"
            Write-Host "doorResult.ReaderMode              : $($doorResult.ReaderMode)"
            Write-Host "doorResult.Side                    : $($doorResult.Side)"
            Write-Host "-----------------------------------------------------------------------------" 
            #Write-Host "selectedDoor                       : $($selectedDoor)"
            #Write-Host "-----------------------------------------------------------------------------" 
            #Write-Host "cardholderResult.Cardholder        : $($cardholderResult.Cardholder)"
            #Write-Host "cardholderResult.CardOnly          : $($cardholderResult.CardOnly)"
            #Write-Host "-----------------------------------------------------------------------------" 
            Write-Host "selectedCardholder.CardholderName  : $($selectedCardholder.CardholderName)"
            Write-Host "selectedCardholder.RawCredential   : $($selectedCardholder.RawCredential)"
            Write-Host "selectedCardholder.BitCount        : $($selectedCardholder.BitCount)"
            Start-Sleep 30

        #>

        $refreshedDoor = $null

        # STEP 1 - Does the door still exist?
        try {
            $refreshedDoor = Get-SWDoors -Session $sclSession -ErrorAction Stop | Where-Object Id -eq $selectedDoor.Id | Select-Object -First 1

            if (-not $refreshedDoor) {
                throw "Door not found..."
            }
        }
        catch {
            Write-Host "Error                          : " -NoNewline -ForegroundColor DarkYellow
            Write-Host "The selected door no longer exists or could not be refreshed." -ForegroundColor Red
            Write-Host ""
            Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
            Start-Sleep 2
            $failedAttempts++
            continue
        }

        # STEP 2 - Is door in maintenance mode?
        try {

            $maintenanceMode = $refreshedDoor.UnlockedForMaintenance
            $doorLocked = $refreshedDoor.IsLocked

            if ($maintenanceMode -eq $true) {
                throw "Door in maintenance mode..."
            }

            if ($doorLocked -eq $false) {
                throw "Door is unlocked..."
            }
        }
        catch {
            Write-Host "Error                          : " -NoNewline -ForegroundColor DarkYellow
            Write-Host "The selected door is now unlocked and/or in maintenance mode." -ForegroundColor Red
            Write-Host ""
            Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
            Start-Sleep 2
            $failedAttempts++
            continue
        }

        # STEP 3 - Check the refreshed door still has the correct hardware as per the variable $doorResult

        # 3.a - Checking Door Sensor (if applicable)
        if ($doorResult.Sensor -ne $null) {

            # Does the door sensor still exist?
            try {
                $openSensorRole = $refreshedDoor.Roles |
                    Where-Object {
                        $_.Type.PSObject.Properties.Name -contains 'OpenSensor' -and
                        $_.Type.OpenSensor.Device -eq $doorResult.Sensor
                    } |
                    Select-Object -First 1

                if (-not $openSensorRole) {
                    throw "Matching sensor role not found..."
                }

                $openSensorDevice = $openSensorRole.Type.OpenSensor.Device

                if (-not $openSensorDevice) {
                    throw "Sensor device not found..."
                }

                if ($openSensorDevice -ne $doorResult.Sensor) {
                    throw "Sensor device no longer matches originally selected hardware..."
                }
            }
            catch {
                Write-Host "Error                          : " -NoNewline -ForegroundColor DarkYellow
                Write-Host "The selected door sensor input no longer exists, could not be refreshed, or no longer matches the originally selected input." -ForegroundColor Red
                Write-Host ""
                Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
                Start-Sleep 2
                $failedAttempts++
                continue
            }

        }

        # 3.b - Checking REX (if applicable)
        if ($doorResult.Method -eq 'REX') {

            # Does the REX still exist?
            try {
                $rexRole = $refreshedDoor.Roles |
                    Where-Object {
                        $_.Type.PSObject.Properties.Name -contains 'REX' -and
                        $_.Type.REX.Device -eq $doorResult.REX
                    } |
                    Select-Object -First 1

                if (-not $rexRole) {
                    throw "Matching REX role not found..."
                }

                $rexDevice = $rexRole.Type.REX.Device

                if (-not $rexDevice) {
                    throw "REX device not found..."
                }

                if ($rexDevice -ne $doorResult.REX) {
                    throw "REX device no longer matches originally selected hardware..."
                }
            }
            catch {
                Write-Host "Error                          : " -NoNewline -ForegroundColor DarkYellow
                Write-Host "The selected REX input no longer exists, could not be refreshed, or no longer matches the originally selected input." -ForegroundColor Red
                Write-Host ""
                Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
                Start-Sleep 2
                $failedAttempts++
                continue
            }

        }

        # 3.c - Checking Reader (if applicable)
        if ($doorResult.Method -eq 'Reader') {

            try {
                # Convert pretty side back to raw side
                $rawSide = switch ($doorResult.Side) {
                    'Entry' { 'A' }
                    'Exit'  { 'B' }
                    'None'  { 'NA' }
                }

                $readerRole = $refreshedDoor.Roles |
                    Where-Object {
                        $_.Type.PSObject.Properties.Name -contains 'ReaderAuth' -and
                        $_.Side.PSObject.Properties.Name -eq $rawSide
                    } |
                    Select-Object -First 1

                if (-not $readerRole) {
                    throw "Reader not found... was it deleted?"
                }

                $readerDevice = $readerRole.Type.ReaderAuth.HardwareReader

                if (-not $readerDevice) {
                    throw "Reader not found... was it deleted?"
                }

                if ($readerDevice -ne $doorResult.Reader) {
                    throw "Reader no longer matches originally selected hardware... was it changed?"
                }

                # Check Reader Mode
                $currentMode = $readerRole.Type.ReaderAuth.ReaderMode.PSObject.Properties.Name

                if ($doorResult.ReaderMode -eq 'CardAndPin') {
                    if (-not ($currentMode -contains 'CardAndPin')) {
                        throw "Reader mode no longer matches (expected Card + Pin it's now Card Only)..."
                    }
                }
                else {
                    if ($currentMode -contains 'CardAndPin') {
                        throw "Reader mode no longer matches (expected Card Only it's now Card + Pin)..."
                    }
                }
            }
            catch {
                Write-Host "Error                          : " -NoNewline -ForegroundColor DarkYellow
                Write-Host $_.Exception.Message -ForegroundColor Red
                Write-Host ""
                Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
                Start-Sleep 2
                $failedAttempts++
                continue
            }
        }

        # 3.d - We don't care about if the cardholder has been deleted, it will just be Access Denied: Unknown Credential

        #----------------------------------------------------------------------------------------------------
        #--> G. Execute command (including checking access and incrementing counters)
        #----------------------------------------------------------------------------------------------------
        Write-Host "Action                         : " -NoNewline -ForegroundColor DarkYellow

        # [1/2] REX is used
        if ($doorResult.Method -eq 'REX') {

            Write-Host "Pressing REX..." -NoNewLine -ForegroundColor Green

            # Simulate pressing then releasing the REX             
            Set-SWInputState -Session $sclSession -InputPointer $doorResult.REX -State Active   -ErrorAction Stop | Out-Null
            Start-Sleep -Milliseconds 100
            Set-SWInputState -Session $sclSession -InputPointer $doorResult.REX -State Inactive -ErrorAction Stop | Out-Null
            Start-Sleep -Milliseconds 100

            # Refresh door object to check access decision
            try {
                $refreshedDoorForDecision = Get-SWDoors -Session $sclSession -ErrorAction Stop | Where-Object Id -eq $selectedDoor.Id | Select-Object -First 1

                if (-not $refreshedDoorForDecision) {
                    throw "Door not found..."
                }
            }
            catch {
                Write-Host "Error                          : " -NoNewline -ForegroundColor DarkYellow
                Write-Host "The selected door no longer exists or could not be refreshed." -ForegroundColor Red
                Write-Host ""
                Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
                Start-Sleep 2
                $failedAttempts++
                continue
            }

            # Now handle door sensor depending on event type (normal or held) and if access was granted or not
            if (
                (-not $refreshedDoorForDecision.IsLocked) -and
                ($refreshedDoorForDecision.LastDecision.Decision.Granted.Reason.PSObject.Properties.Name -contains 'RexActivated')
            ) {

                Write-Host "Access granted..." -NoNewLine -ForegroundColor Green

                if ($eventType -eq "Normal") {

                    # If the door has a sensor, open and close it
                    if ($doorResult.Sensor -ne $null) {

                        Write-Host "Opening door..." -NoNewLine -ForegroundColor Green
                        Set-SWInputState -Session $sclSession -InputPointer $doorResult.Sensor -State Active   -ErrorAction Stop | Out-Null
                        Start-Sleep -Milliseconds 100
                        Write-Host "Closing door..." -ForegroundColor Green
                        Set-SWInputState -Session $sclSession -InputPointer $doorResult.Sensor -State Inactive -ErrorAction Stop | Out-Null
                        Start-Sleep -Milliseconds 100

                        $executedNormalCount++
                        $generatedEvents++
                        continue

                    }

                    # If the door doesnt have a sensor... move on
                    Write-Host "No door sensor to operate!" -ForegroundColor Yellow
                    $executedNormalCount++
                    $generatedEvents++
                    continue

                }

                if ($eventType -eq "Held") {

                    # If the door has a sensor, open and close it
                    if ($doorResult.Sensor -ne $null) {

                        Write-Host "Opening door..." -NoNewLine -ForegroundColor Green
                        Set-SWInputState -Session $sclSession -InputPointer $doorResult.Sensor -State Active   -ErrorAction Stop | Out-Null
                        Start-Sleep -Milliseconds 100
                        Write-Host "And leaving it open!" -ForegroundColor Yellow
                        Start-Sleep -Milliseconds 100

                        $executedHeldCount++
                        $generatedEvents++
                        continue

                    }

                    # If the door doesnt have a sensor... move on
                    Write-Host "No door sensor to operate!" -ForegroundColor Yellow
                    $executedHeldCount++
                    $generatedEvents++
                    continue

                }

                Write-Host "Error                          : " -NoNewline -ForegroundColor DarkYellow
                Write-Host "Something went wrong!" -ForegroundColor Red
                Write-Host ""
                Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
                Start-Sleep 2
                $failedAttempts++
                continue

            }
            else {

                Write-Host "Access denied..." -NoNewLine -ForegroundColor Red

                if ($eventType -eq "Normal") {

                    Write-Host "Cannot open door!" -ForegroundColor Yellow
                    $executedNormalCount++
                    $generatedEvents++
                    continue

                }

                if ($eventType -eq "Held") {

                    Write-Host "Cannot open door!" -ForegroundColor Yellow
                    $executedHeldCount++
                    $generatedEvents++
                    continue

                }

                Write-Host "Error                          : " -NoNewline -ForegroundColor DarkYellow
                Write-Host "Something went wrong!" -ForegroundColor Red
                Write-Host ""
                Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
                Start-Sleep 2
                $failedAttempts++
                continue

            }

        }

        # [2/2] Reader is used
        if ($doorResult.Method -eq 'Reader') {
            
            # [1.A] Carrying out the Read - Card + PIN reader
            if ($doorResult.ReaderMode -eq 'CardAndPin') {

                # Prepare raw credential (strip left padding based on entered bit count) - there seems to be a bug with Card & PIN only where you have to do this?
                $cleanHex = ($selectedCardholder.RawCredential -replace '^0x', '' -replace '\s+', '')
                $bytesNeeded = [math]::Ceiling($selectedCardholder.BitCount / 8)
                $hexCharsNeeded = $bytesNeeded * 2
                $rawToUse = $cleanHex.Substring($cleanHex.Length - $hexCharsNeeded).ToUpper()

                # Card first.
                try {
            
                    Write-Host "Swiping the card..." -NoNewline -ForegroundColor Green

                    Invoke-SWSwipeRaw -Session $sclSession -BitCount $selectedCardholder.BitCount -Bytes $rawToUse -ReaderPointer $doorResult.Reader -ErrorAction Stop | Out-Null
                    
                    Start-Sleep -Milliseconds 100

                }
                catch {

                    Write-Host "Error message: $($_.Exception.Message)" -ForegroundColor Red
                    Start-Sleep 1
                    $failedAttempts++
                    continue

                }

                # Then PIN.
                try {
            
                    Write-Host "Entering the PIN..." -NoNewline -ForegroundColor Green

                    Start-Sleep -Milliseconds 100

                    Invoke-SWSwipeWiegand26 -Session $sclSession -Facility 00 -Card $globalPin -ReaderPointer $doorResult.Reader -ErrorAction Stop | Out-Null
            
                    Start-Sleep -Milliseconds 100

                }
                catch {

                    Write-Host "Error message: $($_.Exception.Message)" -ForegroundColor Red
                    Start-Sleep 1
                    $failedAttempts++
                    continue

                }

            }

            # [1.B] Carrying out the Read - Card Only reader
            if ($doorResult.ReaderMode -eq 'CardOnly') {

                # Card only.
                try {
            
                    Write-Host "Swiping the card..." -NoNewline -ForegroundColor Green

                    Invoke-SWSwipeRaw -Session $sclSession -BitCount $selectedCardholder.BitCount -Bytes $selectedCardholder.RawCredential -ReaderPointer $doorResult.Reader -ErrorAction Stop | Out-Null
                    
                    Start-Sleep -Milliseconds 100

                }
                catch {

                    Write-Host "Error message: $($_.Exception.Message)" -ForegroundColor Red
                    Start-Sleep 1
                    $failedAttempts++
                    continue

                }

            }

            # [2.A] Refresh door
            try {
                $refreshedDoorForDecision = Get-SWDoors -Session $sclSession -ErrorAction Stop | Where-Object Id -eq $selectedDoor.Id | Select-Object -First 1

                if (-not $refreshedDoorForDecision) {
                    throw "Door not found..."
                }
            }
            catch {
                Write-Host "Error                          : " -NoNewline -ForegroundColor DarkYellow
                Write-Host "The selected door no longer exists or could not be refreshed." -ForegroundColor Red
                Write-Host ""
                Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
                Start-Sleep 2
                $failedAttempts++
                continue
            }

            # [2.B] Check access and use door sensor as required
            $decision = $refreshedDoorForDecision.LastDecision.Decision

            # Access granted
            if ($null -ne $decision.Granted) {

                # Access granted
                Write-Host "Access granted..." -NoNewLine -ForegroundColor Green

                if ($eventType -eq "Normal") {

                    # If the door has a sensor, open and close it
                    if ($doorResult.Sensor -ne $null) {

                        Write-Host "Opening door..." -NoNewLine -ForegroundColor Green
                        Set-SWInputState -Session $sclSession -InputPointer $doorResult.Sensor -State Active   -ErrorAction Stop | Out-Null
                        Start-Sleep -Milliseconds 100
                        Write-Host "Closing door..." -ForegroundColor Green
                        Set-SWInputState -Session $sclSession -InputPointer $doorResult.Sensor -State Inactive -ErrorAction Stop | Out-Null
                        Start-Sleep -Milliseconds 100

                        $executedNormalCount++
                        $generatedEvents++
                        continue

                    }

                    # If the door doesnt have a sensor... move on
                    Write-Host "No door sensor to operate!" -ForegroundColor Yellow
                    $executedNormalCount++
                    $generatedEvents++
                    continue

                }

                if ($eventType -eq "Held") {

                    # If the door has a sensor, open and close it
                    if ($doorResult.Sensor -ne $null) {

                        Write-Host "Opening door..." -NoNewLine -ForegroundColor Green
                        Set-SWInputState -Session $sclSession -InputPointer $doorResult.Sensor -State Active   -ErrorAction Stop | Out-Null
                        Start-Sleep -Milliseconds 100
                        Write-Host "And leaving it open!" -ForegroundColor Yellow
                        Start-Sleep -Milliseconds 100

                        $executedHeldCount++
                        $generatedEvents++
                        continue

                    }

                    # If the door doesnt have a sensor... move on
                    Write-Host "No door sensor to operate!" -ForegroundColor Yellow
                    $executedHeldCount++
                    $generatedEvents++
                    continue

                }

                Write-Host "Error                          : " -NoNewline -ForegroundColor DarkYellow
                Write-Host "Something went wrong!" -ForegroundColor Red
                Write-Host ""
                Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
                Start-Sleep 2
                $failedAttempts++
                continue

            }
            # Access denied
            elseif ($null -ne $decision.Denied) {

                # Access denied
                Write-Host "Access denied..." -NoNewLine -ForegroundColor Red

                if ($eventType -eq "Normal") {

                    Write-Host "Cannot open door!" -ForegroundColor Yellow
                    $executedNormalCount++
                    $generatedEvents++
                    continue

                }

                if ($eventType -eq "Held") {

                    Write-Host "Cannot open door!" -ForegroundColor Yellow
                    $executedHeldCount++
                    $generatedEvents++
                    continue

                }

                Write-Host "Error                          : " -NoNewline -ForegroundColor DarkYellow
                Write-Host "Something went wrong!" -ForegroundColor Red
                Write-Host ""
                Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
                Start-Sleep 2
                $failedAttempts++
                continue

            }
            # No decision or weird state
            else {

                # No decision / weird state
                Write-Host "Something unexpected happened!" -ForegroundColor Red
                Write-Host ""
                Write-Host "We will retry this iteration again." -ForegroundColor DarkGray
                Start-Sleep 2
                $failedAttempts++
                continue

            }

        }
          
    } 

    #--------------------------------------------------------------------------------------------------------
    #--> 5. Once done, show a summary and press enter to 'return' to main menu
    #--------------------------------------------------------------------------------------------------------
    # Totals
    $totalExecuted = $executedNormalCount + $executedForcedCount + $executedHeldCount

    # Percent helpers
    function Get-Pct($part, $whole) {
        if ($whole -eq 0) { return 0 }
        return [math]::Round(($part / $whole) * 100, 2)
    }

    # Selection %
    $normalSelPct = Get-Pct $selectedNormalCount $numOfEvents
    $forcedSelPct = Get-Pct $selectedForcedCount $numOfEvents
    $heldSelPct   = Get-Pct $selectedHeldCount   $numOfEvents

    # Execution %
    $normalExePct = Get-Pct $executedNormalCount $totalExecuted
    $forcedExePct = Get-Pct $executedForcedCount $totalExecuted
    $heldExePct   = Get-Pct $executedHeldCount   $totalExecuted

    # Success rates (selected vs executed)
    $normalSuccess = Get-Pct $executedNormalCount $selectedNormalCount
    $forcedSuccess = Get-Pct $executedForcedCount $selectedForcedCount
    $heldSuccess   = Get-Pct $executedHeldCount   $selectedHeldCount

    # Drop-offs
    $normalDrop = $selectedNormalCount - $executedNormalCount
    $forcedDrop = $selectedForcedCount - $executedForcedCount
    $heldDrop   = $selectedHeldCount   - $executedHeldCount

    Write-Host ""
    Write-Host "========== SIMULATION SUMMARY ==========" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "Requested events       : $numOfEvents"
    Write-Host "Successfully executed  : $totalExecuted"
    Write-Host "Failed / retried       : $failedAttempts"

    Write-Host ""
    Write-Host "------------ SELECTION ------------" -ForegroundColor Cyan
    Write-Host "Normal                 : $selectedNormalCount ($normalSelPct%)"
    Write-Host "Forced                 : $selectedForcedCount ($forcedSelPct%)"
    Write-Host "Held                   : $selectedHeldCount ($heldSelPct%)"

    Write-Host ""
    Write-Host "------------ EXECUTION ------------" -ForegroundColor Cyan
    Write-Host "Normal                 : $executedNormalCount ($normalExePct%)  | Success: $normalSuccess%"
    Write-Host "Forced                 : $executedForcedCount ($forcedExePct%)  | Success: $forcedSuccess%"
    Write-Host "Held                   : $executedHeldCount ($heldExePct%)  | Success: $heldSuccess%"

    Write-Host ""
    Write-Host "---------- EFFECTIVENESS ----------" -ForegroundColor Cyan
    Write-Host "Normal drop-off        : $normalDrop"
    Write-Host "Forced drop-off        : $forcedDrop"
    Write-Host "Held drop-off          : $heldDrop"

    Write-Host ""
    $response = (Read-Host "Press ENTER to go back to the previous menu").Trim().ToUpperInvariant()

    ClearConsole-Pretty -Message "The console will clear in"

}

#------------------------------------------------------------------------------------------------------
#--> [b] MISC Helper Functions - used by both of the Main Menu Functions above
#------------------------------------------------------------------------------------------------------
function ClearConsole-Pretty {
    <#
    Purpose:
        - Clears the console in a really 'fun' way
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Message
    )

    Write-Host ""

    for ($i = 3; $i -ge 1; $i--) {
        Write-Host "$Message [$i] " -ForegroundColor Yellow
        Start-Sleep -Milliseconds 500
    }

    Write-Host "$Message [NOW] " -ForegroundColor Green
    Start-Sleep -Milliseconds 500

    Clear-Host
}

function Connect-Softwire {
    <#
    Purpose:
        - Ask user for the Softwire password
        - Keep prompting until login succeeds
        - Return both password and session
    #>

    while ($true) {
        Write-Host "`nPlease enter the admin password for Softwire:" -ForegroundColor Cyan
        $softwirePassword = Read-Host

        try {
            $session = Invoke-SWLogin -Hostname $sclEndpoint -Username $sclUsername -Password $softwirePassword -ErrorAction Stop

            Write-Host "`nLogin suceeded. Password is correct." -ForegroundColor Green
            Start-Sleep 1
            
            ClearConsole-Pretty -Message "The console will wipe the password in"

            return [PSCustomObject]@{
                Password = $softwirePassword
                Session  = $session
            }
        }
        catch {
            Write-Host "`nLogin failed. Please try again." -ForegroundColor Red
        }
    }
} 

function LandingPage {
    <#
    Purpose:
        - Every good script needs a Landing page... right?
        - At the very least, a good script needs to outline what's to happen (even if it's a BETA :)
    #>

    ShowLandingBanner

    # Author + last updated
    Write-Host "Author       : James Savage" -ForegroundColor DarkGray
    Write-Host "Last updated : 27-Mar-2026" -ForegroundColor DarkGray
    Write-Host "Version      : BETA 1.0 (ALPHA WIP)" -ForegroundColor DarkGray
    Write-Host ""

    # Quick summary (short enough that they’ll actually read it)
    Write-Host "This tool will:" -ForegroundColor Cyan
    Write-Host "---------------" -ForegroundColor Cyan
    Write-Host "  - Let you interact with simulated doors manually" -ForegroundColor Gray
    Write-Host "  - Simulate a busy site automatically for demos / training / stress testing" -ForegroundColor Gray
    Write-Host ""

    # Important notes (don't say I didn't warn you...)
    Write-Host "Important notes:" -ForegroundColor Yellow
    Write-Host "----------------" -ForegroundColor Yellow
    Write-Host "  - This is only a BETA/proof of Concept to prove the logic used in this script" -ForegroundColor DarkYellow
    Write-Host "      o The release (Alpha) will be wrapped in a nice interactive GUI (will still be janky though)" -ForegroundColor Gray
    Write-Host "  - Please make sure before pressing enter:" -ForegroundColor DarkYellow
    Write-Host "      o Your system is licensed to support Softwire" -ForegroundColor Gray
    Write-Host "      o Your system has Softwire installed, configured, and added to the Access Manager" -ForegroundColor Gray
    Write-Host "      o You have enabled the simulator in Softwire (https://127.0.0.1/Softwire/Features/duisim/Enabled/Set?value=true)" -ForegroundColor Gray
    Write-Host "      o You have added simulated hardware and configured doors using it" -ForegroundColor Gray
    Write-Host "      o You have Cardholders, Credentials, and Access Rules" -ForegroundColor Gray
    Write-Host ""

    # Waiting for enter key to be pressed
    $response = (Read-Host "press ENTER to continue").Trim().ToUpperInvariant()

    Write-Host ""
}

function ShowLandingBanner {
    <#
    Purpose:
        - None... Well, its for the landing page and I think it looks "cool"
    #>

$art = @'
 _____   _____   _____   _____   _    _   _____   _____   _____           _____   _____   ___  ___ 
/  ___| |  _  | |  ___| |_   _| | |  | | |_   _| | ___ | |  ___|         /  ___| |_   _| |  \/  | 
\ `--.  | | | | | |_      | |   | |  | |   | |   | |_/ / | |__           \ `--.    | |   | .  . | 
 `--. \ | | | | |  _|     | |   | |/\| |   | |   |    /  |  __|           `--. \   | |   | |\/| | 
/\__/ / | |_| | | |       | |   \  /\  /  _| |_  | |\ \  | |___          /\__/ /  _| |_  | |  | | 
\____/  |_____| \_|       \_/    \/  \/  |_____| \_| \_| \____/          \____/  |_____| \_|  |_/ 
'@

    $art -split "`n" | ForEach-Object {
        Write-Host $_ -ForegroundColor Magenta
    }

    # ---------------- FRAME ----------------
    $frameWidth = 97
    $innerWidth = $frameWidth - 4
    $line       = '-' * ($frameWidth - 2)

    # Helper: centre text within the inner frame width
    function Center-Line {
        param([Parameter(Mandatory)][string]$Text)

        if ($Text.Length -ge $innerWidth) {
            return $Text.Substring(0, $innerWidth)
        }

        $leftPad = [Math]::Floor(($innerWidth - $Text.Length) / 2)
        ((' ' * $leftPad) + $Text).PadRight($innerWidth)
    }

    # Top border
    Write-Host "+" -ForegroundColor DarkGray -NoNewline
    Write-Host $line -ForegroundColor DarkGray -NoNewline
    Write-Host "+" -ForegroundColor DarkGray

    # Title line
    $title = "Softwire SIM (Proof of Concept)"
    Write-Host "| " -ForegroundColor DarkGray -NoNewline
    Write-Host (Center-Line -Text $title) -ForegroundColor Magenta -NoNewline
    Write-Host " |" -ForegroundColor DarkGray

    # Tagline line
    $tagline = '"Because using the built in Softwire SIM interface is a crime (but probably easier)"'
    Write-Host "| " -ForegroundColor DarkGray -NoNewline
    Write-Host (Center-Line -Text $tagline) -ForegroundColor DarkYellow -NoNewline
    Write-Host " |" -ForegroundColor DarkGray

    # Bottom border
    Write-Host "+" -ForegroundColor DarkGray -NoNewline
    Write-Host $line -ForegroundColor DarkGray -NoNewline
    Write-Host "+" -ForegroundColor DarkGray

    Write-Host ""
} 

function Write-TitleToConsole {
    <# 
      Write a nice header to the console for each menu selection
    #>
    param(
        [Parameter(Mandatory)] $Title
    )

        # ---------------- FRAME ----------------
    $frameWidth = 80
    $innerWidth = $frameWidth - 4
    $line       = '-' * ($frameWidth - 2)

    # Helper: centre text within the inner frame width
    function Center-Line {
        param([Parameter(Mandatory)][string]$Text)

        if ($Text.Length -ge $innerWidth) {
            return $Text.Substring(0, $innerWidth)
        }

        $leftPad = [Math]::Floor(($innerWidth - $Text.Length) / 2)
        ((' ' * $leftPad) + $Text).PadRight($innerWidth)
    }

    Write-Host ""

    # Top border
    Write-Host "+" -ForegroundColor DarkGray -NoNewline
    Write-Host $line -ForegroundColor DarkGray -NoNewline
    Write-Host "+" -ForegroundColor DarkGray

    # Title line
    Write-Host "| " -ForegroundColor DarkGray -NoNewline
    Write-Host (Center-Line -Text $Title) -ForegroundColor Magenta -NoNewline
    Write-Host " |" -ForegroundColor DarkGray

    # Bottom border
    Write-Host "+" -ForegroundColor DarkGray -NoNewline
    Write-Host $line -ForegroundColor DarkGray -NoNewline
    Write-Host "+" -ForegroundColor DarkGray

    Write-Host ""

} 

#------------------------------------------------------------------------------------------------------
#--> [c] MAIN Helper Functions - used by the main menu functions
#------------------------------------------------------------------------------------------------------
function Get-AllCardholders {
    <#
      This function will return all cardholders, their credentials, and Cardholder Group(s) they belong to. 
      
      Note: We dont show PINs and/or ANPR:
        - PINs will be entered manually, no need to show them.
        - ANPR (number plates) won't work with the SIM so no need for them to show.

      Also note: this was a pain in the arse as 'Get-SWCardholders' doesn't work for simulated devices (they don't sync)
                 so went for some really 'janky' SQL querying... I just know this will break one day...
    #>
    param(
        [string]$SqlServerInstance = "localhost\SQLExpress",
        [string]$Database = "Directory"
    )

    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $sqlConnection.ConnectionString = "Server=$SqlServerInstance;Database=$Database;Integrated Security=true"

    try {
        $sqlConnection.Open()

        $sqlCommand = $sqlConnection.CreateCommand()
        $sqlCommand.CommandText = @"
SELECT
    ce.Name AS CardholderName,
    cre.Name AS CredentialName,
    LEFT(cr.UniqueID, CHARINDEX('|', cr.UniqueID) - 1) AS RawCredential,
    RIGHT(cr.UniqueID, LEN(cr.UniqueID) - CHARINDEX('|', cr.UniqueID)) AS BitCount,
    grp.GroupNames AS GroupName,
    CASE
        WHEN EXISTS (
            SELECT 1
            FROM Credential crCheck
            WHERE crCheck.Cardholder = c.Guid
              AND NOT (
                    CHARINDEX('|', crCheck.UniqueID) > 0
                    AND TRY_CONVERT(int, RIGHT(crCheck.UniqueID, LEN(crCheck.UniqueID) - CHARINDEX('|', crCheck.UniqueID))) IS NOT NULL
              )
              AND crCheck.UniqueID NOT LIKE 'Plate%'
        )
        THEN 'True'
        ELSE 'False'
    END AS HasPin
FROM Cardholder c
LEFT JOIN Entity ce
    ON c.Guid = ce.Guid
LEFT JOIN Credential cr
    ON c.Guid = cr.Cardholder
    AND CHARINDEX('|', cr.UniqueID) > 0
    AND TRY_CONVERT(int, RIGHT(cr.UniqueID, LEN(cr.UniqueID) - CHARINDEX('|', cr.UniqueID))) IS NOT NULL
LEFT JOIN Entity cre
    ON cr.Guid = cre.Guid
LEFT JOIN (
    SELECT
        cm.GuidMember,
        STUFF((
            SELECT ', ' + ge2.Name
            FROM CardholderMembership cm2
            LEFT JOIN CardholderGroup cg2
                ON cm2.GuidGroup = cg2.Guid
            LEFT JOIN Entity ge2
                ON cg2.Guid = ge2.Guid
            WHERE cm2.GuidMember = cm.GuidMember
            FOR XML PATH(''), TYPE
        ).value('.', 'nvarchar(max)'), 1, 2, '') AS GroupNames
    FROM CardholderMembership cm
    GROUP BY cm.GuidMember
) grp
    ON c.Guid = grp.GuidMember
ORDER BY ce.Name
"@

        $sqlDataAdapter = New-Object System.Data.SqlClient.SqlDataAdapter $sqlCommand
        $dataSet = New-Object System.Data.DataSet
        $sqlDataAdapter.Fill($dataSet) | Out-Null

        # Return the table (the headers are: CardholderName, CredentialName, RawCredential, BitCount, GroupName, HasPin)
        return ,$dataSet.Tables[0]
    }
    finally {
        if ($sqlConnection.State -eq 'Open') {
            $sqlConnection.Close()
        }

        $sqlConnection.Dispose()
    }
} 

function Get-Doors {
    <# 
      Get a full list of the doors - refreshes each time it is called
    #>
    Get-SWDoors -Session $sclSession
} 

function Get-DoorHardware {
    <# 
      Get all associated door hardware - refreshes each time it is called
    #>
    param(
        [Parameter(Mandatory)]
        $Door
    )

    foreach ($role in $Door.Roles) {
        $typeName = $role.Type.PSObject.Properties.Name
        $sideName = $role.Side.PSObject.Properties.Name
        $typeData = $role.Type.$typeName

        $device = switch ($typeName) {
            'ReaderAuth'    { $typeData.HardwareReader }
            'REX'           { $typeData.Device }
            'Strike'        { $typeData.Device }
            'OpenSensor'    { $typeData.Device }
            'ManualStation' { $typeData.Device }
            'PassageSensor' { $typeData.Device }
            default         { $null }
        }

        $readerMode = if ($typeName -eq 'ReaderAuth') {
            $typeData.ReaderMode.PSObject.Properties.Name -join ','
        }

        $readerType = if ($typeName -eq 'ReaderAuth') {
            $typeData.ReaderType.PSObject.Properties.Name -join ','
        }

        [pscustomobject]@{
            Door       = $Door.Name
            Side       = $sideName
            RoleType   = $typeName
            Device     = $device
            ReaderMode = $readerMode
            ReaderType = $readerType
        }
    }
} 

function Get-DoorInputState {
    <#
      This function will return the door input state. 
    #>
    param(
        [Parameter(Mandatory)]
        $InputState
    )

    switch ($true) {
        $InputState.IsShunted      { return "Shunted" }
        $InputState.TroubleUnknown { return "Trouble Unknown" }
        $InputState.Cut            { return "Line Cut" }
        $InputState.Shorted        { return "Short Circuit" }
        $InputState.Active         { return "Open" }
        default                    { return "Closed" }
    }
} 

function Get-DoorLastEvent {
    <#
    Purpose:
        - This checks the last event of the door.
    #>
    param(
        [Parameter(Mandatory)] $Door
    )

    # Need to refresh the door so the last decision is updated (if it fails it will kick them to Main Menu)!
    $door = Refresh-CurrentDoor -Door $Door

    <# 
        TESTING:

        $test = $door | ConvertTo-Json -Depth 10
        #$door.LastDecision | ConvertTo-Json -Depth 10

        # Testing
        #$test = $door.LastDecision.Decision | ConvertTo-Json -Depth 10
        Write-Host "RESULT: $($test)"
    #>

    $decision = $door.LastDecision.Decision

    if ($null -ne $decision.Granted) {

        Write-Host "Access Granted" -ForegroundColor Green
        $accessGranted = $true

    }
    elseif ($null -ne $decision.Denied) {

        $reason = $decision.Denied.PSObject.Properties.Name

        Write-Host "Access Denied: $reason" -ForegroundColor Red

        $accessGranted = $false

    }

    return $accessGranted

} 

function Get-DoorOutputState {
    <#
      This function will return the door output state. 
    #>
    param(
        [Parameter(Mandatory)]
        $OutputState
    )

    switch ($true) {
        $OutputState.IsShunted { return "Shunted" }
        $OutputState.Activated { return "Unlocked" }
        default                { return "Locked" }
    }
} 

function Get-DoorReaderLedColor {
    <#
      This function will return the door reader LED colour. 
    #>
    param(
        [Parameter(Mandatory)]
        $ReaderLed
    )

    if (-not $ReaderLed.LedColor) {
        return "Unknown"
    }

    switch ($true) {
        ($ReaderLed.LedColor.PSObject.Properties.Name -contains 'Red')   { return "Red" }
        ($ReaderLed.LedColor.PSObject.Properties.Name -contains 'Green') { return "Green" }
        default                                                          { return "Unknown" }
    }
} 

function Get-DoorReaderState {
    <#
      This function will return the door reader state. 
    #>
    param(
        [Parameter(Mandatory)]
        $ReaderState
    )

    switch ($true) {
        $ReaderState.IsShunted { return "Shunted" }
        $ReaderState.Online    { return "Online" }
        default                { return "Offline" }
    }
} 

function Get-DoorState {
    <#
    Purpose:
        - This checks if the door is in maintenance mode or unlocked.
    #>
    param(
        [Parameter(Mandatory)] $Door
    )

    # is the door unlocked?
    $isDoorLocked = $Door.IsLocked
    
    # is the door in maintenance mode?
    $isDoorInMaintenanceMode = $Door.UnlockedForMaintenance

    # Return door status
    return [pscustomobject]@{
        IsDoorLocked             = $isDoorLocked
        IsDoorInMaintenanceMode  = $isDoorInMaintenanceMode
    }

} 

function Get-PinCredential {
    <#
    Purpose:
        - Allow the user to enter a PIN and return it
    #>
    param(
        [Parameter(Mandatory)] $Message
    )

    Write-Host $Message

    while ($true) {

        Write-Host ""
        Write-Host "Enter the cardholder's PIN (4–5 digits): " -NoNewline -ForegroundColor Cyan
        $cardholdersPin = Read-Host

        if ($cardholdersPin -notmatch '^\d{4,5}$') {
            Write-Host "Invalid PIN. Please enter a 4 or 5 digit PIN." -ForegroundColor Red
            Start-Sleep 1
            continue
        }

        # Pad to 5 digits if needed (SDK/API needs it to be 5 digits always)
        if ($cardholdersPin.Length -eq 4) {
            $cardholdersPin = "0$cardholdersPin"
        }

        return $cardholdersPin
    }

} 

function Get-RawCredential {
    <#
    Purpose:
        - Allow the user to enter a raw credential and return it
        - This function is called from:
            - Manual-ManualEntryCredential (any return in this function will go back there)
    #>

    # Get the bit count ($manuallyEnteredBitCount)
    while ($true) {

        Write-Host "`nWhat 'Bit Count' do you want to use for your card/credential?"
        Write-Host ""
        Write-Host "[1] Standard " -NoNewLine
        Write-Host "26-bit" -ForegroundColor DarkYellow -NoNewLine
        Write-Host " Wiegand with Facility Code and Card Number"
        Write-Host "[2] Standard " -NoNewLine
        Write-Host "32-bit" -ForegroundColor DarkYellow -NoNewLine
        Write-Host " CSN/UID"
        Write-Host "[3] Custom bit-length (" -NoNewLine
        Write-Host "1" -ForegroundColor DarkYellow -NoNewLine
        Write-Host " to " -NoNewLine
        Write-Host "256" -ForegroundColor DarkYellow -NoNewLine
        Write-Host ")"
        Write-Host ""
    
        Write-Host "Enter your choice: " -NoNewline -ForegroundColor Cyan
        $choiceBitCount = (Read-Host).Trim()

        if ($choiceBitCount -eq '1') {
            $manuallyEnteredBitCount = '26'
            break
        }

        if ($choiceBitCount -eq '2') {
            $manuallyEnteredBitCount = '32'
            break
        }

        if ($choiceBitCount -eq '3') {

            while ($true) {
                Write-Host "`nEnter the required bit count (1 to 256): " -NoNewline -ForegroundColor Cyan
                $manuallyEnteredBitCount = (Read-Host).Trim()

                if ($manuallyEnteredBitCount -notmatch '^\d+$' -or [int]$manuallyEnteredBitCount -lt 1 -or [int]$manuallyEnteredBitCount -gt 256) {
                    Write-Host "Invalid choice. Please enter a number from 1 to 256." -ForegroundColor Red
                    Start-Sleep 1
                    continue
                }

                break
            }

            break
        }

        # If the user is unable to enter, 1, 2, or 3, back round we go!
        Write-Host "Invalid choice. Please enter 1, 2, or 3." -ForegroundColor Red
        Start-Sleep 1
    }

    # Get the raw credential ($manuallyEnteredRawCredential)

    # [1] If user chose to use 26-bit with FC and CN, get that and return raw padded hex
    if ($choiceBitCount -eq '1') {
            
        # Get Facility Code
        while ($true) {

            Write-Host "`nEnter the Facility Code (0 to 255): " -NoNewline -ForegroundColor Cyan
            $facilityCode = (Read-Host).Trim()

            if ($facilityCode -notmatch '^\d+$' -or [int]$facilityCode -lt 0 -or [int]$facilityCode -gt 255) {
                Write-Host "Invalid entry. Please enter a number from 0 to 255." -ForegroundColor Red
                Start-Sleep 1
                continue
            }
                
            break

        }

        # Get Card Number
        while ($true) {

            Write-Host "`nEnter the Card Number (0 to 65535): " -NoNewline -ForegroundColor Cyan
            $cardNumber = (Read-Host).Trim()

            if ($cardNumber -notmatch '^\d+$' -or [int]$cardNumber -lt 0 -or [int]$cardNumber -gt 65535) {
                Write-Host "Invalid entry. Please enter a number from 0 to 65535." -ForegroundColor Red
                Start-Sleep 1
                continue
            }
                
            break

        } 

        # Set the credential using the "Get-Standard26BitWiegandHex" function
        $manuallyEnteredRawCredential = Get-Standard26BitWiegandHex -FacilityCode $facilityCode -CardNumber $cardNumber

    } 

    # [2] If user chose 32-bit CSN get that and return raw padded hex
    if ($choiceBitCount -eq '2') {

        # Get 32-bit (4-byte) Hex
        while ($true) {

            Write-Host "`nEnter the 32-bit Hex value (8 Hex characters - 0-9 and A-F): " -NoNewline -ForegroundColor Cyan
            $csnEntry = (Read-Host).Trim()

            if ($csnEntry -notmatch '^[0-9A-Fa-f]{8}$') {
                Write-Host "Invalid entry. Please enter 8 hex characters only (0–9, A–F)." -ForegroundColor Red
                Start-Sleep 1
                continue
            }

            $csnEntry = $csnEntry.ToUpper()
            $manuallyEnteredRawCredential = $csnEntry.PadLeft(32, '0')
            break

        }

    }

    # [3] If user chose custom bit count get that and return raw padded hex
    if ($choiceBitCount -eq '3') {

        # Get raw Hex
        while ($true) {

            Write-Host "`nNote: Shorter values will be padded with leading zeros to 32 characters." -ForegroundColor DarkYellow
            Write-Host "      The selected bit count is handled separately by the software." -ForegroundColor DarkYellow
            Write-Host "      If the value entered exceeds the selected bit count, the higher bits will be ignored." -ForegroundColor DarkYellow

            Write-Host "`nEnter the $($manuallyEnteredBitCount)-bit Hex value (Max 32 Hex characters - 0-9 and A-F): " -NoNewline -ForegroundColor Cyan
            $rawEntry = (Read-Host).Trim()

            if ($rawEntry -notmatch '^[0-9A-Fa-f]{1,32}$') {
                Write-Host "Invalid entry. Please enter 1 to 32 hex characters only (0–9, A–F)." -ForegroundColor Red
                Start-Sleep 1
                continue
            }

            $rawEntry = $rawEntry.ToUpper()
            $manuallyEnteredRawCredential = $rawEntry.PadLeft(32, '0')
            break

        }

    }

    # Return an object containing: 'BitCount' and 'RawCredential'
    return [pscustomobject]@{
        ManuallyEnteredBitCount       = $manuallyEnteredBitCount
        ManuallyEnteredRawCredential  = $manuallyEnteredRawCredential
    }

} 

function Get-RelevantCardholders {
    <#
      This function will get cardholders, their credentials, Cardholder Group(s), and whether they have a PIN.
      
      Logic:
        - If the reader is CardAndPin, only get cardholders who have a card and a PIN.
        - Otherwise, only get cardholders who have a card credential.
    #>
    param(
        [Parameter(Mandatory)] [System.Data.DataTable]$Cardholders,
        [Parameter(Mandatory)] [bool]$ReaderIsCardAndPin
    )

    # Filter rows based on reader mode
    if ($ReaderIsCardAndPin) {
        $rowsToShow = @($Cardholders.Rows | Where-Object { $_['HasPin'] -eq 'True' -and -not [string]::IsNullOrWhiteSpace($_['RawCredential']) })
    }
    else {
        $rowsToShow = @($Cardholders.Rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_['RawCredential']) })
    }

    # If nothing is returned, return to calling function
    if (-not $rowsToShow -or $rowsToShow.Count -eq 0) { Write-Host "`nNo relevant cardholders found..." -NoNewLine -ForegroundColor Red; Start-Sleep 1; return }

    # If there are results, return them to calling function
    return $rowsToShow

} 

function Get-Standard26BitWiegandHex {
    <#
    Purpose:
        - Take the entered Facility Code and Card number and return it as raw hex (with correct parity)
        - This function is called from:
            - Get-RawCredential (any return in this function will go back there)
    #>
    param(
        [Parameter(Mandatory)][ValidateRange(0,255)][int]$FacilityCode,
        [Parameter(Mandatory)][ValidateRange(0,65535)][int]$CardNumber
    )

    $data24 = (($FacilityCode -shl 16) -bor $CardNumber)
    $first12 = ($data24 -shr 12) -band 0xFFF
    $last12  = $data24 -band 0xFFF

    $first12Ones = ([Convert]::ToString($first12, 2).ToCharArray() | Where-Object { $_ -eq '1' }).Count
    $last12Ones  = ([Convert]::ToString($last12, 2).ToCharArray()  | Where-Object { $_ -eq '1' }).Count

    $p1 = if (($first12Ones % 2) -eq 0) { 0 } else { 1 }
    $p2 = if (($last12Ones % 2) -eq 1) { 0 } else { 1 }

    $full26 = (($p1 -shl 25) -bor ($data24 -shl 1) -bor $p2)

    return ('{0:X8}' -f $full26).PadLeft(32, '0')
} 

function Get-StateColor {
    <#
    Purpose:
        - Return the console colour associated with a device state
        - Used by Write-StatusLine to colour status values consistently
    #>
    param(
        [Parameter(Mandatory)] [string]$State
    )

    switch ($State) {
        "Shunted"         { return "Red" }
        "Trouble Unknown" { return "Yellow" }
        "Line Cut"        { return "Yellow" }
        "Short Circuit"   { return "Yellow" }

        "Red"             { return "Red" }
        "Green"           { return "Green" }

        "Online"          { return "Green" }
        "Offline"         { return "Red" }

        "Open"            { return "Green" }
        "Closed"          { return "Green" }

        "Unlocked"        { return "Green" }
        "Locked"          { return "Red" }

        "Unknown"         { return "Yellow" }
        default           { return "White" }
    }
} 

function Get-StateIcon {
     <#
    Purpose:
        - Return a small icon representing a device state
        - Used by Write-StatusLine to make status output easier to read
    #>
    param(
        [Parameter(Mandatory)] [string]$State
    )

    switch ($State) {
        "Shunted"         { return "[X]" }
        "Trouble Unknown" { return "[!]" }
        "Line Cut"        { return "[!]" }
        "Short Circuit"   { return "[!]" }

        "Red"             { return "[R]" }
        "Green"           { return "[G]" }

        "Online"          { return "[+]" }
        "Offline"         { return "[-]" }

        "Open"            { return "[O]" }
        "Closed"          { return "[C]" }

        "Unlocked"        { return "[U]" }
        "Locked"          { return "[L]" }

        "Unknown"         { return "[?]" }
        default           { return "[ ]" }
    }
} 

function Manual-CredentialTypeChoice {
    <#
    Purpose:
        - This allows the user to choose if they want to use an existing cardholder or do a manual entry
    #>
    param(
        [Parameter(Mandatory)] $Door,
        [Parameter(Mandatory)] $Reader
    )

    Write-TitleToConsole -Title "Credential selection"

    while ($true) {

        if ($Reader.Side -eq 'A') { Write-Host "`nHow do you want to interact with the doors ($($Door.Name)) entry reader:" }
        if ($Reader.Side -eq 'B') { Write-Host "`nHow do you want to interact with the doors ($($Door.Name)) exit reader:" }
        
        Write-Host ""
        Write-Host "[1] Use an existing Cardholders credential" 
        Write-Host "[2] Enter a credential manually (e.g., for auto-enrolment or PIN only readers)" 
        Write-Host "`nChoose the type of credential you want to use: " -ForegroundColor Cyan -NoNewline
        $selection = (Read-Host).Trim().ToUpperInvariant()

        if ($selection -notmatch '^[12]$') {
            Write-Host "Invalid selection." -ForegroundColor Red; Start-Sleep 1; continue
        }

        return $selection

    }

} 

function Manual-DoorBehaviour {
    <#
    Purpose:
        - This controls the logic that opens and closes a door (as long as the door has an 'Open Sensor')
    #>
    param(
        [Parameter(Mandatory)] $Door,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array]$Inputs
    )

    # UPDATE PARAMS, DON'T NEED THEM ALL! - THINK IM GOOD
    Write-TitleToConsole -Title "Door Behaviour"

    # Need to refresh the hardware
    $Door = Refresh-CurrentDoor -Door $Door
    $hardware = Get-DoorHardware -Door $Door
    $inputs  = @($hardware | Where-Object RoleType -in @('OpenSensor'))

    if (-not $inputs -or $inputs.Count -eq 0) {
        Write-Host "No door sensors found for this door ($($Door.Name))." -ForegroundColor Yellow
        $waitingToContinue = (Read-Host "`nNote: It is therefore not possible to simulate the door opening and closing, press ENTER to continue").Trim().ToUpperInvariant()
        return $null
    }

    Write-Host "There is $($inputs.Count) door sensor(s) configured for the door ($($Door.Name)).`n" -ForegroundColor Yellow

    # We will use the first (or only) door sensor to use (or not... depending on below choice)
    $inputToUse = $inputs[0].Device

    while ($true) {

        # How would the user like to interact with the door?
        Write-Host "How would you like the door sensor to behave?`n"
        Write-Host "[1] Open and close the door"
        Write-Host "[2] Open the door and leave open"
        Write-Host "[3] Don't open the door"
        Write-Host "`nNote: If you choose option [1] or [2] and there is no access granted, the door will not open!" -ForegroundColor DarkYellow
        Write-Host "      If you want to simulate a door forced event, you need to use 'Operate an Input' from the manual door usage menu." -ForegroundColor DarkYellow
        Write-Host "`nChoose how you would like the door sensor to behave: " -ForegroundColor Cyan -NoNewLine
        $doorBehaviourChoice = (Read-Host).Trim().ToUpperInvariant()

        if ($doorBehaviourChoice -eq "1") {
        
            $doorOpenAndClose    = $true
            $doorOpenAndStayOpen = $false
            $doorNoInputAction   = $false
            break
        
        }

        if ($doorBehaviourChoice -eq "2") {
        
            $doorOpenAndClose    = $false
            $doorOpenAndStayOpen = $true
            $doorNoInputAction   = $false
            break
        
        }

        if ($doorBehaviourChoice -eq "3") {
        
            $doorOpenAndClose    = $false
            $doorOpenAndStayOpen = $false
            $doorNoInputAction   = $true
            break
        
        }

        Write-Host "Incorrect choice. Please try again.`n" -ForegroundColor Red
        Start-Sleep 1
        continue

    }

    # Return users choice and input to use
    return [pscustomobject]@{
        DoorOpenAndClose    = $doorOpenAndClose
        DoorOpenAndStayOpen = $doorOpenAndStayOpen
        DoorNoInputAction   = $doorNoInputAction
        DoorInputToUse      = $inputToUse
    }

} 

function Manual-ExistingCredential {
    <#
    Purpose:
        - Allow the user to select an existing credential 
        - This function is called from:
            - Manual-InvokeDoorRead (any return in this function will go back there)
    #>
    param(
        [Parameter(Mandatory)] $Door,
        [Parameter(Mandatory)] $Reader,
        [Parameter(Mandatory)] $IsReaderCardAndPin
    )

    # Get the cardholders
    $allCardholders = Get-AllCardholders

    # If there are no cardholders, return
    if (-not $allCardholders -or $allCardholders.Rows.Count -eq 0) { 
        Write-Host "No cardholders returned at all... your system doesn't have any?" -ForegroundColor Red
        Start-Sleep 1 
        ClearConsole-Pretty -Message "The console will clear in"
        return 
    }

    # Want to tell user here what's happening... Card & PIN suitable CHs or not
    if ($IsReaderCardAndPin -eq $true)  { 
        Write-Host "`nAs the reader is '" -NoNewLine -ForegroundColor DarkYellow
        Write-Host "Card & PIN" -NoNewLine -ForegroundColor Yellow
        Write-Host "', only searching for Cardholders with a Card and a PIN" -NoNewLine -ForegroundColor DarkYellow; Start-Sleep -Milliseconds 600
        Write-Host "." -NoNewLine -ForegroundColor DarkYellow; Start-Sleep -Milliseconds 600
        Write-Host "." -NoNewLine -ForegroundColor DarkYellow; Start-Sleep -Milliseconds 600
        Write-Host "." -ForegroundColor DarkYellow; Start-Sleep -Milliseconds 600
    }        
    if ($IsReaderCardAndPin -eq $false) { 
        Write-Host "`nAs the reader is '" -NoNewLine -ForegroundColor DarkYellow
        Write-Host "Card only" -NoNewLine -ForegroundColor Yellow
        Write-Host "', only searching for Cardholders with a Card"-NoNewLine -ForegroundColor DarkYellow; Start-Sleep -Milliseconds 600
        Write-Host "." -NoNewLine -ForegroundColor DarkYellow; Start-Sleep -Milliseconds 600
        Write-Host "." -NoNewLine -ForegroundColor DarkYellow; Start-Sleep -Milliseconds 600
        Write-Host "." -ForegroundColor DarkYellow; Start-Sleep -Milliseconds 600
    }

    # Now sort for relevant cardholders e.g., 
    #    - If reader IS card & PIN we need CH's with a card and a PIN
    #    - If reader is NOT card & PIN we need CH's with a card
    $relevantCardholders = Get-RelevantCardholders -Cardholders $allCardholders -ReaderIsCardAndPin $IsReaderCardAndPin

    # If we don't get any returned results... back we go!
    if (-not $relevantCardholders) { Write-Host " No relevant cardholders returned." -ForegroundColor Red; Start-Sleep 2; return }

    # If we DO get relevant returned results... Let's check them out and let the user select a relevant cardholder
    $cardholder = Manual-SelectCardholderToUse -RelevantCardholders $relevantCardholders -ReaderIsCardAndPin $IsReaderCardAndPin

    
    # If the reader IS Card & PIN, we need to get a PIN. 
    if ($IsReaderCardAndPin -eq $true) {
        
        # PIN will never come back empty
        Write-Host ""
        $cardholdersPin = Get-PinCredential -Message "As the reader is card and PIN, please enter a PIN."
        $cardholder | Add-Member -NotePropertyName UseCardAndPin -NotePropertyValue $true -Force
        $cardholder | Add-Member -NotePropertyName UsePinOnly -NotePropertyValue $false -Force
        $cardholder | Add-Member -NotePropertyName PinValue -NotePropertyValue $cardholdersPin -Force

    }

    # If the reader IS NOT Card & PIN AND the cardholder HAS a PIN, the user can choose to use card or PIN (if they choose PIN they have to enter it)
    if ($IsReaderCardAndPin -eq $false -and $cardholder.HasPin -eq $true) {
        
        Write-Host "`nAs the reader is NOT card and PIN, there is no need to enter a PIN unless you want to use only the Cardholders PIN."

        :Choice while ($true) {

            Write-Host "`nDo you want to only use the selected [C] Credential or use only a [P] PIN? " -NoNewLine -ForegroundColor Cyan

            $selection = (Read-Host).Trim().ToUpper() 

            # If credential only they don't want to use a PIN
            if ($selection -eq 'C') {
                $cardholder | Add-Member -NotePropertyName UseCardAndPin -NotePropertyValue $false -Force
                $cardholder | Add-Member -NotePropertyName UsePinOnly -NotePropertyValue $false -Force
                $cardholder | Add-Member -NotePropertyName PinValue -NotePropertyValue "0000" -Force
                break Choice
            }

            # If they want to use a PIN
            if ($selection -eq 'P') {
                Write-Host ""
                $cardholdersPin = Get-PinCredential -Message "As you want to use PIN only, please enter a PIN."
                $cardholder | Add-Member -NotePropertyName UseCardAndPin -NotePropertyValue $false -Force
                $cardholder | Add-Member -NotePropertyName UsePinOnly -NotePropertyValue $true -Force
                $cardholder | Add-Member -NotePropertyName PinValue -NotePropertyValue $cardholdersPin -Force
                break Choice
            }

            Write-Host "Invalid selection. Enter C for credential only or P for PIN only." -ForegroundColor Red
            Start-Sleep 1

        }

    }

    # If the reader IS NOT Card & PIN AND the cardholder does NOT have a PIN, the user can only use the card
    if ($IsReaderCardAndPin -eq $false -and $cardholder.HasPin -eq $false) {
        
        Write-Host "`nAs the reader is NOT card and PIN and the Cardholder doesn't have a PIN, there is no need to enter a PIN."
        Start-Sleep 2

        $cardholder | Add-Member -NotePropertyName UseCardAndPin -NotePropertyValue $false -Force
        $cardholder | Add-Member -NotePropertyName UsePinOnly -NotePropertyValue $false -Force
        $cardholder | Add-Member -NotePropertyName PinValue -NotePropertyValue "0000" -Force
    
    }

    # Return the cardholder to "Manual-InvokeDoorRead"
    return $cardholder

    <#
    # Then I can do things with cardholder... Example useage below

    Write-Host "Name                 : $($cardholder.CardholderName)"
    Write-Host "Cred Name            : $($cardholder.CredentialName)"
    Write-Host "Cred Value (Raw)     : $($cardholder.RawCredential)"
    Write-Host "Bit Count            : $($cardholder.BitCount)"
    Write-Host "Has PIN              : $($cardholder.HasPin)"
    Write-Host "Use Card & PIN       : $($cardholder.UseCardAndPin)"
    Write-Host "Use PIN only         : $($cardholder.UsePinOnly)"
    Write-Host "PIN Value            : $($cardholder.PinValue)" 

    #>

} 

function Manual-ExecuteDoorRead {
    <#
    Purpose:
        - This is the logic for "Manual-InvokeDoorRead" step 7 (door execution)
    #>
    param(
        [Parameter(Mandatory)] $Door,
        [Parameter(Mandatory)] $Cardholder
    )

    <#
        #At this point cardholder contains:
        #----------------------------------

        Write-Host "Cardholder Name      : $($Cardholder.CardholderName)"
        Write-Host "Cred Name            : $($Cardholder.CredentialName)"
        Write-Host "Cred Value (Raw)     : $($Cardholder.RawCredential)"
        Write-Host "Bit Count            : $($Cardholder.BitCount)"
        Write-Host "Has PIN              : $($Cardholder.HasPin)"
        Write-Host "Use Card & PIN       : $($Cardholder.UseCardAndPin)"
        Write-Host "Use PIN only         : $($Cardholder.UsePinOnly)"
        Write-Host "PIN Value            : $($Cardholder.PinValue)"   
        Write-Host "Door Open & Close    : $($Cardholder.DoorOpenAndClose)"  
        Write-Host "Door Open No Close   : $($Cardholder.DoorOpenAndStayOpen)" 
        Write-Host "Door No Open         : $($Cardholder.DoorNoInputAction)"
        Write-Host "Input to Use         : $($Cardholder.DoorInputToUse)"
        Write-Host "Reader to Use        : $($Cardholder.ReaderToUse)"
        Write-Host "Reader Side          : $($Cardholder.ReaderSide)"
        Write-Host "Door Name            : $($Door.Name)"
    #>

    Write-TitleToConsole -Title "Simulated read execution"

    # Let's build some variables to print to console.
    $cardholderName = if ($Cardholder.CardholderName -eq "Unknown - Credential entered manually by user") { "N/A (Raw Entry)" } else { "$($Cardholder.CardholderName)" }
    $credentialType = if ($Cardholder.UseCardAndPin -eq $true) { "Card & PIN" } elseif ($Cardholder.UsePinOnly -eq $true) { "PIN Only" } else { "Card Only" }
    $doorAndReader = if ($Cardholder.ReaderSide -eq "A") { "$($Door.Name) -> Entry Reader" } else { "$($Door.Name) -> Exit Reader" }
    $doorBehaviour = if ($Cardholder.DoorInputToUse -eq $null) { "N/A - No door sensor for this door" } elseif ($Cardholder.DoorInputToUse -ne $null -and $Cardholder.DoorNoInputAction -eq $true) { "No action - door sensor will be left alone" } elseif ($Cardholder.DoorOpenAndStayOpen -eq $true) { "If access is granted - Open only (door will not close)" } else { "If access is granted - Open and close the door" }

    # Print to console what we're going to do
    Write-Host "Credential(s) to use  : " -NoNewline -ForegroundColor Green
    Write-Host "$($cardholderName) -> $($credentialType)"
    Start-Sleep -Milliseconds 600
    Write-Host "Door & reader to use  : " -NoNewline -ForegroundColor Green
    Write-Host "$($doorAndReader)"
    Start-Sleep -Milliseconds 600
    Write-Host "Door sensor behaviour : " -NoNewline -ForegroundColor Green
    Write-Host "$($doorBehaviour)"
    Write-Host ""
    Start-Sleep -Milliseconds 600
    Write-Host "Door/Reader Action    : " -NoNewline -ForegroundColor DarkYellow

  
    # STEP 1 - Simulating the read (3 different ways as below):
    # ---------------------------------------------------------

    # [1/1] Card & PIN
    if ($Cardholder.UseCardAndPin -eq $true) {
        
        # Prepare raw credential (strip left padding based on entered bit count) - there seems to be a bug with Card & PIN only where you have to do this?
        $cleanHex = ($Cardholder.RawCredential -replace '^0x', '' -replace '\s+', '')
        $bytesNeeded = [math]::Ceiling($Cardholder.BitCount / 8)
        $hexCharsNeeded = $bytesNeeded * 2
        $rawToUse = $cleanHex.Substring($cleanHex.Length - $hexCharsNeeded).ToUpper()


        # Card first.
        try {
            
            Write-Host "Swiping the card..." -NoNewline

            Invoke-SWSwipeRaw -Session $sclSession -BitCount $Cardholder.BitCount -Bytes $rawToUse -ReaderPointer $Cardholder.ReaderToUse -ErrorAction Stop | Out-Null

        }
        catch {

            Write-Host "Error message: $($_.Exception.Message)" -ForegroundColor Red
            Start-Sleep 2
            return

        }

        # Then PIN.
        try {
            
            Start-Sleep -Milliseconds 600
            Write-Host " entering the PIN"

            Invoke-SWSwipeWiegand26 -Session $sclSession -Facility 00 -Card $Cardholder.PinValue -ReaderPointer $Cardholder.ReaderToUse -ErrorAction Stop | Out-Null
            
            Start-Sleep -Milliseconds 600

        }
        catch {

            Write-Host "Error message: $($_.Exception.Message)" -ForegroundColor Red
            Start-Sleep 2
            return

        }

    }

    # [2/3] Card Only
    if ($Cardholder.UseCardAndPin -eq $false -and $Cardholder.UsePinOnly -eq $false) {
        
        # Card Only.
        try {
            
            Write-Host "Swiping the card..."

            Invoke-SWSwipeRaw -Session $sclSession -BitCount $Cardholder.BitCount -Bytes $Cardholder.RawCredential -ReaderPointer $Cardholder.ReaderToUse -ErrorAction Stop | Out-Null

            Start-Sleep -Milliseconds 600

        }
        catch {

            Write-Host "Error message: $($_.Exception.Message)" -ForegroundColor Red
            Start-Sleep 2
            return

        }

    }

    # [3/3] PIN Only
    if ($Cardholder.UseCardAndPin -eq $false -and $Cardholder.UsePinOnly -eq $true) {
        
        # PIN Only.
        try {
            
            Write-Host "Entering the PIN..."

            Invoke-SWSwipeWiegand26 -Session $sclSession -Facility 00 -Card $Cardholder.PinValue -ReaderPointer $Cardholder.ReaderToUse -ErrorAction Stop | Out-Null
            
            Start-Sleep -Milliseconds 600

        }
        catch {

            Write-Host "Error message: $($_.Exception.Message)" -ForegroundColor Red
            Start-Sleep 2
            return

        }

    }


    # STEP 2 - Checking the result:
    # ---------------------------------------------------------
    Write-Host "Door/Reader Result    : " -NoNewline -ForegroundColor DarkYellow

    # This will print the result and return true is access granted and false if denied.
    $accessGranted = Get-DoorLastEvent -Door $Door
    Start-Sleep -Milliseconds 600

    # STEP 3 - Operating door input (if applicable):
    # ---------------------------------------------------------
    Write-Host "Door Sensor Action    : " -NoNewline -ForegroundColor DarkYellow

    # [1/4] No door input to use:
    if ($Cardholder.DoorInputToUse -eq $null) { 

        Write-Host "N/A - No door sensor for this door`n"
        Start-Sleep -Milliseconds 600
        return

    }

    # [2/4] User chose not to do anything with the door input:
    if ($Cardholder.DoorInputToUse -ne $null -and $Cardholder.DoorNoInputAction -eq $true) {

        # If access was granted:
        if ($accessGranted -eq $true) {

            Write-Host "I could open the door, but you told me not to`n" -ForegroundColor Yellow
            Start-Sleep -Milliseconds 600
            return

        }

        # If access was denied
        if ($accessGranted -eq $false) {

            Write-Host "I was told not to open the door, just as well... it's still locked`n" -ForegroundColor Yellow
            Start-Sleep -Milliseconds 600
            return

        }

        Write-Host "You found a bug!`n" -ForegroundColor Red
        Start-Sleep -Milliseconds 600
        return
       
    }

    # [3/4] User chose to open and close the door:
    if ($Cardholder.DoorInputToUse -ne $null -and $Cardholder.DoorOpenAndClose -eq $true) {

        # If access was granted:
        if ($accessGranted -eq $true) {

            Write-Host "Opening the door..." -NoNewLine -ForegroundColor Green
            Set-SWInputState -Session $sclSession -InputPointer $Cardholder.DoorInputToUse -State Active | Out-Null
            Start-Sleep 1
            Write-Host " closing the door`n" -ForegroundColor Green
            Set-SWInputState -Session $sclSession -InputPointer $Cardholder.DoorInputToUse -State Inactive | Out-Null
            Start-Sleep -Milliseconds 600
            return

        }

        # If access was denied
        if ($accessGranted -eq $false) {

            Write-Host "I was told to open and close the door... but access was denied`n" -ForegroundColor Red
            Start-Sleep -Milliseconds 600
            return

        }

        Write-Host "You found a bug!`n" -ForegroundColor Red
        Start-Sleep -Milliseconds 600
        return
       
    }

    # [4/4] User chose to open the door and leave it open:
    if ($Cardholder.DoorInputToUse -ne $null -and $Cardholder.DoorOpenAndStayOpen -eq $true) {

        # If access was granted:
        if ($accessGranted -eq $true) {

            Write-Host "Opening the door..." -NoNewLine -ForegroundColor Green
            Set-SWInputState -Session $sclSession -InputPointer $Cardholder.DoorInputToUse -State Active | Out-Null
            Start-Sleep -Milliseconds 600
            Write-Host " leaving it open`n" -ForegroundColor Yellow
            Start-Sleep -Milliseconds 600
            return

        }

        # If access was denied
        if ($accessGranted -eq $false) {

            Write-Host "I was told to open the door and leave it open... but access was denied`n" -ForegroundColor Red
            Start-Sleep -Milliseconds 600
            return

        }

        Write-Host "You found a bug!`n" -ForegroundColor Red
        Start-Sleep -Milliseconds 600
        return
       
    }

    # STEP 4 - bug handling (just in case):
    # ---------------------------------------------------------
    Write-Host "You found a bug!`n" -ForegroundColor Red
    Start-Sleep -Milliseconds 600
    return

}

function Manual-InvokeDoorInput {
    <#
    Purpose:
        - This function allows the user to see the inputs and select one
        - Then the user will see the status of the selected input and alter it
    #>
    param(
        [Parameter(Mandatory)] $Door,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array]$Inputs
    )

    # If no inputs head straight back!
    if (-not $Inputs -or $Inputs.Count -eq 0) {
        Write-Host "How do you plan on using an input if you have no inputs..." -ForegroundColor Yellow
        Start-Sleep 2
        return
    }
   
    while ($true) {
        
        # Need to refresh the hardware
        $Door = Refresh-CurrentDoor -Door $Door
        $hardware = Get-DoorHardware -Door $Door
        $inputs  = @($hardware | Where-Object RoleType -in @('OpenSensor','REX','ManualStation'))

        Write-TitleToConsole -Title "Input operation function for door: $($Door.Name)"

        if (-not $inputs -or $inputs.Count -eq 0) {
            Write-Host "No inputs found for this door." -ForegroundColor Yellow
            return
        }

        Write-Host "There are $($inputs.Count) inputs configured for the door ($($Door.Name)):`n"

        for ($i = 0; $i -lt $inputs.Count; $i++) {
            Write-Host "[$i] $($inputs[$i].RoleType)"
        }

        Write-Host ""
        Write-Host "[C] Clear the console" -ForegroundColor DarkYellow
        Write-Host "[R] Refresh the console" -ForegroundColor DarkYellow
        Write-Host "[Q] Return to door operation menu" -ForegroundColor DarkYellow
        
        Write-Host "`nEnter the number of the input you want to alter: " -ForegroundColor Cyan -NoNewline
        
        $selection = (Read-Host).Trim().ToUpperInvariant()

        if ($selection -eq 'C') {
            ClearConsole-Pretty -Message "The console will clear in"; continue
        }

        if ($selection -eq 'R') {
            cls
            continue
        }

        if ($selection -eq 'Q') {
            ClearConsole-Pretty -Message "The console will clear in"; return
        }

        if ($selection -notmatch '^\d+$' -or [int]$selection -lt 0 -or [int]$selection -ge $inputs.Count) {
            Write-Host "Invalid selection. Please try again.`n" -ForegroundColor Red; Start-Sleep 1; continue
        }

        $selectedInput = $inputs[[int]$selection]
        $selectedInputDevice = $selectedInput.Device
        $selectedInputRole   = $selectedInput.RoleType

        ClearConsole-Pretty -Message "The console will clear in"

        while ($true) {

            # Need to refresh the hardware
            $Door = Refresh-CurrentDoor -Door $Door
            $hardware = Get-DoorHardware -Door $Door
            $input = @(
                $hardware | Where-Object {
                    $_.RoleType -eq $selectedInputRole -and
                    $_.Device   -eq $selectedInputDevice
                } | Select-Object -First 1
            )

            # If input no longer exists break out of this while loop
            if (-not $input) {
                Write-Host "The selected input no longer exists. Returning to the previous menu." -ForegroundColor Red
                Start-Sleep 2
                break
            }

            # Logic
            Write-TitleToConsole -Title "Input commands for the $($input.RoleType) input"

            Write-Host "Note: This information will not refresh live`n" -ForegroundColor Yellow

            Show-DoorStatus -Door $Door -Inputs $input

            Write-Host "`nThere are 5 commands for the $($input.RoleType) input:"
            Write-Host ""
            Write-Host "[1] Open the input"
            Write-Host "[2] Close the input"
            Write-Host "[3] Put the input in an 'unknown trouble' state"
            Write-Host "[4] Put the input in a 'line cut' state"
            Write-Host "[5] Put the input in a 'short circuit' state"
            Write-Host ""
            Write-Host "[C] Clear the console instantly and refresh the input status" -ForegroundColor DarkYellow
            Write-Host "[Q] Return to input selection" -ForegroundColor DarkYellow

            Write-Host "`nChoose what you want to do with this input: " -ForegroundColor Cyan -NoNewline
            $choice = (Read-Host).Trim().ToUpperInvariant()

            if ($choice -eq 'C') {
                cls; continue
            }

            if ($choice -eq 'Q') {
                ClearConsole-Pretty -Message "The console will clear in"; break
            }

            if ($choice -notmatch '^[1-5]$') {
                Write-Host "Invalid selection. Please try again.`n" -ForegroundColor Red; Start-Sleep 1; continue
            }

            switch ($choice) {
                "1" { $state = "Active" }
                "2" { $state = "Inactive" }
                "3" { $state = "Trouble" }
                "4" { $state = "Cut" }
                "5" { $state = "Short" }
            }

            Set-SWInputState -Session $sclSession -InputPointer $input.Device -State $state | Out-Null
            ClearConsole-Pretty -Message "The console will clear in"
        }
    }

    return 

} 

function Manual-InvokeDoorRead {
    <#
    Purpose:
        - This controls the logic that allows a user to swipe a card, use a PIN, 
          or do card and PIN on a door and choose if the door should open or not
    #>
    param(
        [Parameter(Mandatory)] $Door,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array]$Readers,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array]$Inputs,
        [Parameter(Mandatory)] [AllowEmptyCollection()] [array]$Outputs
    )

    # Head back stright away if no readers 
    if (-not $Readers -or $Readers.Count -eq 0) {
        Write-Host "How do you plan on simulating a read if you have no readers..." -ForegroundColor Yellow
        Start-Sleep 2
        return
    }
    
    while ($true) {

        #--------------------------------------------------------------------------------------------------------
        #--> 1. Write title to console
        #--------------------------------------------------------------------------------------------------------
        Write-TitleToConsole -Title "Simulated read function for door: $($Door.Name)"

        #--------------------------------------------------------------------------------------------------------
        #--> 2. Check the door status (if it's unlocked or in maintenance what's the point continuing?)
        #--------------------------------------------------------------------------------------------------------
        $doorState = Get-DoorState -Door $Door

        # We get back if door is unlocked and/or in maintenance mode
        $isDoorLocked            = $doorState.IsDoorLocked
        $isDoorInMaintenanceMode = $doorState.IsDoorInMaintenanceMode

        # Only continue if door is locked and not in maintenance mode/unlocked
        if ($isDoorLocked -eq $false -and $isDoorInMaintenanceMode -eq $true) {
            Write-Host "The selected door is both unlocked and in maintenance mode, returning you to the previous menu.`n" -ForegroundColor Red
            Start-Sleep 1
            ClearConsole-Pretty -Message "The console will clear in"
            return
        }
        if ($isDoorLocked -eq $true -and $isDoorInMaintenanceMode -eq $true) {
            Write-Host "The selected door is in maintenance mode, returning you to the previous menu.`n" -ForegroundColor Red
            Start-Sleep 1
            ClearConsole-Pretty -Message "The console will clear in"
            return
        }
        if ($isDoorLocked -eq $false -and $isDoorInMaintenanceMode -eq $false) {
            Write-Host "The selected door is currently unlocked, returning you to the previous menu.`n" -ForegroundColor Red
            Start-Sleep 1
            ClearConsole-Pretty -Message "The console will clear in"
            return
        }
        if ($isDoorLocked -eq $true -and $isDoorInMaintenanceMode -eq $false) {
            Write-Host "The selected door is neither unlocked or in maintenance mode.`n" -ForegroundColor Green
        }

        #--------------------------------------------------------------------------------------------------------
        #--> 3. Select the reader for use
        #--------------------------------------------------------------------------------------------------------
        $readerSelectionResult = Manual-SelectReader -Door $Door -Readers $Readers

        # If $null is returned push user back to door operation menu
        if ($readerSelectionResult -eq $null) { 
            Write-Host "Returning you to the previous menu."
            Start-Sleep 1
            ClearConsole-Pretty -Message "The console will clear in"
            return 
        }

        # If a reader was selected we get back the reader and a boolean of if it's card only or card & PIN
        $selectedReader     = $readerSelectionResult.SelectedReader
        $isReaderCardAndPin = $readerSelectionResult.IsReaderCardAndPin

        #--------------------------------------------------------------------------------------------------------
        #--> 4. Want to use an existing cardholder or enter a credential manually?
        #--------------------------------------------------------------------------------------------------------
        $credentialTypeChoice = Manual-CredentialTypeChoice -Door $Door -Reader $selectedReader

        # We return 1 or 2 (1 == Existing Cardholder and 2 == Manual credential entry)
        if ($credentialTypeChoice -eq "1") { $cardholderAndCredentials = Manual-ExistingCredential -Door $Door -Reader $selectedReader -IsReaderCardAndPin $isReaderCardAndPin }
        if ($credentialTypeChoice -eq "2") { $cardholderAndCredentials = Manual-ManualEntryCredential -Door $Door -Reader $selectedReader -IsReaderCardAndPin $isReaderCardAndPin }

        # If $cardholderAndCredentials comes back null... handle it here and back we go.
        if ($cardholderAndCredentials -eq $null) { 
            Write-Host "Returning you to the previous menu."
            Start-Sleep 1
            ClearConsole-Pretty -Message "The console will clear in"
            return 
        }

        # Tidy up the console
        ClearConsole-Pretty -Message "The console will clear in"

        #--------------------------------------------------------------------------------------------------------
        #--> 5. Door behaviour
        #--------------------------------------------------------------------------------------------------------
        $doorBehaviour = Manual-DoorBehaviour -Door $Door -Inputs $Inputs
        
        # Need to handle getting back null ... 
        if ($doorBehaviour -eq $null) {
            
            $cardholderAndCredentials | Add-Member -NotePropertyName DoorOpenAndClose -NotePropertyValue $false -Force
            $cardholderAndCredentials | Add-Member -NotePropertyName DoorOpenAndStayOpen -NotePropertyValue $false -Force
            $cardholderAndCredentials | Add-Member -NotePropertyName DoorNoInputAction -NotePropertyValue $true -Force
            $cardholderAndCredentials | Add-Member -NotePropertyName DoorInputToUse -NotePropertyValue $null -Force
            $cardholderAndCredentials | Add-Member -NotePropertyName ReaderToUse -NotePropertyValue $selectedReader.Device -Force
            $cardholderAndCredentials | Add-Member -NotePropertyName ReaderSide -NotePropertyValue $selectedReader.Side -Force

        }

        # ... and getting back some info
        if ($doorBehaviour -ne $null) {
            
            $cardholderAndCredentials | Add-Member -NotePropertyName DoorOpenAndClose -NotePropertyValue $doorBehaviour.DoorOpenAndClose -Force
            $cardholderAndCredentials | Add-Member -NotePropertyName DoorOpenAndStayOpen -NotePropertyValue $doorBehaviour.DoorOpenAndStayOpen -Force
            $cardholderAndCredentials | Add-Member -NotePropertyName DoorNoInputAction -NotePropertyValue $doorBehaviour.DoorNoInputAction -Force
            $cardholderAndCredentials | Add-Member -NotePropertyName DoorInputToUse -NotePropertyValue $doorBehaviour.DoorInputToUse -Force
            $cardholderAndCredentials | Add-Member -NotePropertyName ReaderToUse -NotePropertyValue $selectedReader.Device -Force
            $cardholderAndCredentials | Add-Member -NotePropertyName ReaderSide -NotePropertyValue $selectedReader.Side -Force

        }

        # Tidy up the console
        ClearConsole-Pretty -Message "The console will clear in"

        #--------------------------------------------------------------------------------------------------------
        #--> 6. Final check on the door (hectic isn't it... and well over the top for a POC)
        #--------------------------------------------------------------------------------------------------------

        # STEP 1: Refresh the door - if it's not there it kicks back to main menu
        $door = Refresh-CurrentDoor -Door $Door

        # STEP 2: Check reader and door sensor still exist
        $hardware = Get-DoorHardware -Door $door

        $reader = @(
            $hardware | Where-Object {
                $_.RoleType -eq 'ReaderAuth' -and
                $_.Device   -eq $cardholderAndCredentials.ReaderToUse
            } | Select-Object -First 1
        )

        $input = @(
            $hardware | Where-Object {
                $_.RoleType -eq 'OpenSensor' -and
                $_.Device   -eq $cardholderAndCredentials.DoorInputToUse
            } | Select-Object -First 1
        )

        if (-not $reader) {
            Write-Host "Whoops... looks like the reader you wanted to use no longer exists." -ForegroundColor Red
            Start-Sleep 2
            return
        }

        if (-not $input -and $($cardholderAndCredentials.DoorInputToUse) -ne $null) {
            Write-Host "Whoops... looks like the door sensor you wanted to use no longer exists." -ForegroundColor Red
            Start-Sleep 2
            return
        }

        # STEP 3: Check the door hasnt been put in maintenance mode since the start of this jungle
        $doorState = Get-DoorState -Door $Door

        $isDoorInMaintenanceMode = $doorState.IsDoorInMaintenanceMode

        if ($isDoorInMaintenanceMode -eq $true) {
            Write-Host "The selected door is in maintenance mode, returning you to the previous menu.`n" -ForegroundColor Red
            Start-Sleep 1
            return
        }

        # STEP 4: We don't care about if the cardholder has been deleted, it will just be Access Denied: Unknown Credential

        #--------------------------------------------------------------------------------------------------------
        #--> 7. Door execution (pass in $Door and $cardholderAndCredentials)
        #--------------------------------------------------------------------------------------------------------
        Manual-ExecuteDoorRead -Door $Door -Cardholder $cardholderAndCredentials

        <#
            At this point cardholder contains:
            ----------------------------------

            Write-Host "Cardholder Name      : $($cardholderAndCredentials.CardholderName)"
            Write-Host "Cred Name            : $($cardholderAndCredentials.CredentialName)"
            Write-Host "Cred Value (Raw)     : $($cardholderAndCredentials.RawCredential)"
            Write-Host "Bit Count            : $($cardholderAndCredentials.BitCount)"
            Write-Host "Has PIN              : $($cardholderAndCredentials.HasPin)"
            Write-Host "Use Card & PIN       : $($cardholderAndCredentials.UseCardAndPin)"
            Write-Host "Use PIN only         : $($cardholderAndCredentials.UsePinOnly)"
            Write-Host "PIN Value            : $($cardholderAndCredentials.PinValue)"   
            Write-Host "Door Open & Close    : $($cardholderAndCredentials.DoorOpenAndClose)"  
            Write-Host "Door Open No Close   : $($cardholderAndCredentials.DoorOpenAndStayOpen)" 
            Write-Host "Door No Open         : $($cardholderAndCredentials.DoorNoInputAction)"
            Write-Host "Input to Use         : $($cardholderAndCredentials.DoorInputToUse)"
            Write-Host "Reader to Use        : $($cardholderAndCredentials.ReaderToUse)"
            Write-Host "Reader Side          : $($cardholderAndCredentials.ReaderSide)"
            Write-Host "Door Name            : $($Door.Name)"
        #>

        #--------------------------------------------------------------------------------------------------------
        #--> 8. User can marvel at the wonders of this janky jungle!
        #--------------------------------------------------------------------------------------------------------
        
        # This function ends here... return to "Menu-ManualDoorUsage"
        [void](Read-Host "`nPress ENTER to return to the door usage menu")   
        ClearConsole-Pretty -Message "The console will clear in"
        return

    } 

} 

function Manual-ManualEntryCredential {
    <#
    Purpose:
        - Allow the user to select an existing credential 
        - This function is called from:
            - Manual-InvokeDoorRead (any return in this function will go back there)
    #>
    param(
        [Parameter(Mandatory)] $Door,
        [Parameter(Mandatory)] $Reader,
        [Parameter(Mandatory)] $IsReaderCardAndPin
    )

    # Initialise an object to match what we get if we go down the "Manual-ExistingCredential" function branch of "Manual-InvokeDoorRead"
    $cardholder = [pscustomobject]@{
        CardholderName = "Unknown - Credential entered manually by user" # Fixed
        CredentialName = "Unknown - Credential entered manually by user" # Fixed
        RawCredential  = "00000000000000000000000000000000"              # Need to update based on selection (string)
        BitCount       = "00"                                            # Need to update based on selection (string)
        HasPin         = $false                                          # Need to update based on selection (boolean)
        UseCardAndPin  = $false                                          # Need to update based on selection (boolean)
        UsePinOnly     = $false                                          # Need to update based on selection (boolean)
        PinValue       = "0000"                                          # Need to update based on selection (string)
    }

    # If reader is card and PIN, need both a card and a PIN
    if ($IsReaderCardAndPin -eq $true)  { 
        
        # Want to tell user here what's happening... as is a Card & PIN reader, must enter a card and PIN
        Write-Host "`nAs the reader is '" -NoNewLine -ForegroundColor DarkYellow
        Write-Host "Card & PIN" -NoNewLine -ForegroundColor Yellow
        Write-Host "', you will need to enter both a Card and a PIN" -NoNewLine -ForegroundColor DarkYellow; Start-Sleep -Milliseconds 600
        Write-Host "." -NoNewLine -ForegroundColor DarkYellow; Start-Sleep -Milliseconds 600
        Write-Host "." -NoNewLine -ForegroundColor DarkYellow; Start-Sleep -Milliseconds 600
        Write-Host "." -ForegroundColor DarkYellow; Start-Sleep -Milliseconds 600

        # Get the raw credential first, will return the 'BitCount' and 'RawCredential'
        $rawCredential = Get-RawCredential 

        # Update the cardholder variables 'BitCount' and 'RawCredential' fields
        $cardholder.BitCount      = $rawCredential.ManuallyEnteredBitCount
        $cardholder.RawCredential = $rawCredential.ManuallyEnteredRawCredential

        # Now get the PIN and update 'PinValue' with it
        Write-Host ""
        $cardholder.PinValue = Get-PinCredential -Message "As the reader is 'Card & PIN', you also need to enter a PIN..."

        # Now update 'HasPin', 'UseCardAndPin', and 'UsePinOnly'
        $cardholder.HasPin        = $true
        $cardholder.UseCardAndPin = $true
        $cardholder.UsePinOnly    = $false

    }    
    
    # If reader is NOT card and PIN, need either a card and a PIN (depending on how user wants to interact with the reader) 
    if ($IsReaderCardAndPin -eq $false) { 
        Write-Host "`nAs the reader is '" -NoNewLine -ForegroundColor DarkYellow
        Write-Host "Card only" -NoNewLine -ForegroundColor Yellow
        Write-Host "', you will need to enter either a Card or a PIN"-NoNewLine -ForegroundColor DarkYellow; Start-Sleep -Milliseconds 600
        Write-Host "." -NoNewLine -ForegroundColor DarkYellow; Start-Sleep -Milliseconds 600
        Write-Host "." -NoNewLine -ForegroundColor DarkYellow; Start-Sleep -Milliseconds 600
        Write-Host "." -ForegroundColor DarkYellow; Start-Sleep -Milliseconds 600

        # We won't be using card & PIN seeing as that isn't supported by the reader
        $cardholder.UseCardAndPin = $false

        while ($true) {

            # CODE
            Write-Host "`nDo you want to only use a [C] Credential/Card or use only a [P] PIN?"
            Write-Host ""
            Write-Host "[C] Credential/Card only"
            Write-Host "[P] PIN only"
            Write-Host ""
            Write-Host "Enter your choice: " -NoNewline -ForegroundColor Cyan
            $userChoice = (Read-Host).Trim().ToUpperInvariant()
            Write-Host ""

            # Card only
            if ($userChoice -eq 'C') {

                # Get the raw credential, will return the 'BitCount' and 'RawCredential'
                $rawCredential = Get-RawCredential 

                # Update the cardholder variables 'BitCount' and 'RawCredential' fields
                $cardholder.BitCount      = $rawCredential.ManuallyEnteredBitCount
                $cardholder.RawCredential = $rawCredential.ManuallyEnteredRawCredential

                # Now update 'HasPin', 'UseCardAndPin', and 'UsePinOnly'
                $cardholder.HasPin        = $false
                $cardholder.UseCardAndPin = $false
                $cardholder.UsePinOnly    = $false

                # break from while loop
                break

            }

            # Pin only
            if ($userChoice -eq 'P') {

                # Get the PIN and update 'PinValue' with it
                Write-Host ""
                $cardholder.PinValue = Get-PinCredential -Message "As you will be using PIN only, you need to enter a PIN..."

                # Now update 'HasPin', 'UseCardAndPin', and 'UsePinOnly'
                $cardholder.HasPin        = $true
                $cardholder.UseCardAndPin = $false
                $cardholder.UsePinOnly    = $true

                # break from while loop
                break

            }

            Write-Host "Invalid selection. Enter C for Credential/Card only or P for PIN only." -ForegroundColor Red
            Start-Sleep 1

        }

    }

    # Return the cardholder to "Manual-InvokeDoorRead"
    return $cardholder

    <#
    # Then I can do things with cardholder... Example useage below

    Write-Host "Name                 : $($cardholder.CardholderName)"
    Write-Host "Cred Name            : $($cardholder.CredentialName)"
    Write-Host "Cred Value (Raw)     : $($cardholder.RawCredential)"
    Write-Host "Bit Count            : $($cardholder.BitCount)"
    Write-Host "Has PIN              : $($cardholder.HasPin)"
    Write-Host "Use Card & PIN       : $($cardholder.UseCardAndPin)"
    Write-Host "Use PIN only         : $($cardholder.UsePinOnly)"
    Write-Host "PIN Value            : $($cardholder.PinValue)"

    #>
    
} 

function Manual-SelectCardholderToUse {
    <#
      This function will show the relevant cardholders (as called by "Manual-ExistingCredential") and allow users to select one.
    #>
    param(
        [Parameter(Mandatory)] [array]$RelevantCardholders,
        [Parameter(Mandatory)] [bool]$ReaderIsCardAndPin
    )

    $names = for ($i = 0; $i -lt $RelevantCardholders.Count; $i++) { $RelevantCardholders[$i]['CardholderName'] }

    $uniqueCardholderCount = ($names | Sort-Object -Unique).Count

    if ($ReaderIsCardAndPin) {
        Write-Host "`nThere are $uniqueCardholderCount unique relevant cardholders (that have a Card and a PIN) and $($RelevantCardholders.Count) returned rows:`n" -ForegroundColor Green
    }
    else {
        Write-Host "`nThere are $uniqueCardholderCount unique relevant cardholders (that have a Card) and $($RelevantCardholders.Count) returned rows:`n" -ForegroundColor Green
    }

    Write-Host ("[{0,-2}] {1,-20} | {2,-25} | {3,-35} | {4,-3} | {5,-7}" -f `
        "#", "Cardholder", "Credential Name", "Raw Credential", "Bit", "Has PIN")

    Write-Host ("".PadRight(108, '-'))

    for ($i = 0; $i -lt $RelevantCardholders.Count; $i++) {
        $name = $RelevantCardholders[$i]['CardholderName']
        $credentialName = $RelevantCardholders[$i]['CredentialName']
        $rawCredential = $RelevantCardholders[$i]['RawCredential']
        $bitCount = $RelevantCardholders[$i]['BitCount']
        $hasPin = $RelevantCardholders[$i]['HasPin']

        Write-Host ("[{0,-2}] {1,-20} | {2,-25} | {3,-35} | {4,-3} | {5,-7}" -f `
            $i, $name, $credentialName, $rawCredential, $bitCount, $hasPin) 
    }
   
    while ($true) {

        Write-Host ""
        Write-Host "Enter the row number of the cardholder/credential you want to use: " -NoNewline -ForegroundColor Cyan
        $selection = Read-Host 

        if ($selection -notmatch '^\d+$') {
            Write-Host "Invalid selection. Please enter a valid number." -ForegroundColor Red
            Start-Sleep 1
            continue
        }

        $selection = [int]$selection

        if ($selection -lt 0 -or $selection -ge $RelevantCardholders.Count) {
            Write-Host "Invalid selection. Please enter a valid number." -ForegroundColor Red
            Start-Sleep 1
            continue
        }

        # Return selected cardholder/credential to "Manual-ExistingCredential"
        return $RelevantCardholders[$selection]
    }

} 

function Manual-SelectReader {
    <#
    Purpose:
        - This allows the user to select the reader they want to use for "Manual-InvokeDoorRead"
    #>
    param(
        [Parameter(Mandatory)] $Door,
        [Parameter(Mandatory)] [array]$Readers
    )

    while ($true) {
        
        # Need to refresh the hardware
        $door = Refresh-CurrentDoor -Door $Door
        $hardware = Get-DoorHardware -Door $door
        $readers = @($hardware | Where-Object RoleType -eq 'ReaderAuth' | Sort-Object Side)

        # If no readers tied to the door return $null to "Manual-InvokeDoorRead"
        if (-not $readers -or $readers.Count -eq 0) { Write-Host "No readers found for this door." -ForegroundColor Yellow; Start-Sleep 1; return $null }

        # If there are readers, list them and allow user to select
        Write-Host "There are $($readers.Count) readers configured for the door ($($door.Name)):`n"
        
        Show-DoorHardwareSection -Title "Readers available:" -Items $readers

        Write-Host ""
        Write-Host "      [C] Clear the console" -ForegroundColor DarkYellow
        Write-Host "      [R] Refresh the console" -ForegroundColor DarkYellow
        Write-Host "      [Q] Return to door operation menu" -ForegroundColor DarkYellow
        
        Write-Host "`nEnter the number of the reader you want to use: " -ForegroundColor Cyan -NoNewline
        
        $selection = (Read-Host).Trim().ToUpperInvariant()

        if ($selection -eq 'C') { ClearConsole-Pretty -Message "The console will clear in"; Write-TitleToConsole -Title "Simulated read function for door: $($door.Name)"; continue }
        if ($selection -eq 'R') { cls; Write-TitleToConsole -Title "Simulated read function for door: $($door.Name)"; continue }
        if ($selection -eq 'Q') { return $null } # If user wants to quit return $null to "Manual-InvokeDoorRead"
        if ($selection -notmatch '^\d+$' -or [int]$selection -lt 0 -or [int]$selection -ge $readers.Count) { Write-Host "Invalid selection. Please try again.`n" -ForegroundColor Red; Start-Sleep 1; continue }

        $selectedReader = $readers[[int]$selection]
        # $selectedReader: @{Door=Carpark Entrance; Side=A; RoleType=ReaderAuth; Device=/Devices/Bus/Sim/Port_A/Iface/1/Reader/READER_01; ReaderMode=Normal; ReaderType=Wiegand}
        # Example future use: Invoke-SWSwipeRaw -Session $sclSession -BitCount ?? -Bytes ?? -ReaderPointer $selectedReader.Device

        :ReaderState while ($true) {

            # If someone refreshes this will update the reader
            $door = Refresh-CurrentDoor -Door $Door
            $hardware = Get-DoorHardware -Door $door

            $selectedReader = @(
                $hardware | Where-Object {
                    $_.RoleType -eq 'ReaderAuth' -and
                    $_.Device   -eq $selectedReader.Device
                } | Select-Object -First 1
            )

            if (-not $selectedReader) {
                Write-Host "Whoops... looks like the reader you wanted to use no longer exists." -ForegroundColor Red
                Start-Sleep 2
                return $null
            }

            # Write to console chosen reader of the chosen door and the status of the reader
            Write-Host "`nYou are using " -NoNewLine
            if ($selectedReader.Side -eq "A") { Write-Host "Reader Side $($selectedReader.Side) (In) " -NoNewline -ForegroundColor Green }
            if ($selectedReader.Side -eq "B") { Write-Host "Reader Side $($selectedReader.Side) (Out) " -NoNewline -ForegroundColor Green }
            Write-Host "of door: $($selectedReader.Door)`n"

            Show-DoorStatus -Door $door -Readers $selectedReader

            Write-Host ""
            Write-Host "To avoid a wasted read:" -ForegroundColor Gray
            Write-Host "-----------------------" -ForegroundColor Gray
            Write-Host "   - If the State is " -NoNewLine -ForegroundColor Gray
            Write-Host "Shunted" -NoNewLine -ForegroundColor Red
            Write-Host ": unshunt the reader in Security Desk before continuing..." -ForegroundColor Gray
            Write-Host "   - If the LED is " -NoNewLine -ForegroundColor Gray
            Write-Host "Green" -NoNewLine -ForegroundColor Green
            Write-Host ": the door is currently unlocked/open, check in Security Desk before continuing..." -ForegroundColor Gray
            Write-Host "   - Note: The state of the reader will not update live, use [R] to refresh." -ForegroundColor DarkYellow
            Write-Host ""
            
            $response = (Read-Host "Press R to clear the console and refresh the state, or press ENTER to continue").Trim().ToUpperInvariant()

            if ($response -eq 'R') { cls; continue ReaderState}
            else { ClearConsole-Pretty -Message "The console will clear in"; break ReaderState }

        }

        # initialise a variable for if the chosen reader is CardAndPin
        $isReaderCardAndPin = ($selectedReader.ReaderMode -eq "CardAndPin")

        # Return the selected reader and if reader is card and PIN or not
        return [pscustomobject]@{
            SelectedReader      = $selectedReader
            IsReaderCardAndPin  = $isReaderCardAndPin
        }

    }

} 

function Refresh-CurrentDoor {
    <#
    Purpose:
        - Refresh the selected door so the latest state is used
    #>
    param(
        [Parameter(Mandatory)] $Door
    )

    try {
        $refreshedDoor = Get-Doors -ErrorAction Stop | Where-Object Id -eq $Door.Id | Select-Object -First 1

        if (-not $refreshedDoor) {
            throw "Door not found..."
        }
    }
    catch {
        Write-Host "The selected door no longer exists or could not be refreshed. Returning to the main menu.`n" -ForegroundColor Red
        Start-Sleep 2
        Menu-ManualDoorUsage
        return
    }

    return $refreshedDoor
}  

function Show-DoorHardwareSection {
    <#
    Purpose:
        - Display a list of door hardware in a neat, aligned format
    #>
    param(
        [Parameter(Mandatory)]
        [string]$Title,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [array]$Items
    )

    Write-Host "   $Title" -ForegroundColor DarkYellow

    if (-not $Items -or $Items.Count -eq 0) {
        Write-Host "      [X] None found." -ForegroundColor DarkGray
        return
    }

    for ($i = 0; $i -lt $Items.Count; $i++) {
        $item = $Items[$i]

        $display = switch ($item.RoleType) {
            'ReaderAuth' {
                $direction = switch ($item.Side) {
                    'A' { '(In)' }
                    'B' { '(Out)' }
                    default { '' }
                }

                "Reader Side $($item.Side) $direction"
            }
            'REX'        { "$($item.RoleType) Side $($item.Side)" }
            default      { $item.RoleType }
        }

        Write-Host ("      [{0}] {1,-20} | {2}" -f $i, $display, $item.Device)
    }
} 

function Show-DoorStatus {
    <#
    Purpose:
        - Show the current live status of the selected door's readers, inputs, and outputs
        - Uses helper functions to translate raw device data into readable states
    #>
    param(
        [Parameter(Mandatory)] $Door, [array]$Readers, [array]$Inputs, [array]$Outputs
    )

    # Query Softwire once and reuse the results.
    # This avoids repeatedly calling the API for every device lookup.
    $liveDevices = Get-SWDevices -Session $sclSession

    # ---------------- Readers ----------------
    if ($PSBoundParameters.ContainsKey('Readers')) {

        Write-Host "   Current Reader status:" -ForegroundColor DarkYellow

        if (-not $Readers -or $Readers.Count -eq 0) {
            Write-Host "      None found." -ForegroundColor DarkGray
        }
        else {
            for ($i = 0; $i -lt $Readers.Count; $i++) {
                $reader = $Readers[$i]
                $liveReader = $liveDevices | Where-Object Href -eq $reader.Device

                if (-not $liveReader) {
                    Write-Host ("      [{0}] {1,-20} | Not found" -f $i, "Reader Side $($reader.Side)") -ForegroundColor Red
                    continue
                }

                $readerState = Get-DoorReaderState -ReaderState $liveReader
                $readerLed   = Get-DoorReaderLedColor -ReaderLed $liveReader

                $display = "Reader Side $($reader.Side)"

                Write-StatusLine `
                    -Index $i `
                    -Label $display `
                    -PrimaryLabel "State:" `
                    -PrimaryState $readerState `
                    -SecondaryLabel "LED:" `
                    -SecondaryState $readerLed
            }
        }
    }

    # ---------------- Inputs ----------------
    if ($PSBoundParameters.ContainsKey('Inputs')) { 
        
        Write-Host "   Current Input status:" -ForegroundColor DarkYellow

        if (-not $Inputs -or $Inputs.Count -eq 0) {
            Write-Host "      None found." -ForegroundColor DarkGray
        }
        else {
            for ($i = 0; $i -lt $Inputs.Count; $i++) {
                $input = $Inputs[$i]
                $liveInput = $liveDevices | Where-Object Href -eq $input.Device

                if (-not $liveInput) {
                    Write-Host ("      [{0}] {1,-20} | Not found" -f $i, $input.RoleType) -ForegroundColor Red
                    continue
                }

                $inputState = Get-DoorInputState -InputState $liveInput

                $display = if ($input.RoleType -eq 'REX') {
                    "$($input.RoleType) Side $($input.Side)"
                }
                else {
                    $input.RoleType
                }

                Write-StatusLine `
                    -Index $i `
                    -Label $display `
                    -PrimaryLabel "State:" `
                    -PrimaryState $inputState
            }
        }
    }

    # ---------------- Outputs ----------------
    if ($PSBoundParameters.ContainsKey('Outputs')) {   
        
        Write-Host "   Current Output status:" -ForegroundColor DarkYellow

        if (-not $Outputs -or $Outputs.Count -eq 0) {
            Write-Host "      None found." -ForegroundColor DarkGray
        }
        else {
            for ($i = 0; $i -lt $Outputs.Count; $i++) {
                $output = $Outputs[$i]
                $liveOutput = $liveDevices | Where-Object Href -eq $output.Device

                if (-not $liveOutput) {
                    Write-Host ("      [{0}] {1,-20} | Not found" -f $i, $output.RoleType) -ForegroundColor Red
                    continue
                }

                $outputState = Get-DoorOutputState -OutputState $liveOutput

                Write-StatusLine `
                    -Index $i `
                    -Label $output.RoleType `
                    -PrimaryLabel "State:" `
                    -PrimaryState $outputState
            }
        }
    }
} 

function Simulation-GatherVariables {
    <#
    Purpose:
        - Gather the variables from "Menu-SimulateDoorUsage" step 2 and return them
    #>

    :main while($true) {

        Write-TitleToConsole -Title "Simulated door usage setup"

        # STEP 1 - Get timers to use (low and high)
        :timeBetweenEvents while ($true) {

            Write-Host "[1/4] Which mode you would like to use (delay between events):" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "     [1] EXTREME MODE: Roughly an event per second"
            Write-Host "     [2] RELAXED MODE: An event every 3 to 7 seconds"
            Write-Host "     [3] CUSTOM  MODE: Set your own timer ('n' to 'n' seconds)"
            Write-Host ""
            Write-Host "     Please select the mode you would like to use: " -NoNewline -ForegroundColor Cyan
            $choice = (Read-Host).Trim().ToUpperInvariant()

            # [1] selected
            if ($choice -eq "1") {

                $lowTimer  = 0
                $highTimer = 0
                break timeBetweenEvents

            }

            # [2] selected
            if ($choice -eq "2") {

                $lowTimer  = 3
                $highTimer = 7
                break timeBetweenEvents

            }

            # [3] selected
            if ($choice -eq "3") {
            
                # 1/2 - Get the low timer
                :CustomLowTimerMenu while ($true) {

                    Write-Host ""
                    Write-Host "     Please enter the minimum delay between events (0 to 60): " -NoNewline -ForegroundColor Cyan
                    $lowTimerChoice = (Read-Host).Trim().ToUpperInvariant()

                    if ($lowTimerChoice -notmatch '^\d+$' -or [int]$lowTimerChoice -lt 0 -or [int]$lowTimerChoice -gt 60) {

                        Write-Host "     Invalid choice, please enter a valid number (0 to 60)." -ForegroundColor Red
                        Start-Sleep 1
                        continue CustomLowTimerMenu

                    }

                    $lowTimer  = [int]$lowTimerChoice
                    break CustomLowTimerMenu

                }

                # 2/2 - Get the high timer
                :CustomHighTimerMenu while ($true) {

                    Write-Host ""
                    Write-Host "     Please enter the maximum delay between events (0 to 120): " -NoNewline -ForegroundColor Cyan
                    $highTimerChoice = (Read-Host).Trim().ToUpperInvariant()

                    if ($highTimerChoice -notmatch '^\d+$' -or [int]$highTimerChoice -lt 0 -or [int]$highTimerChoice -gt 120 -or [int]$highTimerChoice -lt [int]$lowTimer) {

                        Write-Host "     Invalid choice, please enter a valid number (0 to 120)." -ForegroundColor Red
                        Write-Host "     Note: This number must be equal to or greater than the minimum delay you entered: $($lowTimer) seconds." -ForegroundColor DarkYellow
                        Start-Sleep 1
                        continue CustomHighTimerMenu

                    }

                    $highTimer = [int]$highTimerChoice
                    break CustomHighTimerMenu

                }

                break timeBetweenEvents

            }

            Write-Host "     Invalid choice, please select a valid number (1 to 3)." -ForegroundColor Red
            Write-Host ""
            Start-Sleep 1
            continue timeBetweenEvents

        }

        # STEP 2 - How many events?
        :howManyEvents while ($true) {

            Write-Host ""
            Write-Host "[2/4] How many events should occur (Min: 1 | Max: 999):" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "     Please enter the number of events you'd like: " -NoNewline -ForegroundColor Cyan
            $howManyEvents = (Read-Host).Trim().ToUpperInvariant()

            if ($howManyEvents -notmatch '^\d+$' -or [int]$howManyEvents -lt 1 -or [int]$howManyEvents -gt 999) {

                Write-Host "     Invalid choice, please enter a valid number (1 to 999)." -ForegroundColor Red
                Start-Sleep 1
                continue howManyEvents

            }

            $numOfEvents = [int]$howManyEvents
            break howManyEvents

        }

        # STEP 3 - Additional events?
        :additionalEvents while ($true) {

            Write-Host ""
            Write-Host "[3/4] Select event profile (affects Door Forced / Door Held frequency):" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "     [1] Normal operation     (Normal = 100% | Forced =  0% | Held =  0%)"
            Write-Host "     [2] Low anomalies        (Normal =  90% | Forced =  5% | Held =  5%)"
            Write-Host "     [3] Typical environment  (Normal =  80% | Forced = 10% | Held = 10%)"
            Write-Host "     [4] Elevated anomalies   (Normal =  70% | Forced = 20% | Held = 10%)"
            Write-Host "     [5] Fault / misuse       (Normal =  50% | Forced = 25% | Held = 25%)"
            Write-Host ""
            Write-Host "     Please select the profile you'd like: " -NoNewline -ForegroundColor Cyan
            $whatEventProfile = (Read-Host).Trim().ToUpperInvariant()

            switch ($whatEventProfile) {

                "1" { $normalEvents = 100; $forcedEvents =  0; $heldEvents =  0; break }
                "2" { $normalEvents =  90; $forcedEvents =  5; $heldEvents =  5; break }
                "3" { $normalEvents =  80; $forcedEvents = 10; $heldEvents = 10; break }
                "4" { $normalEvents =  70; $forcedEvents = 20; $heldEvents = 10; break }
                "5" { $normalEvents =  50; $forcedEvents = 25; $heldEvents = 25; break }
                default { Write-Host "     Invalid choice, please select a valid number (1 to 5)." -ForegroundColor Red; Start-Sleep 1; continue additionalEvents }

            }

            break additionalEvents

        }

        # STEP 4 - Get PIN
        :getPIN while ($true) {

            Write-Host ""
            Write-Host "[4/4] Enter the required PIN:" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "     NOTE: As a reminder all PINs for all Cardholders must match your entered value." -ForegroundColor DarkYellow
            Write-Host ""
            Write-Host "     Please enter the PIN (4-5 digits): " -NoNewline -ForegroundColor Cyan
            $cardholdersPin = (Read-Host).Trim().ToUpperInvariant()

            if ($cardholdersPin -notmatch '^\d{4,5}$') {
                Write-Host "     Invalid PIN. Please enter a 4 or 5 digit PIN." -ForegroundColor Red
                Start-Sleep 1
                continue getPIN
            }

            # Pad to 5 digits if needed (SDK/API needs it to be 5 digits always)
            if ($cardholdersPin.Length -eq 4) {
                $pinToUse = "0$cardholdersPin"
                break getPIN
            }

            $pinToUse = $cardholdersPin
            break getPIN

        }

        # Variable collection summary
        :summary while ($true) {

            Write-Host ""
            Write-Host "Summary" -ForegroundColor Yellow
            Write-Host "-------" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "     1. " -NoNewline
            Write-Host "$($numOfEvents) " -NoNewline -ForegroundColor Green
            Write-Host "events will happen between " -NoNewline
            Write-Host "$($lowTimer) " -NoNewline -ForegroundColor Green
            Write-Host "and " -NoNewLine
            Write-Host "$($highTimer) " -NoNewline -ForegroundColor Green
            Write-Host "seconds."
            Write-Host "     2. Normal events will occur " -NoNewline
            Write-Host "$($normalEvents)% " -NoNewline -ForegroundColor Green
            Write-Host "of the time (Forced: " -NoNewline
            Write-Host "$($forcedEvents)% " -NoNewline -ForegroundColor Green
            Write-Host "& Held: " -NoNewline
            Write-Host "$($heldEvents)% " -NoNewline -ForegroundColor Green
            Write-Host "of the time)."
            Write-Host "     3. Global PIN to use: " -NoNewline
            Write-Host "$($cardholdersPin)" -ForegroundColor Green # Note: Not showing padded variable (no need to confuse the user)
            Write-Host ""
            Write-Host "Enter R to re-enter the variables, or press ENTER to continue: " -NoNewline -ForegroundColor Cyan
            $summaryChoice = (Read-Host).Trim().ToUpperInvariant()

            if ([string]::IsNullOrWhiteSpace($summaryChoice)) {
                break main
            }

            if ($summaryChoice -eq "R") {
                ClearConsole-Pretty -Message "The console will clear in"
                continue main
            }

            Write-Host "Invalid choice. Press ENTER to continue or type R to start again." -ForegroundColor Red
            Start-Sleep 1
            continue summary

        } 

    } 

    ClearConsole-Pretty -Message "The console will clear in"

    return [pscustomobject]@{

        LowTimer     = $lowTimer     # INT
        HighTimer    = $highTimer    # INT
        NumOfEvents  = $numOfEvents  # INT
        NormalEvents = $normalEvents # INT
        ForcedEvents = $forcedEvents # INT
        HeldEvents   = $heldEvents   # INT
        PinToUse     = $pinToUse     # STRING
    }

}

function Simulation-GetSuitableCardholder{
    <# 
      Get a suitable cardholder
    #>
    param(
        [Parameter(Mandatory)]
        $ReaderMode
    )

    # Get the cardholders
    $allCardholders = Get-AllCardholders

    # If there are no cardholders, return
    if (-not $allCardholders -or $allCardholders.Rows.Count -eq 0) {     
        return [pscustomobject]@{
            Success    = $false
            Reason     = "No cardholders returned at all... your system doesn't have any?"
            Cardholder = $null
        }    
    }

    # If there are cardholders then depending on ReaderMode ('CardAndPin' or 'CardOnly'):

    #    - If reader IS card & PIN we need CH's with a card and a PIN
    if ($ReaderMode -eq 'CardAndPin') {
        
        # Filter the cardholders for ones with both cards + PINs
        $cardholderCandidates = @($allCardholders.Rows | Where-Object { $_['HasPin'] -eq 'True' -and -not [string]::IsNullOrWhiteSpace($_['RawCredential']) })

        # If there are no cardholders with Card+PIN, return
        if (-not $cardholderCandidates -or $cardholderCandidates.Count -eq 0) {     
            return [pscustomobject]@{
                Success    = $false
                Reason     = "No cardholders returned with both card and PIN credentials"
                Cardholder = $null
            }    
        }

        # If there are cardholders with both Card + PIN, select one at random
        $selectedCardholder = $cardholderCandidates | Get-Random

        # Return the cardholder
        return [pscustomobject]@{
            Success    = $true
            Reason     = $null
            Cardholder = $selectedCardholder
            CardOnly   = $false
        }

    }

    #    - If reader is NOT card & PIN we need CH's with a card
    if ($ReaderMode -eq 'CardOnly') {

        # Filter the cardholders for ones with just cards
        $cardholderCandidates = @($allCardholders.Rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_['RawCredential']) })

        # If there are no cardholders with just a Card, return
        if (-not $cardholderCandidates -or $cardholderCandidates.Count -eq 0) {     
            return [pscustomobject]@{
                Success    = $false
                Reason     = "No cardholders returned with a card credential"
                Cardholder = $null
            }    
        }

        # If there are cardholders with a Card, select one at random
        $selectedCardholder = $cardholderCandidates | Get-Random

        # Return the cardholder
        return [pscustomobject]@{
            Success    = $true
            Reason     = $null
            Cardholder = $selectedCardholder
            CardOnly   = $true
        }

    }

}

function Simulation-GetSuitableDoor {
    <# 
      Get a suitable door based on the event type as called from STEP 4.D of "Menu-SimulateDoorUsage"
    #>
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Normal', 'Held', 'Forced')]
        [string]$EventType
    )

    # Get suitable doors that aren't unlocked or in maintenance mode
    $baseDoors = Get-SWDoors -Session $sclSession | Where-Object {
        $_.UnlockedForMaintenance -ne $true -and
        $_.IsLocked -ne $false
    }
    
    # If no suitable doors (at all)... return and handle back in "Menu-SimulateDoorUsage"
    if (-not $baseDoors) { 
        return [pscustomobject]@{
            Success = $false
            Reason  = 'No suitable doors found (either none configured or all in maintenance mode/unlocked)'
            Door    = $null
        } 
    }

    # Based on event type filter doors down to a suitable (random) candidate
    switch ($EventType) {

        "Forced" {

            # For a door to be forced it needs to be configured to generate door forced and have a door sensor:
            $forcedCandidates = $baseDoors | Where-Object {
                $_.EnforceDoorForcedOpen -and
                ($_.Roles | ForEach-Object { $_.Type.PSObject.Properties.Name }) -contains 'OpenSensor'
            }

            # If no suitable door for forcing... return and we handle it back in "Menu-SimulateDoorUsage"
            if (-not $forcedCandidates) { 
                return [pscustomobject]@{
                    Success = $false
                    Reason  = 'No suitable doors found (either not configured to allow door forced events and/or no door sensor)'
                    Door    = $null
                } 
            }

            # If there are suitable doors lets select one at random, get the right HW, and return it
            $selectedDoor = $forcedCandidates | Get-Random
            $openSensorDevice = ($selectedDoor.Roles | Where-Object { $_.Type.PSObject.Properties.Name -contains 'OpenSensor' } | Select-Object -First 1).Type.OpenSensor.Device

            return [pscustomobject]@{
                Success    = $true
                Reason     = $null
                Door       = $selectedDoor
                Sensor     = $openSensorDevice
                EventType  = 'Forced'
            }

        }

        "Held"   {

            # For a door to be held it needs to be configured to generate door held and have a door sensor (plus a REX or reader):
            $heldCandidates = $baseDoors | Where-Object {
                $roleTypes = $_.Roles | ForEach-Object { $_.Type.PSObject.Properties.Name }

                ($_.PSObject.Properties.Name -contains 'DoorHeldTime') -and
                ($null -ne $_.DoorHeldTime) -and
                ($roleTypes -contains 'OpenSensor') -and
                (
                    ($roleTypes -contains 'ReaderAuth') -or
                    (
                        ($roleTypes -contains 'REX') -and
                        ($_.AutoUnlockOnRex -eq $true)
                    )
                )
            }

            <# TESTING
            if (-not $heldCandidates) {
                Write-Host "No held candidates found"
            }
            else {
                Write-Host "Held candidates count: $(@($heldCandidates).Count)"
            }
            #>
    
            # If no suitable door for holding open... return and we handle it back in "Menu-SimulateDoorUsage"
            if (-not $heldCandidates) { 
                return [pscustomobject]@{
                    Success = $false
                    Reason  = 'No suitable doors found (either not configured to allow door held events and/or no door sensor)'
                    Door    = $null
                } 
            }

            # If there are suitable doors lets select one at random
            $selectedDoor = $heldCandidates | Get-Random

            # Now let's get the relevant HW...
            $openSensorDevice = ($selectedDoor.Roles | Where-Object { $_.Type.PSObject.Properties.Name -contains 'OpenSensor' } | Select-Object -First 1).Type.OpenSensor.Device
            $reader           = $null
            $rex              = $null
            $readerMode       = $null
            $side             = $null

            # Let's check if we have just readers, just REX's, or both
            $readers = $selectedDoor.Roles | Where-Object { $_.Type.PSObject.Properties.Name -contains 'ReaderAuth' }
            $rexes = if ($selectedDoor.AutoUnlockOnRex -eq $true) { $selectedDoor.Roles | Where-Object { $_.Type.PSObject.Properties.Name -contains 'REX' } } else { @() }

            # If we only have readers we will use a reader
            if ($readers -and -not $rexes) {
                
                $method = 'Reader'
                $selectedReader = $readers | Get-Random
                $reader = $selectedReader.Type.ReaderAuth.HardwareReader
                $side = $selectedReader.Side.PSObject.Properties.Name

                $modeName = $selectedReader.Type.ReaderAuth.ReaderMode.PSObject.Properties.Name

                if ($modeName -contains 'CardAndPin') {
                    $readerMode = 'CardAndPin'
                }
                else {
                    $readerMode = 'CardOnly'
                }

            }
            # If we only have REX's we will use a REX
            elseif ($rexes -and -not $readers) {
                
                $method = 'REX'
                $selectedRex = $rexes | Get-Random
                $rex = $selectedRex.Type.REX.Device
                $side = $selectedRex.Side.PSObject.Properties.Name

            }
            # If we have both we do a weighted choice: Reader 80%, Rex 20%
            else {
                
                $method = if ((Get-Random -Minimum 1 -Maximum 101) -le 80) {
                    'Reader'
                }
                else {
                    'REX'
                }

                if ($method -eq 'Reader') {
                    $selectedReader = $readers | Get-Random
                    $reader = $selectedReader.Type.ReaderAuth.HardwareReader
                    $side = $selectedReader.Side.PSObject.Properties.Name

                    $modeName = $selectedReader.Type.ReaderAuth.ReaderMode.PSObject.Properties.Name

                    if ($modeName -contains 'CardAndPin') {
                        $readerMode = 'CardAndPin'
                    }
                    else {
                        $readerMode = 'CardOnly'
                    }
                }
                else {
                    $selectedRex = $rexes | Get-Random
                    $rex = $selectedRex.Type.REX.Device
                    $side = $selectedRex.Side.PSObject.Properties.Name
                }

            }

            # Switch on $side to make it 'more pretty'.. i dunno I reckon A and B will confuse some people entry and exit is better
            switch ($side) {
                'A'  { $side = 'Entry' }
                'B'  { $side = 'Exit' }
                'NA' { $side = 'None' }
            }

            # Return the door
            return [pscustomobject]@{
                Success    = $true
                Reason     = $null
                Door       = $selectedDoor
                Sensor     = $openSensorDevice
                EventType  = 'Held'
                Method     = $method
                Reader     = $reader
                REX        = $rex
                ReaderMode = $readerMode
                Side       = $side
            }
            
        }

        "Normal" {

            # For a door to be held it needs to be configured with a REX or reader (dont care about door sensor etc.):
            $normalCandidates = $baseDoors | Where-Object {
                $roleTypes = $_.Roles | ForEach-Object { $_.Type.PSObject.Properties.Name }

                ($roleTypes -contains 'ReaderAuth') -or
                    (
                        ($roleTypes -contains 'REX') -and
                        ($_.AutoUnlockOnRex -eq $true)
                    )
            }

            <# TESTING
            if (-not $normalCandidates) {
                Write-Host "No held candidates found"
            }
            else {
                Write-Host "Held candidates count: $(@($normalCandidates).Count)"
            }
            Start-Sleep 1000
            #>
    
            # If no suitable door for holding open... return and we handle it back in "Menu-SimulateDoorUsage"
            if (-not $normalCandidates) { 
                return [pscustomobject]@{
                    Success = $false
                    Reason  = 'No suitable doors found (doors must have at least one reader or REX)'
                    Door    = $null
                } 
            }

            # If there are suitable doors lets select one at random
            $selectedDoor = $normalCandidates | Get-Random

            # Now let's get the relevant HW...
            $openSensorDevice = ($selectedDoor.Roles | Where-Object { $_.Type.PSObject.Properties.Name -contains 'OpenSensor' } | Select-Object -First 1).Type.OpenSensor.Device
            $reader           = $null
            $rex              = $null
            $readerMode       = $null
            $side             = $null

            # Let's check if we have just readers, just REX's, or both
            $readers = $selectedDoor.Roles | Where-Object { $_.Type.PSObject.Properties.Name -contains 'ReaderAuth' }
            $rexes = if ($selectedDoor.AutoUnlockOnRex -eq $true) { $selectedDoor.Roles | Where-Object { $_.Type.PSObject.Properties.Name -contains 'REX' } } else { @() }

            # If we only have readers we will use a reader
            if ($readers -and -not $rexes) {
                
                $method = 'Reader'
                $selectedReader = $readers | Get-Random
                $reader = $selectedReader.Type.ReaderAuth.HardwareReader
                $side = $selectedReader.Side.PSObject.Properties.Name

                $modeName = $selectedReader.Type.ReaderAuth.ReaderMode.PSObject.Properties.Name

                if ($modeName -contains 'CardAndPin') {
                    $readerMode = 'CardAndPin'
                }
                else {
                    $readerMode = 'CardOnly'
                }

            }
            # If we only have REX's we will use a REX
            elseif ($rexes -and -not $readers) {
                
                $method = 'REX'
                $selectedRex = $rexes | Get-Random
                $rex = $selectedRex.Type.REX.Device
                $side = $selectedRex.Side.PSObject.Properties.Name

            }
            # If we have both we do a weighted choice: Reader 80%, Rex 20%
            else {
                
                $method = if ((Get-Random -Minimum 1 -Maximum 101) -le 80) {
                    'Reader'
                }
                else {
                    'REX'
                }

                if ($method -eq 'Reader') {
                    $selectedReader = $readers | Get-Random
                    $reader = $selectedReader.Type.ReaderAuth.HardwareReader
                    $side = $selectedReader.Side.PSObject.Properties.Name

                    $modeName = $selectedReader.Type.ReaderAuth.ReaderMode.PSObject.Properties.Name

                    if ($modeName -contains 'CardAndPin') {
                        $readerMode = 'CardAndPin'
                    }
                    else {
                        $readerMode = 'CardOnly'
                    }
                }
                else {
                    $selectedRex = $rexes | Get-Random
                    $rex = $selectedRex.Type.REX.Device
                    $side = $selectedRex.Side.PSObject.Properties.Name
                }

            }

            # Switch on $side to make it 'more pretty'.. i dunno I reckon A and B will confuse some people entry and exit is better
            switch ($side) {
                'A'  { $side = 'Entry' }
                'B'  { $side = 'Exit' }
                'NA' { $side = 'None' }
            }

            # Return the door
            return [pscustomobject]@{
                Success    = $true
                Reason     = $null
                Door       = $selectedDoor
                Sensor     = $openSensorDevice
                EventType  = 'Normal'
                Method     = $method
                Reader     = $reader
                REX        = $rex
                ReaderMode = $readerMode
                Side       = $side
            }


        }

    }

}

function Simulation-LandingPage {
    <#
    Purpose:
        - Explains what the simulation is for and its limitations
    #>

    Write-Host "                  `"Simulating reality… badly, but usefully`"" -ForegroundColor Green
    Write-Host ""

    # Quick summary (short enough that they’ll actually read it)
    Write-Host "This menu will:" -ForegroundColor Cyan
    Write-Host "---------------" -ForegroundColor Cyan
    Write-Host "  - Allow you to stress test your system, or" -ForegroundColor Gray
    Write-Host "  - Allow you to simulate an active site for use during training or demos." -ForegroundColor Gray
    Write-Host ""

    # Important notes
    Write-Host "Important notes:" -ForegroundColor Yellow
    Write-Host "----------------" -ForegroundColor Yellow
    Write-Host "  - Behaviour and Limitations:" -ForegroundColor DarkYellow
    Write-Host "      1. Only card and/or card+PIN readers are supported (there is no PIN only ability)" -ForegroundColor Gray
    Write-Host "          o NOTE: All Cardholder PIN's MUST be the same (example: every cardholder has a PIN of 1234 for card + PIN doors)" -ForegroundColor DarkGray
    Write-Host "      2. REX (Request to Exit) will only be used on doors configured to 'Unlock on REX'" -ForegroundColor Gray
    Write-Host "          o NOTE: REX events do not identify a cardholder, therefore readers will be selected more than REX's" -ForegroundColor DarkGray
    Write-Host "      3. Doors will only open and close if access is granted and the door unlocks" -ForegroundColor Gray
    Write-Host "          o NOTE: If `"Door Held`" is configured/enabled, doors may be left open instead of closing immediately" -ForegroundColor DarkGray
    Write-Host "          o NOTE: If `"Door Forced`" is configured/enabled, doors may be forced open" -ForegroundColor DarkGray
    Write-Host "      4. Doors that are unlocked and/or in maintenance mode will be skipped for that iteration only" -ForegroundColor Gray
    Write-Host "          o NOTE: If you take the door out of maintenance mode/lock the door it can be used on the next iteration" -ForegroundColor DarkGray
    Write-Host "      5. If you want anti-passback events you must configure anti-passback in your system" -ForegroundColor Gray
    Write-Host "          o NOTE: These events are not guaranteed and depend on system configuration and randomness" -ForegroundColor DarkGray
    Write-Host "      6. Events are randomly generated (door selection, cardholder selection, timing, etc.)" -ForegroundColor Gray
    Write-Host "  - Other Notes:" -ForegroundColor DarkYellow
    Write-Host "      1. You can add/delete/modify doors/door hardware/areas/cardholders/credentials/access rules while the script is running" -ForegroundColor Gray
    Write-Host "          o NOTE: If a door/reader/input becomes invalid during simulation, that attempt will be skipped for that iteration" -ForegroundColor DarkGray
    Write-Host "          o NOTE: The simulation is designed to continue running even if changes occur in the system" -ForegroundColor DarkGray
    Write-Host "      2. Extreme mode may generate a high volume of events and may impact system performance (polite way of saying 'it might melt your VM mate')" -ForegroundColor Gray
    Write-Host "      3. Any comments/complaints/suggestions please email: " -ForegroundColor Gray -NoNewLine
    Write-Host "jsavage@genetec.com" -ForegroundColor Cyan
    Write-Host ""

    # Waiting for enter key to be pressed
    $response = (Read-Host "Press ENTER to continue").Trim().ToUpperInvariant()
    ClearConsole-Pretty -Message "The console will clear in"

}

function Write-StatusLine {
     <#
    Purpose:
        - Print a formatted status line with aligned columns
        - Applies colours and icons based on device state
    #>
    param(
        [Parameter(Mandatory)] [int]$Index,
        [Parameter(Mandatory)] [string]$Label,
        [Parameter(Mandatory)] [string]$PrimaryLabel,
        [Parameter(Mandatory)] [string]$PrimaryState, [string]$SecondaryLabel, [string]$SecondaryState
    )

    $primaryColor = Get-StateColor -State $PrimaryState
    $primaryIcon  = Get-StateIcon  -State $PrimaryState
    $primaryText  = "$primaryIcon $PrimaryState"

    Write-Host ("      [{0}] {1,-20} | {2,-6}" -f $Index, $Label, $PrimaryLabel) -NoNewline
    Write-Host (" {0,-12}" -f $primaryText) -ForegroundColor $primaryColor -NoNewline

    if ($SecondaryLabel -and $SecondaryState) {
        $secondaryColor = Get-StateColor -State $SecondaryState
        $secondaryIcon  = Get-StateIcon  -State $SecondaryState
        $secondaryText  = "$secondaryIcon $SecondaryState"

        Write-Host (" | {0,-4}" -f $SecondaryLabel) -NoNewline
        Write-Host (" {0,-10}" -f $secondaryText) -ForegroundColor $secondaryColor -NoNewline
    }

    Write-Host ""
} 


###############################################################################################################################################################################################################
#                                                    [4/4] USING THE SCRIPT
#                                         --------------------------------------------
###############################################################################################################################################################################################################

#------------------------------------------------------------------------------------------------------
#--> STEP 1 - Show Landing Page (explains what the script does)
#------------------------------------------------------------------------------------------------------
LandingPage

#------------------------------------------------------------------------------------------------------
#--> STEP 2 - Get Softwire Password (will not progress until it's correct)
#------------------------------------------------------------------------------------------------------
$softwireConnection = Connect-Softwire

# Password entered by the user (stored for optional reuse) e.g.; Write-Host "The Softwire password is: $($sclPassword)"
$sclPassword = $softwireConnection.Password 

# Authenticated Softwire session used by cmdlets requiring -Session e.g.; Get-SWDevices -Session $sclSession
$sclSession  = $softwireConnection.Session 

#------------------------------------------------------------------------------------------------------
#--> STEP 3 - Main Menu (choose to interact with doors manually or simulate/stress test)
#------------------------------------------------------------------------------------------------------
while ($true) {

    Write-TitleToConsole -Title "Main menu"

    Write-Host "[1] Interact with doors manually"
    Write-Host "[2] Simulate a busy site"
    Write-Host ""
    Write-Host "[C] Clear the console" -ForegroundColor DarkYellow
    Write-Host "[Q] Quit" -ForegroundColor DarkYellow

    Write-Host "`nChoose how you want to use this tool: " -ForegroundColor Cyan -NoNewline
    $usersChoice = (Read-Host).Trim().ToUpperInvariant()

    if ($usersChoice -eq 'C') { ClearConsole-Pretty -Message "The console will clear in"; continue }

    if ($usersChoice -eq 'Q') { break }

    if ($usersChoice -notmatch '^[1-2]$') { Write-Host "Invalid selection. Please try again." -ForegroundColor Red; Start-Sleep 1; continue }

    switch ($usersChoice) {
        "1" { ClearConsole-Pretty -Message "The console will clear in"; Menu-ManualDoorUsage }
        "2" { ClearConsole-Pretty -Message "The console will clear in"; Menu-SimulateDoorUsage }
    }
}

#--> The end, everyone lived happily ever after... :)