# Rightfully stolen from https://gist.github.com/asheroto/10d69b6bcb93296c5684a5ca750927aa -- asheroto
# Modified to add a check for xxth generation prepending the intel cpu section ( which is shown on 13th gen cpus )

function Get-IntelProcessorGeneration {
    <#
    .SYNOPSIS
        Returns the generation number of the Intel processor.
    .DESCRIPTION
        Returns the generation number of the Intel processor by parsing the processor name.
        More information: https://www.intel.com/content/www/us/en/support/articles/000032203/processors/intel-core-processors.html
    .EXAMPLE
        PS C:\> Get-IntelProcessorGeneration
    #>

    # Get processor information
    $ProcessorInfo = (Get-CimInstance -ClassName Win32_Processor).Name

    # Check if the processor name contains "xxth Gen"
    if ($ProcessorInfo -match '(\d+)(th|nd|rd|st) Gen') {
        # Extract and return the numeric value preceding "Gen"
        return [int]$matches[1]
    } elseif ($ProcessorInfo -match 'i\d+-\d+') {
        # Existing logic to handle other cases
        $procMatch = $matches[0]
        $genString = $procMatch.Split('-')[1]

        if ($genString.Length -eq 4) {
            # If it's a 4-digit number after the dash, get the first number
            $generationNumber = [int]($genString.Substring(0, 1))
        } elseif ($genString.Length -eq 5) {
            # If it's a 5-digit number after the dash, get the first two numbers
            $generationNumber = [int]($genString.Substring(0, 2))
        } else {
            return -1  # Return -1 if unable to determine generation
        }

        return $generationNumber
    } else {
        return -1  # Return -1 if pattern doesn't match
    }
}
Get-IntelProcessorGeneration