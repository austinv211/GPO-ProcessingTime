<#
NAME: GPOProcessingMonitor.ps1
AUTHOR: Austin Vargason
DESC: Gets GPO Processing Time from Winevents on remote computers
SYNTAX: .\GPOProcessingMonitor.ps1 computers.txt output.csv
NOTE: The first arg is the input text file and the second is the output csv file, these can be full filepaths
CONT... The Second arg is optional and if ommited will just output to the default csv name in the same directory as the session
#>

param(
    [Parameter(Mandatory=$true)]
    [String]$FilePath,
    [Parameter(Mandatory=$false)]
    [String]$OutputCSVPath
)

# function to get the win events for GPO processing on a remote computer
# to speed this up when calling the Get-GPOProcessingTime on each system in the text file provided you can try the experimental
# foreach-parallel or start jobs and join the job data upon completion for "multithreading"
Function Get-GPOProcessingTime() {

    param (
        [Parameter(Mandatory=$true)]
        [String]$ComputerName
    )

    #use try catch to catch timeout to the machine in the list, script may run slow due to system connectrion timeout
    try {
        #get the win event 8004 from Windows Group Policy Operational Logs
        $GPTime = Get-WinEvent -ComputerName $ComputerName -logname Microsoft-Windows-GroupPolicy/Operational | Where-Object {$_.id -eq "8004"} | Select-Object TimeCreated, Message

        # if the GP Time is not equal to null, split the message contents and create a custom psobject
        if ($null -ne $GPTime) {

            #initialize an empty array
            $results = @()

            # loop through the split array and create the objects to be exported in the result array
            ForEach($GPT in $GPTime)
            {
                $tmparray = $GPT.Message.ToString().Split(" ")
                [int]$GPOTime = $tmparray[$tmparray.Count - 2]

                $obj = New-Object -TypeName psobject
                $obj | Add-Member -Name "ComputerName" -Value $ComputerName -MemberType NoteProperty
                $obj | Add-Member -Name "TimeCreated" -Value $GPT.TimeCreated -MemberType NoteProperty
                $obj | Add-Member -Name "ProcessingTime" -Value $GPOTime -MemberType NoteProperty

                $results += $obj
            }

            #return the array
            return $results
        }
        else {
            Write-Host "No Results Found"
        }
    }
    catch {
        Write-Host "Error Gathering Data from Machine: $ComputerName"
    }
}

#function to get the systems to run on via the commandline args and run the Get-GPOProcessing Time function against
Function Start-ProcessingQuery() {
    param(
        [Parameter(Mandatory=$true)]
        [String]$FilePath
    )

    # test the input file path
    if (!(Test-Path -Path $FilePath)) {
        Write-Host "Could Not Find Input File"
        return -1
    }

    # get the computers to query
    $computerList = Get-Content -Path $FilePath

    # create an empty array
    $results = @()

    # loop through and call the function and store the result
    foreach ($computer in $computerList) {
        $results += Get-GPOProcessingTime -ComputerName $computer
    }

    return $results
}

# check the commandline args to determine how to output the data
if ($null -ne $OutputCSVPath) {
    Start-ProcessingQuery -FilePath $FilePath | Export-Csv -NoTypeInformation -Path $OutputCSVPath
}
else {
    Start-ProcessingQuery -FilePath $FilePath | Export-Csv -NoTypeInformation -Path .\GPOOutputData.csv
}


