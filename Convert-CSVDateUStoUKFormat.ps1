

# Takes a CSV, Regex matches each Cell for a date format, casts that to a date time object
# Then Formats that into a UK date format
# Rebuilds a new object to output as a new csv with UK date format dd/MM/yy
# Checks if there are any UK dates in the file and skips the file

$VerbosePreference = 'Continue'

$FolderPath = 'C:\Users\tom\SampleFolder' 

$CSVsInFolderPath = Get-ChildItem -Path "$FolderPath" -Filter *.csv

Foreach ($CSV in $CSVsInFolderPath)
{
  Write-Host "Checking $CSV"
  
  # Reset Process check to null
  $Process = $null
  $CheckFileForUKDates = $null
  $ContainsUKDates = $null
  
  $CSVRecords = Import-csv "$($CSV.FullName)"
  
  $NewCSV = @()
  
  # Get CSV Record Headings
  $CSVHeadings = $CSVRecords[0] | gm | Where-Object {$_.MemberType -eq 'NoteProperty'} | Select-Object Name
  
  ######################################################################################
  # Check for any Date Cells, if none found, skip file. Check the 2nd Row
  Foreach ($testrow in $CSVRecords[2])
  {
    $testrow
    Foreach ($heading in $CSVHeadings)
    {
      # Write-Verbose "Checking $($heading.Name)"
      # Write-Verbose "Value is $($testrow.$($heading.Name))"
      # Regex check to find a date record
      
      IF ($($testrow.$($heading.Name)) -match '([0-9]+)/+([0-9]+)/+([0-9]+)')
      {
        Write-Verbose "Match for $($testrow.$($heading.Name))"
        $CheckFileForUKDates = 'Yes'
        $Process = 'Yes'
        $ColumToCheck = $($heading.Name)
        Write-Verbose "ColumToCheck is $ColumToCheck"
        
      }
      
    }
  } # close test Foreach
  
  #########################################
  # Check all the dates in the file to see if there are any UK format Dates, where xx/yy/zz , xx goes over 12
  IF ($CheckFileForUKDates -eq 'Yes')
  {
    # Search each Row
    Foreach ($row in $CSVRecords)
    {
      
      # Regex check to find a date record
      IF ($row.$($ColumToCheck) -match '([0-9]+)/+([0-9]+)/+([0-9]+)' )
      {
        # Write-Verbose "Match for $($row.$($ColumToCheck))"        
        
        $matches = $row.$($ColumToCheck) | Select-String -Pattern '([0-9]+)/+([0-9]+)/+([0-9]+)'
        
        # Get the match for the first value before the /
        [int]$firstvalue = $Matches.Matches.groups[1].Captures.value
        
        IF ($firstvalue -gt 12)
        {
          
          $ContainsUKDates = 'Yes'
        }
        
      } # close regex check
    }
  } # close if
  
  IF ($ContainsUKDates -eq 'Yes')
    {
    Write-Warning "Do not Process this file, it looks like it has some UK Dates: $CSV"
    }
  
  # IF we should process the file and it contains no UK dates,process the file
  IF ($Process -eq 'Yes' -and $ContainsUKDates -ne 'Yes')
  {
    # Search each Row
    Foreach ($row in $CSVRecords)
    {
      
      # Regex check to find a date record
      IF ($row.$($ColumToCheck) -match '([0-9]+)/+([0-9]+)/+([0-9]+)' )
      {
        Write-Debug "Match for $($row.$($ColumToCheck))"        
        
        $RawValue = $($row.$($ColumToCheck))
        
        $orgCulture = Get-Culture
        [System.Threading.Thread]::CurrentThread.CurrentCulture = New-Object "System.Globalization.CultureInfo" "en-US"
        
        
        If (($RawValue.Length) -gt 8)
        {
          # This works fine for 05/12/2014 00:00 format, but not without the time 
          [datetime]$DateTime = $RawValue
          Write-Debug "Using cast method"
        }
        If (($RawValue.Length) -lt 10)
        {
          # else case is 8 or less characters, so is date without time
          $DateTime = Get-Date $RawValue
          Write-Debug 'using get-date Method'
        }
        
        [System.Threading.Thread]::CurrentThread.CurrentCulture = $orgCulture
        
        $DateInUKFormat = Get-Date $DateTime -Format dd/MM/yyyy
        
        
        # Replace the US date with the UK Date
        $row.$($ColumToCheck) = $DateInUKFormat
        
      } # close If
      
      # Add row to the new CSV
      $NewCSV += $row
      
    } # close Foreach Row
    
    
    $NewCSVName = $CSV.FullName.Replace('.csv','-UKDateFormat.csv')
    
    # Create the New CSV
    Write-Verbose "Creating new file $NewCSVName"
    $NewCSV | Export-Csv -Path $NewCSVName -NoTypeInformation
    
  } # close IF Process 'Yes'
  
  IF ($Process -ne 'Yes')
  {
    Write-Verbose "Skipping File, No dates found: $CSV"
  }
  
  
} # close foreach CSV in Path

