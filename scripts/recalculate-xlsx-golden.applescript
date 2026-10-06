on run argv
  if (count of argv) is not 1 then error "Expected one copied XLSX path" number 8001

  set inputPath to item 1 of argv
  set openedWorkbook to missing value
  set previousSecurity to missing value
  set previousAlerts to missing value
  set previousScreenUpdating to missing value
  set previousCalculation to missing value
  set previousAskToUpdateLinks to missing value
  set previousCalculateBeforeSave to missing value
  set currentStage to "initialize"

  tell application "Microsoft Excel"
    try
      with timeout of 870 seconds
        set my currentStage to "read-settings"
        set my previousSecurity to automation security
        set my previousAlerts to display alerts
        set my previousScreenUpdating to screen updating
        set my previousCalculation to calculation
        set my previousAskToUpdateLinks to ask to update links
        set my previousCalculateBeforeSave to calculate before save

        set my currentStage to "secure-settings"
        set automation security to msoAutomationSecurityForceDisable
        set display alerts to false
        set screen updating to false
        set calculation to calculation manual
        set ask to update links to false
        set calculate before save to false

        set my currentStage to "open-copy"
        set my openedWorkbook to open workbook workbook file name inputPath update links do not update links read only false ignore read only recommended true editable true add to mru false

        set my currentStage to "calculate-full-rebuild"
        calculate full rebuild

        set my currentStage to "save-copy"
        save workbook as my openedWorkbook filename inputPath file format Excel XML file format add to most recently used list false

        set my currentStage to "close-copy"
        close my openedWorkbook saving no
        set my openedWorkbook to missing value

        set my currentStage to "restore-settings"
        if my previousSecurity is not missing value then set automation security to my previousSecurity
        if my previousAlerts is not missing value then set display alerts to my previousAlerts
        if my previousScreenUpdating is not missing value then set screen updating to my previousScreenUpdating
        if my previousCalculation is not missing value then set calculation to my previousCalculation
        if my previousAskToUpdateLinks is not missing value then set ask to update links to my previousAskToUpdateLinks
        if my previousCalculateBeforeSave is not missing value then set calculate before save to my previousCalculateBeforeSave
      end timeout
    on error errorMessage number errorNumber
      set originalErrorMessage to "Excel stage " & (my currentStage) & ": " & (errorMessage as text)
      set originalErrorNumber to errorNumber as integer
      try
        with timeout of 10 seconds
          if my openedWorkbook is not missing value then close my openedWorkbook saving no
        end timeout
      end try
      try
        with timeout of 10 seconds
          if my previousSecurity is not missing value then set automation security to my previousSecurity
          if my previousAlerts is not missing value then set display alerts to my previousAlerts
          if my previousScreenUpdating is not missing value then set screen updating to my previousScreenUpdating
          if my previousCalculation is not missing value then set calculation to my previousCalculation
          if my previousAskToUpdateLinks is not missing value then set ask to update links to my previousAskToUpdateLinks
          if my previousCalculateBeforeSave is not missing value then set calculate before save to my previousCalculateBeforeSave
        end timeout
      end try
      error originalErrorMessage number originalErrorNumber
    end try
  end tell
end run
