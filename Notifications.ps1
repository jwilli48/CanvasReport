function Send-Notification {
    <#
  .DESCRIPTION

  Sends a desktop notification with the name of the course, time it took, and a button to open the report.
  #>
    $ButtonContent = @{
        Content   = "Open Report"
        Arguments = $ExcelReport
    }
    $Button = New-BTButton @ButtonContent

    $NotificationContent = @{
        Text   = "Report for $courseName Generated", "Time taken: $($sw.Elapsed.ToString('hh\:mm\:ss'))"
        Button = $Button
    }
    New-BurntToastNotification @NotificationContent
}
