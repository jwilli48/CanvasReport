function Send-Notification{
  $ButtonContent = @{
    Content = "Open Report"
    Arguments = $ExcelReport
  }
  $Button = New-BTButton @ButtonContent

  $NotificationContent = @{
    Text = "Report for $courseName Generated", "Time taken: $($sw.Elapsed.ToString('hh\:mm\:ss'))"
    Button = $Button
  }
  New-BurntToastNotification @NotificationContent
}
