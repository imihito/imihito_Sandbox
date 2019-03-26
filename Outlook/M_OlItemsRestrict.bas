Attribute VB_Name = "M_OlItemsRestrict"
Option Explicit

Sub SampleItemsRestrict()
    Dim targetDate As Date, tomorrowDate As Date
    targetDate = Date
    tomorrowDate = VBA.DateAdd("d", 1, targetDate)
    
    Dim myFolder As Outlook.Folder
    Set myFolder = Session.GetDefaultFolder(olFolderCalendar)
    
    Dim filterString As String
    filterString = "[Start] >= '{TODAY}' AND [Start] <= '{Tomorrow}'"
    filterString = VBA.Replace(filterString, "{TODAY}", Format$(targetDate, "ddddd"))
    filterString = VBA.Replace(filterString, "{Tomorrow}", Format$(tomorrowDate, "ddddd"))
    
    'ファイル > オプション > 予定表 > 稼働時間 外の予定は指定範囲を広げないと検知されない？
    Dim filteredAppointments As Outlook.Items
    Set filteredAppointments = myFolder.Items.Restrict(filterString)
    
    Dim appoint As Outlook.AppointmentItem
    For Each appoint In filteredAppointments
        Debug.Print appoint.Start, appoint.Subject
    Next appoint
    
End Sub
