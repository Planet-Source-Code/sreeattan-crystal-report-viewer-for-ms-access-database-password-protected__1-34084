<div align="center">

## Crystal Report Viewer for MS Access Database \(Password Protected\)


</div>

### Description

Crystal Report Viewer.. How to Display a report in Crystal Report Viewer (MS Access Password Protected Database)..
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[sreeattan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sreeattan.md)
**Level**          |Advanced
**User Rating**    |4.0 (16 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sreeattan-crystal-report-viewer-for-ms-access-database-password-protected__1-34084/archive/master.zip)





### Source Code

Private Sub Report(ByVal rptname As String) <br>
Dim crxApplication As New CRAXDRT.Application<br>
Dim Report As CRAXDRT.Report <br>
Dim a As Integer<br>
Set Report = crxApplication.OpenReport(App.Path & rptname, 1)<br>
For a = 1 To Report.Database.Tables.Count<br>
 Report.Database.Tables(a).Location = App.Path & "\database\XXX.mdb"<br>
Report.Database.Tables(a).SetLogOnInfo App.Path & "\database\XXX.mdb", "XXX.mdb", "admin", "pwdhere"<br>
Next<br>
Set frmReport.report = Report<br>
frmReport.Show vbModal, Me <br>
End Sub<br>

