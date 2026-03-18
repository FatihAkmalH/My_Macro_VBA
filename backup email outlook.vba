Option Explicit

'========================================
' OUTLOOK STARTUP (AUTO)
'========================================
Private Sub Application_Startup()

    On Error Resume Next
    
    If Day(Date) = 6 Then  'auto backup date every month eg: 06 April 2026, the macro will run automatic when app startup
        RunMonthlyBackup
    End If

End Sub


'========================================
' MANUAL BACKUP (ALT + F8)
'========================================
Sub ManualBackup()

    RunMonthlyBackup

End Sub


'========================================
' AUTO DELETE UNDELIVERABLE EMAIL
'========================================
Private Sub Application_NewMailEx(ByVal EntryIDCollection As String)

    On Error Resume Next
    
    Dim ns As Outlook.NameSpace
    Dim arrIDs() As String
    Dim i As Long
    Dim itm As Object
    Dim mai As Outlook.MailItem
    
    Set ns = Application.GetNamespace("MAPI")
    
    arrIDs = Split(EntryIDCollection, ",")
    
    For i = 0 To UBound(arrIDs)
    
        Set itm = ns.GetItemFromID(arrIDs(i))
        
        If Not itm Is Nothing Then
        
            If itm.Class = olMail Then
            
                Set mai = itm
                
                If InStr(1, mai.Subject, "Undeliverable", vbTextCompare) > 0 _
                Or InStr(1, mai.Subject, "Delivery Status Notification", vbTextCompare) > 0 _
                Or InStr(1, mai.Subject, "Delivery has failed", vbTextCompare) > 0 Then
                
                    mai.Move ns.GetDefaultFolder(olFolderDeletedItems)
                
                End If
            
            End If
        
        End If
    
    Next i

End Sub


'========================================
' BACKUP FUNCTION
'========================================
Sub RunMonthlyBackup()

    On Error GoTo SafeExit
    
    Dim ns As Outlook.NameSpace
    Dim inbox As Outlook.Folder
    Dim backupStore As Outlook.Folder
    Dim targetFolder As Outlook.Folder
    
    Dim items As Outlook.Items
    Dim itm As Object
    
    Dim dtStart As Date
    Dim dtEnd As Date
    
    Dim prevMonth As Integer
    Dim prevYear As Integer
    
    Dim movedCount As Long
    Dim i As Long
    
    
    '==============================
    ' COUNT LAST MONTH
    '==============================
    
    prevMonth = Month(Date) - 1
    prevYear = Year(Date)
    
    If prevMonth = 0 Then
        prevMonth = 12
        prevYear = prevYear - 1
    End If
    
    dtStart = DateSerial(prevYear, prevMonth, 1)
    dtEnd = DateSerial(prevYear, prevMonth + 1, 1)
    
    
    '==============================
    ' NAMESPACE
    '==============================
    
    Set ns = Application.GetNamespace("MAPI")
    
    
    '==============================
    ' SOURCE EMAIL (CHANGE ONE OR BOTH)
    '==============================
    
    ' ===== OPTION 1: INBOX =====
    Set inbox = GetInboxByEmail("youremail@yourdomain.com")
    
    ' ===== OPTION 2: SENT ITEMS =====
    ' Gunakan ini jika ingin backup email yang dikirim
    'Set inbox = GetSentItemsByEmail("youremail@yourdomain.com")
    
    
    If inbox Is Nothing Then
        MsgBox "Folder not found!", vbCritical
        Exit Sub
    End If
    
    
    '==============================
    ' PST BACKUP
    '==============================
    
    Set backupStore = ns.Folders("your pst backup file from outlook. eg name: my backup email")
    
    If backupStore Is Nothing Then
        MsgBox "PST File Not Found", vbCritical
        Exit Sub
    End If
    
    
    '==============================
    ' Make Monthly Folder automatic
    '==============================
    
    On Error Resume Next
    
    Set targetFolder = backupStore.Folders(Format(dtStart, "mmm-yyyy"))
    
    If targetFolder Is Nothing Then
        Set targetFolder = backupStore.Folders.Add(Format(dtStart, "mmm-yyyy"))
    End If
    
    On Error GoTo SafeExit
    
    
    '==============================
    ' GET YOUR ITEM
    '==============================
    
    Set items = inbox.Items
    items.Sort "[ReceivedTime]", True
    
    ' NOTE:
    ' If you use Sent Items, change to:
    ' items.Sort "[SentOn]", True
    
    
    movedCount = 0
    
    
    '==============================
    ' LOOP EMAIL (REVERSE LOOP)
    '==============================
    
    For i = items.Count To 1 Step -1
    
        Set itm = items(i)
        
        If itm.Class = olMail Then
        
            ' ===== FOR INBOX =====
            If itm.ReceivedTime >= dtStart And itm.ReceivedTime < dtEnd Then
            
                itm.Move targetFolder
                movedCount = movedCount + 1
                
            End If
            
            ' ===== FOR SENT ITEMS =====
            'If itm.SentOn >= dtStart And itm.SentOn < dtEnd Then
            '
            '    itm.Move targetFolder
            '    movedCount = movedCount + 1
            '
            'End If
        
        End If
    
    Next i
    
    
    MsgBox movedCount & " BACKUP SUCCESS!!! " & targetFolder.Name, vbInformation
    
    
SafeExit:

End Sub


'========================================
' GET INBOX BY EMAIL NAME
'========================================
Function GetInboxByEmail(ByVal accountEmail As String) As Outlook.Folder

    Dim acc As Outlook.Account
    
    For Each acc In Application.Session.Accounts
    
        If LCase(acc.SmtpAddress) = LCase(accountEmail) Then
        
            Set GetInboxByEmail = acc.DeliveryStore.GetDefaultFolder(olFolderInbox)
            Exit Function
        
        End If
    
    Next
    
    Set GetInboxByEmail = Nothing

End Function


'========================================
' GET SENT ITEMS BY EMAIL NAME
'========================================
Function GetSentItemsByEmail(ByVal accountEmail As String) As Outlook.Folder

    Dim acc As Outlook.Account
    
    For Each acc In Application.Session.Accounts
    
        If LCase(acc.SmtpAddress) = LCase(accountEmail) Then
        
            Set GetSentItemsByEmail = acc.DeliveryStore.GetDefaultFolder(olFolderSentMail)
            Exit Function
        
        End If
    
    Next
    
    Set GetSentItemsByEmail = Nothing

End Function

