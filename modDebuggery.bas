Attribute VB_Name = "modDebuggery"
Option Explicit

Private Const LOG_TO_FILE As Boolean = True
Private Const MINIMUM_LOG_LEVEL As Integer = 1
Private Const LOG_BIG_FILE_SIZES As Boolean = False
Private Const LOG_FILE As String = "phlegmoirs_err.log"

Public Sub DebugLog(ByVal sMsg As String, Optional ByVal iLogLevel As Integer = 1)
      If LOG_TO_FILE And iLogLevel >= MINIMUM_LOG_LEVEL Then
            Debug.Print sMsg
            Dim iFile As Integer
            iFile = FreeFile
            Open LOG_FILE For Append As #iFile
            Print #iFile, Now & ": " & sMsg
            Close #iFile
      End If
End Sub

Public Sub PrintComboBoxW(cboBox As ComboBoxW, Optional ByVal iLevel As Integer = 0)
      DebugLog cboBox.Name & " (" & cboBox.SelStart & ", " & cboBox.SelLength & "): """ & cboBox.Text & """", iLevel
      Dim i As Integer
      For i = 0 To cboBox.ListCount - 1
            If i = cboBox.ListIndex Then
                  DebugLog "  * " & cboBox.List(i), iLevel
            Else
                  DebugLog "    " & cboBox.List(i), iLevel
            End If
      Next
End Sub

Public Function WhatsAtThisAddress(ByVal lAddr As Long) As String
      Const HOW_FAR As Integer = 128
      Dim s As String
      s = ""
      Dim iIndex As Integer
      Dim b() As Byte
      ReDim b(0 To HOW_FAR - 1)

      CopyMemory b(0), ByVal lAddr, HOW_FAR
      For iIndex = 0 To HOW_FAR - 1
            s = s & Format(Hex(b(iIndex)), "00") & " "
      Next iIndex
      WhatsAtThisAddress = s & "."
End Function

Public Sub LogBigFileSize(ByRef rtWfd As WIN32_FIND_DATA, ByVal sFileName As String, _
            ByVal sDir As String, ByRef roFso As Object, ByVal oTry1 As Variant)
      
      If LOG_BIG_FILE_SIZES And rtWfd.nFileSizeHigh > 0 Or rtWfd.nFileSizeLow < 0 Then
            Dim oTry2
            oTry2 = roFso.GetFile(sDir + sFileName).Size
            DebugLog ""
            DebugLog "File: " & sFileName, 2
            DebugLog "    size high: " & rtWfd.nFileSizeHigh & "; low: " & rtWfd.nFileSizeLow, 2
            DebugLog "    try1: " & oTry1, 2
            DebugLog "    try2: " & oTry2, 2
      End If
End Sub

Public Function GetMsgName(ByVal lMsg As Long) As String
      Select Case lMsg
            Case (LVM_FIRST + 0): GetMsgName = "LVM_GETBKCOLOR              "
            Case (LVM_FIRST + 1): GetMsgName = "LVM_SETBKCOLOR              "
            Case (LVM_FIRST + 2): GetMsgName = "LVM_GETIMAGELIST            "
            Case (LVM_FIRST + 3): GetMsgName = "LVM_SETIMAGELIST            "
            Case (LVM_FIRST + 4): GetMsgName = "LVM_GETITEMCOUNT            "
            Case (LVM_FIRST + 5): GetMsgName = "LVM_GETITEM                 "
            Case (LVM_FIRST + 76): GetMsgName = "LVM_SETITEM                 "
            Case (LVM_FIRST + 77): GetMsgName = "LVM_INSERTITEM              "
            Case (LVM_FIRST + 8): GetMsgName = "LVM_DELETEITEM              "
            Case (LVM_FIRST + 9): GetMsgName = "LVM_DELETEALLITEMS          "
            Case (LVM_FIRST + 10): GetMsgName = "LVM_GETCALLBACKMASK         "
            Case (LVM_FIRST + 11): GetMsgName = "LVM_SETCALLBACKMASK         "
            Case (LVM_FIRST + 12): GetMsgName = "LVM_GETNEXTITEM             "
            Case (LVM_FIRST + 83): GetMsgName = "LVM_FINDITEM                "
            Case (LVM_FIRST + 14): GetMsgName = "LVM_GETITEMRECT             "
            Case (LVM_FIRST + 16): GetMsgName = "LVM_GETITEMPOSITION         "
            Case (LVM_FIRST + 18): GetMsgName = "LVM_HITTEST                 "
            Case (LVM_FIRST + 19): GetMsgName = "LVM_ENSUREVISIBLE           "
            Case (LVM_FIRST + 20): GetMsgName = "LVM_SCROLL                  "
            Case (LVM_FIRST + 21): GetMsgName = "LVM_REDRAWITEMS             "
            Case (LVM_FIRST + 22): GetMsgName = "LVM_ARRANGE                 "
            Case (LVM_FIRST + 23): GetMsgName = "LVM_EDITLABEL               "
            Case (LVM_FIRST + 24): GetMsgName = "LVM_GETEDITCONTROL          "
            Case (LVM_FIRST + 95): GetMsgName = "LVM_GETCOLUMN               "
            Case (LVM_FIRST + 96): GetMsgName = "LVM_SETCOLUMN               "
            Case (LVM_FIRST + 97): GetMsgName = "LVM_INSERTCOLUMN            "
            Case (LVM_FIRST + 28): GetMsgName = "LVM_DELETECOLUMN            "
            Case (LVM_FIRST + 29): GetMsgName = "LVM_GETCOLUMNWIDTH          "
            Case (LVM_FIRST + 30): GetMsgName = "LVM_SETCOLUMNWIDTH          "
            Case (LVM_FIRST + 31): GetMsgName = "LVM_GETHEADER               "
            Case (LVM_FIRST + 35): GetMsgName = "LVM_GETTEXTCOLOR            "
            Case (LVM_FIRST + 36): GetMsgName = "LVM_SETTEXTCOLOR            "
            Case (LVM_FIRST + 37): GetMsgName = "LVM_GETTEXTBKCOLOR          "
            Case (LVM_FIRST + 38): GetMsgName = "LVM_SETTEXTBKCOLOR          "
            Case (LVM_FIRST + 39): GetMsgName = "LVM_GETTOPINDEX             "
            Case (LVM_FIRST + 40): GetMsgName = "LVM_GETCOUNTPERPAGE         "
            Case (LVM_FIRST + 42): GetMsgName = "LVM_UPDATE                  "
            Case (LVM_FIRST + 43): GetMsgName = "LVM_SETITEMSTATE            "
            Case (LVM_FIRST + 44): GetMsgName = "LVM_GETITEMSTATE            "
            Case (LVM_FIRST + 115): GetMsgName = "LVM_GETITEMTEXT             "
            Case (LVM_FIRST + 116): GetMsgName = "LVM_SETITEMTEXT             "
            Case (LVM_FIRST + 48): GetMsgName = "LVM_SORTITEMS               "
            Case (LVM_FIRST + 50): GetMsgName = "LVM_GETSELECTEDCOUNT        "
            Case (LVM_FIRST + 51): GetMsgName = "LVM_GETITEMSPACING          "
            Case (LVM_FIRST + 53): GetMsgName = "LVM_SETICONSPACING          "
            Case (LVM_FIRST + 54): GetMsgName = "LVM_SETEXTENDEDLISTVIEWSTYLE"
            Case (LVM_FIRST + 55): GetMsgName = "LVM_GETEXTENDEDLISTVIEWSTYLE"
            Case (LVM_FIRST + 56): GetMsgName = "LVM_GETSUBITEMRECT          "
            Case (LVM_FIRST + 60): GetMsgName = "LVM_SETHOTITEM              "
            Case (LVM_FIRST + 61): GetMsgName = "LVM_GETHOTITEM              "
            Case (LVM_FIRST + 62): GetMsgName = "LVM_SETHOTCURSOR            "
            Case (LVM_FIRST + 63): GetMsgName = "LVM_GETHOTCURSOR            "
            Case (LVM_FIRST + 66): GetMsgName = "LVM_GETSELECTIONMARK        "
            Case (LVM_FIRST + 67): GetMsgName = "LVM_SETSELECTIONMARK        "
            Case (LVM_FIRST + 68): GetMsgName = "LVM_SETBKIMAGE              "
            Case (LVM_FIRST + 69): GetMsgName = "LVM_GETBKIMAGE              "
            Case (LVM_FIRST + 81): GetMsgName = "LVM_SORTITEMSEX             "
            Case (LVM_FIRST + 142): GetMsgName = "LVM_SETVIEW                 "
            Case (LVM_FIRST + 143): GetMsgName = "LVM_GETVIEW                 "
            Case (LVM_FIRST + 92): GetMsgName = "LVM_GETGROUPSTATE           "
            Case (LVM_FIRST + 93): GetMsgName = "LVM_GETFOCUSEDGROUP         "
            Case (LVM_FIRST + 98): GetMsgName = "LVM_GETGROUPRECT            "
            Case (LVM_FIRST + 140): GetMsgName = "LVM_SETSELECTEDCOLUMN       "
            Case (LVM_FIRST + 145): GetMsgName = "LVM_INSERTGROUP             "
            Case (LVM_FIRST + 147): GetMsgName = "LVM_SETGROUPINFO            "
            Case (LVM_FIRST + 149): GetMsgName = "LVM_GETGROUPINFO            "
            Case (LVM_FIRST + 150): GetMsgName = "LVM_REMOVEGROUP             "
            Case (LVM_FIRST + 151): GetMsgName = "LVM_MOVEGROUP               "
            Case (LVM_FIRST + 152): GetMsgName = "LVM_GETGROUPCOUNT           "
            Case (LVM_FIRST + 153): GetMsgName = "LVM_GETGROUPINFOBYINDEX     "
            Case (LVM_FIRST + 154): GetMsgName = "LVM_MOVEITEMTOGROUP         "
            Case (LVM_FIRST + 155): GetMsgName = "LVM_SETGROUPMETRICS         "
            Case (LVM_FIRST + 156): GetMsgName = "LVM_GETGROUPMETRICS         "
            Case (LVM_FIRST + 157): GetMsgName = "LVM_ENABLEGROUPVIEW         "
            Case (LVM_FIRST + 158): GetMsgName = "LVM_SORTGROUPS              "
            Case (LVM_FIRST + 159): GetMsgName = "LVM_INSERTGROUPSORTED       "
            Case (LVM_FIRST + 160): GetMsgName = "LVM_REMOVEALLGROUPS         "
            Case (LVM_FIRST + 161): GetMsgName = "LVM_HASGROUP                "
            Case (LVM_FIRST + 175): GetMsgName = "LVM_ISGROUPVIEWENABLED      "
            Case (LVM_FIRST + 205): GetMsgName = "LVM_GETFOOTERRECT           "
            Case (LVM_FIRST + 206): GetMsgName = "LVM_GETFOOTERINFO           "
            Case (LVM_FIRST + 207): GetMsgName = "LVM_GETFOOTERITEMRECT       "
            Case (LVM_FIRST + 208): GetMsgName = "LVM_GETFOOTERITEM           "
            Case Else
                  GetMsgName = "uMsg: H" & CStr(Hex(lMsg))
      End Select
      GetMsgName = Trim(GetMsgName)
End Function

Public Function GetCodeName(ByVal lCode As Long) As String
      Const NM_FIRST As Long = 0
      Select Case lCode
            Case 0: GetCodeName = "NM_FIRST"
            Case (NM_FIRST - 2): GetCodeName = "NM_CLICK"
            Case (NM_FIRST - 3): GetCodeName = "NM_DBLCLK"
            Case (NM_FIRST - 5): GetCodeName = "NM_RCLICK"
            Case (NM_FIRST - 6): GetCodeName = "NM_RDBLCLK"
            Case (NM_FIRST - 12): GetCodeName = "NM_CUSTOMDRAW"
            Case (NM_FIRST - 16): GetCodeName = "NM_RELEASEDCAPTURE"
            
            Case (HDN_FIRST - 0):  GetCodeName = "HDN_ITEMCHANGINGA      "
            Case (HDN_FIRST - 1):  GetCodeName = "HDN_ITEMCHANGEDA       "
            Case (HDN_FIRST - 2):  GetCodeName = "HDN_ITEMCLICKA           "
            Case (HDN_FIRST - 3):  GetCodeName = "HDN_ITEMDBLCLICKA      "
            Case (HDN_FIRST - 5):  GetCodeName = "HDN_DIVIDERDBLCLICKA   "
            Case (HDN_FIRST - 6):  GetCodeName = "HDN_BEGINTRACKA          "
            Case (HDN_FIRST - 7):  GetCodeName = "HDN_ENDTRACKA            "
            Case (HDN_FIRST - 8):  GetCodeName = "HDN_TRACKA               "
            Case (HDN_FIRST - 10): GetCodeName = "HDN_BEGINDRAG      "
            Case (HDN_FIRST - 11): GetCodeName = "HDN_ENDDRAG        "
            Case (HDN_FIRST - 12): GetCodeName = "HDN_FILTERCHANGE   "
            Case (HDN_FIRST - 13): GetCodeName = "HDN_FILTERBTNCLICK "
            Case (HDN_FIRST - 16): GetCodeName = "HDN_ITEMCHECK      "
            Case (HDN_FIRST - 18): GetCodeName = "HDN_DROPDOWN       "
            Case (HDN_FIRST - 20): GetCodeName = "HDN_ITEMCHANGINGW      "
            Case (HDN_FIRST - 21): GetCodeName = "HDN_ITEMCHANGEDW       "
            Case (HDN_FIRST - 22): GetCodeName = "HDN_ITEMCLICKW           "
            Case (HDN_FIRST - 23): GetCodeName = "HDN_ITEMDBLCLICKW      "
            Case (HDN_FIRST - 25): GetCodeName = "HDN_DIVIDERDBLCLICKW   "
            Case (HDN_FIRST - 26): GetCodeName = "HDN_BEGINTRACKW          "
            Case (HDN_FIRST - 27): GetCodeName = "HDN_ENDTRACKW            "
            Case (HDN_FIRST - 28): GetCodeName = "HDN_TRACKW               "
            Case (HDN_FIRST - 99): GetCodeName = "HDN_GETDISPINFO    "
            Case (-520): GetCodeName = "TTN_GETDISPINFO"
            Case Else
                  GetCodeName = "code: " & CStr(lCode)
      End Select
      GetCodeName = Trim(GetCodeName)
End Function
