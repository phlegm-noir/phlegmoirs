Attribute VB_Name = "modTypes"
Option Explicit

Enum eIconType
      Directory = 1
      Binary = 2
      Drive = 3
      Text = 4
      Other = 5
      Picture = 6
      IconError = 7
      Bookmark = 8
      Floppy = 9
      Network = 10
      Cdrom = 11
      Rtf = 12
      Music = 13
      Video = 14
End Enum

Enum eViewMode
      Error = -2
      TextView = 1
      PictureView = 2
      PropertiesView = 3
End Enum

Enum eStat
      BrowserStats = 1
      CharStats = 2
      Encoding = 3
      Modified = 4
      SelText = 5
      StatTips = 6
End Enum

Enum eDirection
      Forward = 1
      back = -1
End Enum

Enum eQuery
      Find = 0
End Enum

Enum eTextEncoding
      ASCII = 0
      UNICODE = -1
      Error = -2
End Enum

Enum eIoMode
      ForReading = 1
      ForWriting = 2
      ForAppending = 8
End Enum

Enum eCreate
      Yes = True
      No = False
End Enum

Enum eOverwrite
      Yes = True
      No = False
End Enum

Enum eImageSizingMode
      AlwaysFit = 0
      Default100 = 1
End Enum

Enum eFilerMode
      Files = 0
      Drives = 1
      Bookmarks = 2
      History = 3
End Enum

Public Function GetIconType(sEx As String) As eIconType
      ' This function takes an extension (no dot) and returns a mode
      
      Select Case LCase(sEx)
            Case "bmp", "gif", "jpg", "jpeg", "ico", "cur", "png", "webp", "tif", "tiff"
                  GetIconType = eIconType.Picture
            
            Case "dll", "ocx", "exe", "zip", "msi", "sys", "cab", "7z"
                  GetIconType = eIconType.Binary
            
            Case "mp3", "ogg", "wav", "flac"
                  GetIconType = eIconType.Music
            
            Case "avi", "mpeg", "mp4", "webm", "flv"
                  GetIconType = eIconType.Video
            
            Case "rtf"
                  GetIconType = eIconType.Rtf
            
            Case "txt", "log"
                  GetIconType = eIconType.Text
            
            Case Else
                  GetIconType = eIconType.Other
      End Select
End Function

