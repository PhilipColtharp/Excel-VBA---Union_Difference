'requires Microsoft Scripting Runtime libraries

Function Union_difference(left As String, right As String) As Variant
  Dim Ud() As Variant
  Dim tags As New Scripting.Dictionary
  Dim sortedTags As New Scripting.Dictionary
  Dim thisTag As String
  Dim lastTag As String
  'MsgBox left + right
  
  thisTag = ""
  For i = 1 To Len(left)
    char = Mid(left, i, 1)
    Select Case (char)
      Case ","
        tagsAdd tags, thisTag, 1, 1
        thisTag = ""
      Case Else
        thisTag = thisTag + char
    End Select
  Next
  tagsAdd tags, thisTag, 1, 1
  
  thisTag = ""
  For i = 1 To Len(right)
    char = Mid(right, i, 1)
    Select Case (char)
      Case ","
        tagsAdd tags, thisTag, 3, 2
        thisTag = ""
      Case Else
        thisTag = thisTag + char
    End Select
  Next
  tagsAdd tags, thisTag, 3, 2
  
'sort
  For i = 1 To tags.Count
    If i = 1 Then
      thisTag = tags.Keys(1)
      For Each Key In tags
        If Key < thisTag Then thisTag = Key
      Next
    Else
      lastTag = thisTag
      thisTag = ""
      For Each Key In tags
        If lastTag < Key Then
          If thisTag = "" Then thisTag = Key
          If Key < thisTag Then thisTag = Key
        End If
      Next
    End If
    tagsAdd sortedTags, thisTag, 0, 0
  Next
  
  ReDim Ud(1 To 3)
  For i = 1 To 3
     Ud(i) = ""
  Next
  
  'MsgBox Str(tags.Count)
  For Each Key In sortedTags
    i = (tags(Key) Mod 3) + 1
    If Ud(i) <> "" _
      Then Ud(i) = Ud(i) + ", "
    Ud(i) = Ud(i) + Key
    'MsgBox Key + ":" + Str(tags(Key)) + ":" + Str(i) + ":" + Ud(i)
  Next
   Union_difference = Ud
End Function


Sub tagsAdd(ByRef tags As Dictionary, thisTag As String, OldItem As Integer, newItem As Integer)
  thisTag = Trim(thisTag)
  If thisTag = "" Then Exit Sub
  
  If tags.Exists(thisTag) Then
    tags(thisTag) = OldItem
  Else
    For Each Key In tags
      If Key > thisTag Then Exit For
    Next
    tags.Add Key:=thisTag, item:=newItem
  End If
End Sub
