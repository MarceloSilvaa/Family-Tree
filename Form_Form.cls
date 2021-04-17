VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Dim curId As Long
Dim newForm As Boolean
Dim pic As String

' ---------------------------
'        Main functions
' ---------------------------

Private Sub Form_Load()
    New_Person_Click
End Sub

Private Sub Form_Close()
    CleanForm
    RestartLists
End Sub

Private Sub New_Person_Click()
    newForm = True
    curId = -1
    CleanForm
    RestartLists
    DisableRelations
End Sub

Private Function LoadData(id As Long)
    Dim i As Long
    Dim count As Long
    Dim con As ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim rst1 As New ADODB.Recordset
    Dim id1 As Long
    Dim n As String
    Dim k As Long
    
    CleanForm
    EnableRelations
    
    curId = id
    newForm = False
    
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient
    
    rst.Open "SELECT * FROM [People] WHERE [Id] = " & id, con, , adLockOptimistic
    Do Until rst.EOF
        BirthName.Value = rst.Fields("BirthName")
        BirthDate.Value = rst.Fields("BirthDate")
        DeathDate.Value = rst.Fields("DeathDate")
        Address.Value = rst.Fields("Address")
        PostalCode.Value = rst.Fields("PostalCode")
        Email.Value = rst.Fields("Email")
        Mobile.Value = rst.Fields("MobileNumber")
        
        For k = 0 To PeopleList.ListCount - 1
            If PeopleList.Column(0, k) = id Then
                PeopleList.Selected(k) = True
                Exit For
            End If
        Next k
        
        PeopleCombo.Value = rst.Fields("BirthName")
        
        If IsNull(rst.Fields("Foto")) Or Len(Trim(rst.Fields("Foto"))) = 0 Then
            If rst.Fields("Gender") Like "F" Then
                pic = "Perfil1.jpg"
                Foto.Picture = Application.CurrentProject.Path & "\Fotos\" & "Perfil1.jpg"
            ElseIf rst.Fields("Gender") Like "M" Then
                pic = "Perfil2.jpg"
                Foto.Picture = Application.CurrentProject.Path & "\Fotos\" & "Perfil2.jpg"
            Else
                pic = ""
                Foto.Picture = ""
            End If
        Else
            Foto.Picture = Application.CurrentProject.Path & "\Fotos\" & rst.Fields("Foto")
        End If
        
        If rst.Fields("Gender") Like "F" Then
            Gender = 1
        ElseIf rst.Fields("Gender") Like "M" Then
            Gender = 2
        End If
        
        rst.MoveNext
    Loop
    rst.Close
    
    rst.Open "People", con, , adLockOptimistic
    Do Until rst.EOF
        id1 = rst.Fields("Id")
        n = rst.Fields("BirthName")
        If rst.Fields("Id") <> id Then
            Siblings.AddItem id1 & ";" & n
            Child.AddItem id1 & ";" & n
            Marriage.AddItem id1 & ";" & n
        End If
        rst.MoveNext
    Loop
    rst.Close
    
    ' Load parents
    rst.Open "SELECT * FROM [IsParent] WHERE [IDPerson2] = " & id, con, , adLockOptimistic
    Do Until rst.EOF
        rst1.Open "SELECT * FROM [People] WHERE [Id] = " & rst.Fields("IDPerson1"), con, , adLockOptimistic
        Do Until rst1.EOF
            If rst1.Fields("Gender") = "F" Then
                MotherImage.Picture = Application.CurrentProject.Path & "\Fotos\" & rst1.Fields("Foto")
                Mother.Caption = rst1.Fields("BirthName")
            ElseIf rst1.Fields("Gender") = "M" Then
                FatherImage.Picture = Application.CurrentProject.Path & "\Fotos\" & rst1.Fields("Foto")
                Father.Caption = rst1.Fields("BirthName")
            End If
            rst1.MoveNext
        Loop
        rst1.Close
        rst.MoveNext
    Loop
    rst.Close
    
    ' Load children
    rst.CursorLocation = adUseClient
    rst.Sort = "IDPerson2 ASC"
    rst.Open "SELECT * FROM [IsParent] WHERE [IDPerson1] = " & id, con, , adLockOptimistic
    Do Until rst.EOF
        rst1.Open "SELECT * FROM [People] WHERE [Id] = " & rst.Fields("IDPerson2"), con, , adLockOptimistic
        Do Until rst1.EOF
            ChildrenList.AddItem rst1.Fields("Id") & ";" & rst1.Fields("BirthName")
            rst1.MoveNext
        Loop
        rst1.Close
        rst.MoveNext
    Loop
    rst.Close
    
    ' Load siblings
    rst.CursorLocation = adUseClient
    rst.Sort = "IDPerson1 ASC"
    rst.Open "SELECT * FROM [IsSibling] WHERE [IDPerson2] = " & id, con, , adLockOptimistic
    Do Until rst.EOF
        rst1.Open "SELECT * FROM [People] WHERE [Id] = " & rst.Fields("IDPerson1"), con, , adLockOptimistic
        Do Until rst1.EOF
            SiblingsList.AddItem rst1.Fields("Id") & ";" & rst1.Fields("BirthName")
            rst1.MoveNext
        Loop
        rst1.Close
        rst.MoveNext
    Loop
    rst.Close
    
    ' Load marriage
    rst.CursorLocation = adUseClient
    rst.Sort = "IDPerson1 ASC"
    rst.Open "SELECT * FROM [IsMarried] WHERE [IDPerson2] = " & id, con, , adLockOptimistic
    Do Until rst.EOF
        rst1.Open "SELECT * FROM [People] WHERE [Id] = " & rst.Fields("IDPerson1"), con, , adLockOptimistic
        Do Until rst1.EOF
            MarriageList.AddItem rst1.Fields("Id") & ";" & rst1.Fields("BirthName")
            rst1.MoveNext
        Loop
        rst1.Close
        rst.MoveNext
    Loop
    rst.Close
End Function

Private Sub DeleteData_Click()
    Dim con As ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim answer As Integer
    
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient
    
    answer = MsgBox("Are you sure that you want to remove this person from all records?", vbExclamation + vbYesNo + vbDefaultButton2, _
    "Delete record")
    
    If answer = vbNo Then
        Exit Sub
    End If
    
    ' Remove from main table
    rst.Open "SELECT * FROM [People] WHERE [Id] = " & curId, con, adOpenStatic, adLockOptimistic
    Do Until rst.EOF
        rst.Delete
        rst.MoveNext
    Loop
    rst.Close
        
    ' Remove from IsParent table
    rst.Open "SELECT * FROM [IsParent] WHERE [IDPerson1] = " & curId, con, adOpenStatic, adLockOptimistic
    Do Until rst.EOF
        rst.Delete
        rst.MoveNext
    Loop
    rst.Close
    
    rst.Open "SELECT * FROM [IsParent] WHERE [IDPerson2] = " & curId, con, adOpenStatic, adLockOptimistic
    Do Until rst.EOF
        rst.Delete
        rst.MoveNext
    Loop
    rst.Close
    
    ' Remove from IsMarried table
    rst.Open "SELECT * FROM [IsMarried] WHERE [IDPerson1] = " & curId, con, adOpenStatic, adLockOptimistic
    Do Until rst.EOF
        rst.Delete
        rst.MoveNext
    Loop
    rst.Close
    
    rst.Open "SELECT * FROM [IsMarried] WHERE [IDPerson2] = " & curId, con, adOpenStatic, adLockOptimistic
    Do Until rst.EOF
        rst.Delete
        rst.MoveNext
    Loop
    rst.Close
    
    ' Remove from IsSibling table
    rst.Open "SELECT * FROM [IsSibling] WHERE [IDPerson1] = " & curId, con, adOpenStatic, adLockOptimistic
    Do Until rst.EOF
        rst.Delete
        rst.MoveNext
    Loop
    rst.Close
    
    rst.Open "SELECT * FROM [IsSibling] WHERE [IDPerson2] = " & curId, con, adOpenStatic, adLockOptimistic
    Do Until rst.EOF
        rst.Delete
        rst.MoveNext
    Loop
    rst.Close
    
    New_Person_Click
End Sub

Private Sub CleanForm()
    ' Main info
    Foto.Picture = ""
    BirthName.Value = ""
    BirthDate.Value = ""
    DeathDate.Value = ""
    Gender.Value = False
    Address.Value = ""
    PostalCode.Value = ""
    Email.Value = ""
    Mobile.Value = ""
    
    ' Parents
    Mother.Caption = "Undefined"
    Father.Caption = "Undefined"
    FatherImage.Picture = ""
    MotherImage.Picture = ""
    
    ' Siblings
    SiblingsList.RowSource = ""
    SiblingsList.Requery
    SiblingsList.Value = ""
    Siblings.RowSource = ""
    Siblings.Requery
    Siblings.Value = ""
    SiblingImage.Picture = ""
    
    ' Marriage
    MarriageList.RowSource = ""
    MarriageList.Requery
    MarriageList.Value = ""
    Marriage.RowSource = ""
    Marriage.Requery
    Marriage.Value = ""
    MarriageImage.Picture = ""
    MarriageDate.Value = ""
    DivorceDate.Value = ""
    
    ' Children
    ChildrenList.RowSource = ""
    ChildrenList.Requery
    ChildrenList.Value = ""
    Child.RowSource = ""
    Child.Requery
    Child.Value = ""
    ChildImage.Picture = ""
End Sub

Private Sub RestartLists()
    PeopleList.RowSource = "All People"
    PeopleList.Requery
    PeopleList.Value = ""
    
    PeopleCombo.RowSource = "All People"
    PeopleCombo.Requery
    PeopleCombo.Value = ""
    
    Dim con As ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim id As Long
    Dim n As String
    
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient
    
    rst.Open "People", con, , adLockOptimistic
    Do Until rst.EOF
        id = rst.Fields("Id")
        n = rst.Fields("BirthName")
        
        Siblings.AddItem id & ";" & n
        Child.AddItem id & ";" & n
        Marriage.AddItem id & ";" & n
        
        rst.MoveNext
    Loop
    rst.Close
End Sub

' ---------------------------
'        Search people
' ---------------------------

Private Sub PeopleList_Click()
    Dim i As Long
    Dim id As Long
    
    i = PeopleList.ListIndex
    
    If i = -1 Then
        Exit Sub
    End If
    
    id = PeopleList.Column(0, i)
    
    LoadData (id)
End Sub

Private Sub PeopleCombo_Click()
    Dim i As Long
    Dim id As Long
    
    i = PeopleCombo.ListIndex
    
    If i = -1 Then
        Exit Sub
    End If
    
    id = PeopleCombo.Column(0, i)
    
    LoadData (id)
End Sub

' ---------------------------
'       Person main info
' ---------------------------

Private Sub SaveMainData_Click()
    If Len(Trim(BirthName.Value)) = 0 Then
        MsgBox "Invalid name"
        Exit Sub
    End If
    
    Dim con As ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient
    
    rst.Open "People", con, adOpenStatic, adLockOptimistic
    If newForm = True Then
        rst.AddNew
    Else
        Do Until rst.EOF
            If rst.Fields("Id") = curId Then
                Exit Do
            End If
            rst.MoveNext
        Loop
    End If
    
    'Fields
        
    rst.Fields("BirthName") = BirthName.Value
        
    If Gender = 1 Then
        rst.Fields("Gender") = "F"
    ElseIf Gender = 2 Then
        rst.Fields("Gender") = "M"
    Else
        rst.Fields("Gender") = ""
    End If
    
    If Len(Trim(BirthDate.Value)) > 0 Then
        rst.Fields("BirthDate") = BirthDate.Value
    Else
        rst.Fields("BirthDate") = ""
    End If
    
    If Len(Trim(DeathDate.Value)) > 0 Then
        rst.Fields("DeathDate") = DeathDate.Value
    Else
        rst.Fields("DeathDate") = ""
    End If
    
    If Len(Trim(pic)) = 0 Then
        rst.Fields("Foto") = ""
    Else
        rst.Fields("Foto") = pic
    End If
    
    If Len(Trim(Address.Value)) > 0 Then
        rst.Fields("Address") = Address.Value
    Else
        rst.Fields("Address") = ""
    End If
    
    If Len(Trim(PostalCode.Value)) > 0 Then
        rst.Fields("PostalCode") = PostalCode.Value
    Else
        rst.Fields("PostalCode") = ""
    End If
    
    If Len(Trim(Email.Value)) > 0 Then
        rst.Fields("Email") = Email.Value
    Else
        rst.Fields("Email") = ""
    End If
    
    If Len(Trim(Mobile.Value)) > 0 Then
        rst.Fields("MobileNumber") = Mobile.Value
    Else
        rst.Fields("MobileNumber") = ""
    End If
    
    rst.Update
    rst.Close
    
    If newForm = True Then
        ' Get the last id (Warning: Id's are not sequencial if records were removed)
        
        rst.Open "People", con, , adLockOptimistic
        rst.MoveLast
        
        curId = rst.Fields("Id")
        newForm = False
        EnableRelations
        
        rst.Close
    End If
    
    RestartLists
    LoadData (curId)
End Sub

Private Sub Foto_Click()
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Show
        If .SelectedItems.count > 0 Then
            strPath = .SelectedItems(1)
        End If
    End With
    
    If Not IsEmpty(strPath) Then
        dirs = Split(strPath, "\Fotos\", , vbTextCompare)
        strFileName = dirs(UBound(dirs))
        
        Foto.Picture = strFileName
        pic = strFileName
    Else
        pic = ""
    End If
End Sub

Private Sub BirthDate_AfterUpdate()
    Dim birth As String
    Dim death As String
    
    ' If death has no value then there is no validation do be made in the birth date
    ' It is possible to store birth date without death date
    
    If IsNull(DeathDate.Value) Then
        DeathDate.Value = ""
        Exit Sub
    End If
    
    death = DeathDate.Value
    
    If Len(Trim(death)) = 0 Then
        Exit Sub
    End If
    
    ' Case where date was removed
    ' Warning: If field is changed to empty (date was removed) then it will have value NULL
    
    If IsNull(BirthDate.Value) Then
        BirthDate.Value = ""
        Exit Sub
    End If
    
    birth = BirthDate.Value
    
    If Len(Trim(birth)) = 0 Then
        Exit Sub
    End If
    
    ' Normal case - Make sure that death date is after birth date
    
    If DateValue(birth) > DateValue(death) Then
        MsgBox "Invalid birth/death date relation.", vbOKOnly, "Invalid Field"
        BirthDate.Value = ""
    End If
End Sub

Private Sub DeathDate_AfterUpdate()
    Dim birth As String
    Dim death As String
    
    ' If birth has no value then there is no validation do be made in the death date
    ' It is possible to store death date without birth date
    
    If IsNull(BirthDate.Value) Then
        BirthDate.Value = ""
        Exit Sub
    End If
    
    birth = BirthDate.Value
    
    If Len(Trim(birth)) = 0 Then
        Exit Sub
    End If
    
    ' Case where date was removed
    ' Warning: If field is changed to empty (date was removed) then it will have value NULL
    
    If IsNull(DeathDate.Value) Then
        DeathDate.Value = ""
        Exit Sub
    End If
    
    death = DeathDate.Value
    
    If Len(Trim(death)) = 0 Then
        Exit Sub
    End If
    
    ' Normal case - Make sure that death date is after birth date
    
    If DateValue(birth) > DateValue(death) Then
        MsgBox "Invalid birth/death date relation.", vbOKOnly, "Invalid Field"
        DeathDate.Value = ""
    End If
End Sub

' ---------------------------
'       Person relations
' ---------------------------

Private Sub EnableRelations()
    ' Parents
    CenterFather.Enabled = True
    CenterMother.Enabled = True
    
    ' Siblings
    SiblingsList.Enabled = True
    Siblings.Enabled = True
    CenterSibling.Enabled = True
    SaveSibling.Enabled = True
    RemoveSibling.Enabled = True
    
    ' Marriage
    MarriageList.Enabled = True
    Marriage.Enabled = True
    CenterMarriage.Enabled = True
    MarriageDate.Enabled = True
    DivorceDate.Enabled = True
    SaveMarriage.Enabled = True
    RemoveMarriage.Enabled = True
    
    ' Children
    ChildrenList.Enabled = True
    Child.Enabled = True
    CenterChild.Enabled = True
    SaveChild.Enabled = True
    RemoveChild.Enabled = True
End Sub

Private Sub DisableRelations()
    ' Parents
    CenterFather.Enabled = False
    CenterMother.Enabled = False
    
    ' Siblings
    SiblingsList.Enabled = False
    Siblings.Enabled = False
    CenterSibling.Enabled = False
    SaveSibling.Enabled = False
    RemoveSibling.Enabled = False
    
    ' Marriage
    MarriageList.Enabled = False
    Marriage.Enabled = False
    CenterMarriage.Enabled = False
    MarriageDate.Enabled = False
    DivorceDate.Enabled = False
    SaveMarriage.Enabled = False
    RemoveMarriage.Enabled = False
    
    ' Children
    ChildrenList.Enabled = False
    Child.Enabled = False
    CenterChild.Enabled = False
    SaveChild.Enabled = False
    RemoveChild.Enabled = False
End Sub

Private Sub CenterFather_Click()
    Dim id As Long
    
    id = FatherId(curId)
    
    If id = -1 Then
        Exit Sub
    End If
    
    LoadData (id)
End Sub

Private Sub CenterMother_Click()
    Dim id As Long
    
    id = MotherId(curId)
    
    If id = -1 Then
        Exit Sub
    End If
    
    LoadData (id)
End Sub

Private Sub Siblings_Click()
    Dim i As Long
    Dim id As Long
    
    i = Siblings.ListIndex
    
    If i = -1 Then
        Exit Sub
    End If
    
    id = Siblings.Column(0, i)
    
    For i = 0 To SiblingsList.ListCount - 1
        If id = SiblingsList.Column(0, i) Then
            SiblingsList.Selected(i) = True
            Exit Sub
        End If
    Next i
    
    SiblingsList.Value = ""
End Sub

Private Sub SiblingsList_Click()
    Dim i As Long
    Dim id As Long
    
    Dim con As ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient

    i = SiblingsList.ListIndex

    If i = -1 Then
        Exit Sub
    End If
    
    id = SiblingsList.Column(0, i)
    
    Siblings.Value = id
    
    rst.Open "People", con, adOpenStatic, adLockOptimistic
    Do Until rst.EOF
        If rst.Fields("Id") = id Then
            Exit Do
        End If
        rst.MoveNext
    Loop
    If Not IsNull(rst.Fields("Foto")) Then
        SiblingImage.Picture = Application.CurrentProject.Path & "\Fotos\" & rst.Fields("Foto")
    Else
        SiblingImage.Picture = ""
    End If
    rst.Close
End Sub

Private Sub CenterSibling_Click()
    Dim id As Long
    
    id = SiblingId()
    
    If id = -1 Then
        Exit Sub
    End If
    
    LoadData (id)
End Sub

Private Sub SaveSibling_Click()
    Dim i As Long
    Dim id As Long
    Dim answer As Integer
    
    i = Siblings.ListIndex
    
    If i = -1 Then
        Exit Sub
    End If
    
    id = Siblings.Column(0, i)
    
    If SiblingRelation(curId, id) = True Then
        MsgBox "This relation was already added."
        Exit Sub
    End If
    
    If MarriageRelation(curId, id) = True Then
        answer = MsgBox("Current person is married to the person you are trying to add as sibling. Do you want to continue?", _
                    vbInformation + vbYesNo + vbDefaultButton2, "Marriage relation")
        If answer = vbNo Then
            Exit Sub
        End If
    End If
    
    If ParentRelation(curId, id) = True Then
        answer = MsgBox("Current person is parent of the person you are trying to add as sibling. Do you want to continue?", _
                    vbInformation + vbYesNo + vbDefaultButton2, "Family relation")
        If answer = vbNo Then
            Exit Sub
        End If
    End If
    
    If ChildRelation(curId, id) = True Then
        answer = MsgBox("Current person is child of the person you are trying to add as sibling. Do you want to continue?", _
                    vbInformation + vbYesNo + vbDefaultButton2, "Family relation")
        If answer = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Add marriage to DB
    Dim con As ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient
    
    rst.Open "IsSibling", con, adOpenStatic, adLockOptimistic
    
    rst.AddNew
    rst.Fields("IDPerson1") = curId
    rst.Fields("IDPerson2") = id
    rst.Update
    rst.Close
    
    rst.Open "IsSibling", con, adOpenStatic, adLockOptimistic
    
    rst.AddNew
    rst.Fields("IDPerson1") = id
    rst.Fields("IDPerson2") = curId
    rst.Update
    rst.Close
    
    LoadData (curId)
End Sub

Private Sub RemoveSibling_Click()
    Dim i As Long
    Dim id As Long
    
    i = Siblings.ListIndex
    
    If i = -1 Then
        Exit Sub
    End If
    
    id = Siblings.Column(0, i)
    
    Dim con As ADODB.Connection
    Dim rst As New ADODB.Recordset
        
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient
    
    rst.Open "SELECT * FROM [IsSibling] WHERE [IDPerson1] = " & curId, con, adOpenStatic, adLockOptimistic
    Do Until rst.EOF
        If id = rst.Fields("IDPerson2") Then
            rst.Delete
            rst.MoveNext
            Exit Do
        End If
        rst.MoveNext
    Loop
    rst.Close
        
    rst.Open "SELECT * FROM [IsSibling] WHERE [IDPerson1] = " & id, con, adOpenStatic, adLockOptimistic
    Do Until rst.EOF
        If curId = rst.Fields("IDPerson2") Then
            rst.Delete
            rst.MoveNext
            Exit Do
        End If
        rst.MoveNext
    Loop
    rst.Close
    
    LoadData (curId)
End Sub

Private Sub Marriage_Click()
    Dim i As Long
    Dim id As Long
    
    i = Marriage.ListIndex
    
    If i = -1 Then
        Exit Sub
    End If
    
    id = Marriage.Column(0, i)
    
    For i = 0 To MarriageList.ListCount - 1
        If id = MarriageList.Column(0, i) Then
            MarriageList.Selected(i) = True
            Exit Sub
        End If
    Next i
    
    MarriageList.Value = ""
End Sub

Private Sub MarriageList_Click()
    Dim i As Long
    Dim id As Long
    
    i = MarriageList.ListIndex
    
    If i = -1 Then
        Exit Sub
    End If
    
    id = MarriageList.Column(0, i)
    
    Marriage.Value = id
    
    Dim con As ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient

    rst.Open "SELECT * FROM [IsMarried] WHERE [IDPerson1] = " & curId, con, , adLockOptimistic
    Do Until rst.EOF
        If id = rst.Fields("IDPerson2") Then
            MarriageDate.Value = rst.Fields("WeddingDate")
            DivorceDate.Value = rst.Fields("DivorceDate")
            Exit Do
        End If
        rst.MoveNext
    Loop
    rst.Close
    
    rst.Open "People", con, adOpenStatic, adLockOptimistic
    Do Until rst.EOF
        If rst.Fields("Id") = id Then
            Exit Do
        End If
        rst.MoveNext
    Loop
    If Not IsNull(rst.Fields("Foto")) Then
        MarriageImage.Picture = Application.CurrentProject.Path & "\Fotos\" & rst.Fields("Foto")
    Else
        MarriageImage.Picture = ""
    End If
    rst.Close
End Sub

Private Sub CenterMarriage_Click()
    Dim id As Long
    
    id = MarriageId()
    
    If id = -1 Then
        Exit Sub
    End If
    
    LoadData (id)
End Sub

Private Sub MarriageDate_AfterUpdate()
    Dim Marriage As String
    Dim divorce As String
    
    ' If divorce has no value then there is no validation do be made in the marriage date
    ' It is possible to store marriage date without divorce date
    
    If IsNull(DivorceDate.Value) Then
        DivorceDate.Value = ""
        Exit Sub
    End If
    
    divorce = DivorceDate.Value
    
    If Len(Trim(divorce)) = 0 Then
        Exit Sub
    End If
    
    ' Case where date was removed
    ' Warning: If field is changed to empty (date was removed) then it will have value NULL
    
    If IsNull(MarriageDate.Value) Then
        MarriageDate.Value = ""
        Exit Sub
    End If
    
    Marriage = MarriageDate.Value
    
    If Len(Trim(Marriage)) = 0 Then
        Exit Sub
    End If
    
    ' Normal case - Make sure that marriage date is after divorce date
    
    If DateValue(Marriage) > DateValue(divorce) Then
        MsgBox "Invalid marriage/divorce date relation.", vbOKOnly, "Invalid Field"
        MarriageDate.Value = ""
    End If
End Sub

Private Sub DivorceDate_AfterUpdate()
    Dim Marriage As String
    Dim divorce As String
    
    ' If marriage has no value then there is no validation do be made in the divorce date
    ' It is possible to store divorce date without marriage date
    
    If IsNull(MarriageDate.Value) Then
        MarriageDate.Value = ""
        Exit Sub
    End If
    
    Marriage = MarriageDate.Value
    
    If Len(Trim(Marriage)) = 0 Then
        Exit Sub
    End If
    
    ' Case where date was removed
    ' Warning: If field is changed to empty (date was removed) then it will have value NULL
    
    If IsNull(DivorceDate.Value) Then
        DivorceDate.Value = ""
        Exit Sub
    End If
    
    divorce = DivorceDate.Value
    
    If Len(Trim(divorce)) = 0 Then
        Exit Sub
    End If
    
    ' Normal case - Make sure that marriage date is after divorce date
    
    If DateValue(Marriage) > DateValue(divorce) Then
        MsgBox "Invalid marriage/divorce date relation.", vbOKOnly, "Invalid Field"
        DivorceDate.Value = ""
    End If
End Sub

Private Sub SaveMarriage_Click()
    Dim i As Long
    Dim id As Long
    Dim answer As Integer
    
    i = Marriage.ListIndex
    
    If i = -1 Then
        Exit Sub
    End If
    
    id = Marriage.Column(0, i)
    
    Dim con As ADODB.Connection
    Dim rst As New ADODB.Recordset
        
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient
    
    If MarriageRelation(curId, id) = True Then
        ' Update marriage dates
    
        rst.Open "SELECT * FROM [IsMarried] WHERE [IDPerson1] = " & curId, con, adOpenStatic, adLockOptimistic
        Do Until rst.EOF
            If id = rst.Fields("IDPerson2") Then
                rst.Fields("WeddingDate") = MarriageDate.Value
                rst.Fields("DivorceDate") = DivorceDate.Value
                rst.Update
                Exit Do
            End If
            rst.MoveNext
        Loop
        rst.Close
        
        rst.Open "SELECT * FROM [IsMarried] WHERE [IDPerson1] = " & id, con, adOpenStatic, adLockOptimistic
        Do Until rst.EOF
            If curId = rst.Fields("IDPerson2") Then
                rst.Fields("WeddingDate") = MarriageDate.Value
                rst.Fields("DivorceDate") = DivorceDate.Value
                rst.Update
                Exit Do
            End If
            rst.MoveNext
        Loop
        rst.Close
        
        LoadData (curId)
        Exit Sub
    End If
     
    If ParentRelation(curId, id) = True Then
        answer = MsgBox("Current person is the parent of the person you are trying to add as married. Do you want to continue?", _
                    vbInformation + vbYesNo + vbDefaultButton2, "Family relation")
        If answer = vbNo Then
            Exit Sub
        End If
    End If
    
    If ChildRelation(curId, id) = True Then
        answer = MsgBox("Current person is the child of the person you are trying to add as married. Do you want to continue?", _
                    vbInformation + vbYesNo + vbDefaultButton2, "Family relation")
        If answer = vbNo Then
            Exit Sub
        End If
    End If
        
    If SiblingRelation(curId, id) = True Then
        answer = MsgBox("Current person is the sibling of the person you are trying to add as married. Do you want to continue?", _
                    vbInformation + vbYesNo + vbDefaultButton2, "Family relation")
        If answer = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Add marriage to DB
    rst.Open "IsMarried", con, adOpenStatic, adLockOptimistic
    
    rst.AddNew
    rst.Fields("IDPerson1") = curId
    rst.Fields("IDPerson2") = id
    rst.Fields("WeddingDate") = MarriageDate.Value
    rst.Fields("DivorceDate") = DivorceDate.Value
    rst.Update
    rst.Close
    
    rst.Open "IsMarried", con, adOpenStatic, adLockOptimistic
    
    rst.AddNew
    rst.Fields("IDPerson1") = id
    rst.Fields("IDPerson2") = curId
    rst.Fields("WeddingDate") = MarriageDate.Value
    rst.Fields("DivorceDate") = DivorceDate.Value
    rst.Update
    rst.Close
    
    LoadData (curId)
End Sub

Private Sub RemoveMarriage_Click()
    Dim i As Long
    Dim id As Long
    
    i = Marriage.ListIndex
    
    If i = -1 Then
        Exit Sub
    End If
    
    id = Marriage.Column(0, i)
    
    Dim con As ADODB.Connection
    Dim rst As New ADODB.Recordset
        
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient
    
    rst.Open "SELECT * FROM [IsMarried] WHERE [IDPerson1] = " & curId, con, adOpenStatic, adLockOptimistic
    Do Until rst.EOF
        If id = rst.Fields("IDPerson2") Then
            rst.Delete
            rst.MoveNext
            Exit Do
        End If
        rst.MoveNext
    Loop
    rst.Close
        
    rst.Open "SELECT * FROM [IsMarried] WHERE [IDPerson1] = " & id, con, adOpenStatic, adLockOptimistic
    Do Until rst.EOF
        If curId = rst.Fields("IDPerson2") Then
            rst.Delete
            rst.MoveNext
            Exit Do
        End If
        rst.MoveNext
    Loop
    rst.Close
    
    LoadData (curId)
End Sub

Private Sub Child_Click()
    Dim i As Long
    Dim id As Long
    
    i = Child.ListIndex
    
    If i = -1 Then
        Exit Sub
    End If
    
    id = Child.Column(0, i)
    
    For i = 0 To ChildrenList.ListCount - 1
        If id = ChildrenList.Column(0, i) Then
            ChildrenList.Selected(i) = True
            Exit Sub
        End If
    Next i
    
    ChildrenList.Value = ""
End Sub

Private Sub ChildrenList_Click()
    Dim i As Long
    Dim id As String
    
    Dim con As ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient
    
    i = ChildrenList.ListIndex
    
    If i = -1 Then
        Exit Sub
    End If
    
    id = ChildrenList.Column(0, i)
    
    Child.Value = id
    
    rst.Open "People", con, adOpenStatic, adLockOptimistic
    Do Until rst.EOF
        If rst.Fields("Id") = id Then
            Exit Do
        End If
        rst.MoveNext
    Loop
    If Not IsNull(rst.Fields("Foto")) Then
        ChildImage.Picture = Application.CurrentProject.Path & "\Fotos\" & rst.Fields("Foto")
    Else
        ChildImage.Picture = ""
    End If
    rst.Close
End Sub

Private Sub CenterChild_Click()
    Dim id As Long
    
    id = ChildId()
    
    If id = -1 Then
        Exit Sub
    End If
    
    LoadData (id)
End Sub

Private Sub SaveChild_Click()
    Dim i As Long
    Dim id As Long
    Dim answer As Integer
    
    i = Child.ListIndex
    
    If i = -1 Then
        Exit Sub
    End If
    
    id = Child.Column(0, i)
    
    If ParentRelation(curId, id) = True Then
        MsgBox "This relation was already added."
        Exit Sub
    End If
    
    If ChildRelation(curId, id) = True Then
        MsgBox "This child can not be added because the child is already marked as parent of the current person."
        Exit Sub
    End If
    
    If MarriageRelation(curId, id) = True Then
        answer = MsgBox("Current person is married to the person you are trying to add as child. Do you want to continue?", _
                    vbInformation + vbYesNo + vbDefaultButton2, "Marriage relation")
        If answer = vbNo Then
            Exit Sub
        End If
    End If
        
    If SiblingRelation(curId, id) = True Then
        answer = MsgBox("Current person is sibling of the person you are trying to add as child. Do you want to continue?", _
                    vbInformation + vbYesNo + vbDefaultButton2, "Family relation")
        If answer = vbNo Then
            Exit Sub
        End If
    End If
    
    ' Add child to DB
    Dim con As ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient
    
    rst.Open "IsParent", con, adOpenStatic, adLockOptimistic
    
    rst.AddNew
    rst.Fields("IDPerson1") = curId
    rst.Fields("IDPerson2") = id
    rst.Update
    rst.Close
    
    LoadData (curId)
End Sub

Private Sub RemoveChild_Click()
    Dim i As Long
    Dim id As Long
    
    i = Child.ListIndex
    
    If i = -1 Then
        Exit Sub
    End If
    
    id = Child.Column(0, i)
    
    Dim con As ADODB.Connection
    Dim rst As New ADODB.Recordset
        
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient
    
    rst.Open "SELECT * FROM [IsParent] WHERE [IDPerson1] = " & curId, con, adOpenStatic, adLockOptimistic
    Do Until rst.EOF
        If id = rst.Fields("IDPerson2") Then
            rst.Delete
            rst.MoveNext
            Exit Do
        End If
        rst.MoveNext
    Loop
    rst.Close
    
    LoadData (curId)
End Sub

' ---------------------------
'      Auxiliary functions
' ---------------------------

Private Function FatherId(ByVal sonId As Long) As Long
    Dim con As ADODB.Connection
    Dim rst1 As New ADODB.Recordset
    Dim rst2 As New ADODB.Recordset
    
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient
    FatherId = -1
    
    rst1.Open "SELECT * FROM [IsParent] WHERE [IDPerson2] = " & sonId, con, adOpenForwardOnly, adLockOptimistic
    Do Until rst1.EOF
    
        rst2.Open "SELECT * FROM [People] WHERE [Id] = " & rst1.Fields("IDPerson1"), con, adOpenForwardOnly, adLockOptimistic
        Do Until rst2.EOF
            If rst2.Fields("Gender") = "M" Then
                FatherId = rst2.Fields("Id")
            End If
            rst2.MoveNext
        Loop
        
        rst2.Close
        rst1.MoveNext
    Loop
    
    rst1.Close
End Function

Private Function MotherId(ByVal sonId As Long) As Long
    Dim con As ADODB.Connection
    Dim rst1 As New ADODB.Recordset
    Dim rst2 As New ADODB.Recordset
    
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient
    MotherId = -1
    
    rst1.Open "SELECT * FROM [IsParent] WHERE [IDPerson2] = " & sonId, con, adOpenForwardOnly, adLockOptimistic
    Do Until rst1.EOF
    
        rst2.Open "SELECT * FROM [People] WHERE [Id] = " & rst1.Fields("IDPerson1"), con, adOpenForwardOnly, adLockOptimistic
        Do Until rst2.EOF
            If rst2.Fields("Gender") = "F" Then
                MotherId = rst2.Fields("Id")
            End If
            rst2.MoveNext
        Loop
        
        rst2.Close
        rst1.MoveNext
    Loop
    
    rst1.Close
End Function

Private Function SiblingId() As Long
    ' This function looks for the sibling that is CURRENTLY SELECTED in the listbox
    Dim i As Long
    
    i = SiblingsList.ListIndex
    SiblingId = -1
    
    If i = -1 Then
        Exit Function
    End If
    
    SiblingId = SiblingsList.Column(0, i)
End Function

Private Function MarriageId() As Long
    ' This function looks for the marriage that is CURRENTLY SELECTED in the listbox
    Dim i As Long
    
    i = MarriageList.ListIndex
    MarriageId = -1
    
    If i = -1 Then
        Exit Function
    End If
    
    MarriageId = MarriageList.Column(0, i)
End Function

Private Function ChildId() As Long
    ' This function looks for the child that is CURRENTLY SELECTED in the listbox
    Dim i As Long
    
    i = ChildrenList.ListIndex
    ChildId = -1
    
    If i = -1 Then
        Exit Function
    End If
    
    ChildId = ChildrenList.Column(0, i)
End Function

Private Function RelationsCount(ByVal id1 As Integer, ByVal id2 As Integer) As Integer
    ' Count the number of relations, each person can only be in one relation with another
    Dim i As Integer

    i = 0

    If ParentRelation(id1, id2) = True Then
        i = i + 1
    End If

    If SiblingRelation(id1, id2) = True Then
        i = i + 1
    End If

    If MarriageRelation(id1, id2) = True Then
        i = i + 1
    End If
    
    If ChildRelation(id1, id2) = True Then
        i = i + 1
    End If
    
    RelationsCount = i
End Function

Private Function ParentRelation(ByVal id1 As Integer, ByVal id2 As Integer) As Boolean
    Dim con As ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    ParentRelation = False
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient
    
    rst.Open "SELECT * FROM [IsParent] WHERE [IDPerson1] = " & id1, con, , adLockOptimistic
    Do Until rst.EOF
        If id2 = rst.Fields("IDPerson2") Then
            ParentRelation = True
            Exit Do
        End If
        rst.MoveNext
    Loop
    
    rst.Close
End Function

Private Function SiblingRelation(ByVal id1 As Integer, ByVal id2 As Integer) As Boolean
    Dim con As ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    SiblingRelation = False
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient
    
    rst.Open "SELECT * FROM [IsSibling] WHERE [IDPerson1] = " & id1, con, , adLockOptimistic
    Do Until rst.EOF
        If id2 = rst.Fields("IDPerson2") Then
            SiblingRelation = True
            Exit Do
        End If
        rst.MoveNext
    Loop

    rst.Close
End Function

Private Function MarriageRelation(ByVal id1 As Integer, ByVal id2 As Integer) As Boolean
    Dim con As ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    MarriageRelation = False
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient
    
    rst.Open "SELECT * FROM [IsMarried] WHERE [IDPerson1] = " & id1, con, , adLockOptimistic
    Do Until rst.EOF
        If id2 = rst.Fields("IDPerson2") Then
            MarriageRelation = True
            ' This exit is important, the others are only for efficiency
            Exit Do
        End If
        rst.MoveNext
    Loop
    
    rst.Close
End Function

Private Function ChildRelation(ByVal id1 As Integer, ByVal id2 As Integer) As Boolean
    Dim con As ADODB.Connection
    Dim rst As New ADODB.Recordset
    
    ChildRelation = False
    Set con = CurrentProject.Connection
    con.CursorLocation = adUseClient
    
    rst.Open "SELECT * FROM [IsParent] WHERE [IDPerson1] = " & id2, con, , adLockOptimistic
    Do Until rst.EOF
        If id1 = rst.Fields("IDPerson2") Then
            ChildRelation = True
            Exit Do
        End If
        rst.MoveNext
    Loop
    
    rst.Close
End Function
