Attribute VB_Name = "TestDicSort"
Option Explicit
Option Private Module

' TestDicSort: �A�z�z��Dictionary���L�[�̏����ɕ��ׂ�����Sub DicSort���e�X�g����

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Private Sub TestCleanup()
    'this method runs after every test in the module.
End Sub


'@TestMethod("DicSort���e�X�g")
Private Sub TestAccessTable1()
    On Error GoTo TestFail
    'Arrange:
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    dic("g") = "gggg"
    dic("9") = "999"
    dic("��") = "��������"
    dic("4") = "444"
    dic("��") = "��������"
    dic("(") = "(((("
    dic("a") = "aaaa"
    'Assert:
    Dim Keys() As Variant    ' Keys�Ƃ������O��dynamic array��錾���Ă���
    Keys = dic.Keys
    Assert.AreEqual "g", Keys(0)
    Assert.AreEqual "9", Keys(1)
    Assert.AreEqual "��", Keys(2)
    
    'Act:
    Dim output As String
    Dim key As Variant
    output = "##before------------" & vbNewLine
    For Each key In dic
        output = output & key & ":" & dic(key) & vbNewLine
    Next key
    
    ' now sort it
    Call DicSort(dic)
    
    output = output & vbNewLine & "##after-------------" & vbNewLine
    For Each key In dic
        output = output & key & ":" & dic(key) & vbNewLine
    Next key
    
    Call KzCls
    Debug.Print output
    
    'Assert:
    Keys = dic.Keys
    Assert.AreEqual "(", Keys(0)
    Assert.AreEqual "4", Keys(1)
    Assert.AreEqual "9", Keys(2)

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
