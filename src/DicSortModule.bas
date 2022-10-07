Attribute VB_Name = "DicSortModule"
Option Explicit


' DicSortModule: Dictionary�A�z�z����\�[�g����Sub DicSort��񋟂���

' Qiita�Ɍf�ڂ��ꂽ�L�� [VBA��Dictionary�i�A�z�z��j���������Ƀ\�[�g����](https://qiita.com/daik/items/682743bb8bcd8b5f0689)��
' ���J�����R�[�h���܂�܂�R�s�[�����B


'' Dictionary���Q�ƈ����ɂ��A������\�[�g����j��I�v���V�[�W���B
Public Sub DicSort(ByRef dic As Object)

  Dim i As Long, j As Long, dicSize As Long
  Dim varTmp() As String
  Dim key As Variant

  dicSize = dic.Count

  ReDim varTmp(dicSize + 1, 2)

  ' Dictionary���󂩁A�T�C�Y��1�ȉ��ł���΃\�[�g�s�v
  If dic Is Nothing Or dicSize < 2 Then
    Exit Sub
  End If

  ' Dictionary����񌳔z��ɓ]��
  i = 0
  For Each key In dic
    varTmp(i, 0) = key
    varTmp(i, 1) = dic(key)
    i = i + 1
  Next

  '�N�C�b�N�\�[�g
  Call QuickSort(varTmp, 0, dicSize - 1)

  dic.RemoveAll

  For i = 0 To dicSize - 1
    dic(varTmp(i, 0)) = varTmp(i, 1)
  Next
End Sub


'' String�^��2��̓񎟌��z����󂯎��A�����1��ڂŃN�C�b�N�\�[�g����i�ق�Ƃ�Compare���\�b�h��n��Adapter�p�^�[���ŏ��������Ƃ���AVBA�̃I�u�W�F�N�g�w�������������őË��j
Private Sub QuickSort(ByRef targetVar() As String, ByVal min As Long, ByVal max As Long)
    Dim i, j As Long
    Dim tmp As String
    Dim pivot As Variant

    If min < max Then
        i = min
        j = max
        pivot = strMed3(targetVar(i, 0), targetVar(Int((i + j) / 2), 0), targetVar(j, 0))
        Do
            Do While StrComp(targetVar(i, 0), pivot) < 0
                i = i + 1
            Loop
            Do While StrComp(pivot, targetVar(j, 0)) < 0
                j = j - 1
            Loop
            If i >= j Then Exit Do

            tmp = targetVar(i, 0)
            targetVar(i, 0) = targetVar(j, 0)
            targetVar(j, 0) = tmp

            tmp = targetVar(i, 1)
            targetVar(i, 1) = targetVar(j, 1)
            targetVar(j, 1) = tmp

            i = i + 1
            j = j - 1

        Loop
        Call QuickSort(targetVar, min, i - 1)
        Call QuickSort(targetVar, j + 1, max)

    End If
End Sub


'' String�^��x, y, z ����������r����Ԗڂ̂��̂�Ԃ�
Private Function strMed3(ByVal x As String, ByVal y As String, ByVal z As String)
    If StrComp(x, y) < 0 Then
        If StrComp(y, z) < 0 Then
            strMed3 = y
        ElseIf StrComp(z, x) < 0 Then
            strMed3 = x
        Else
            strMed3 = z
        End If
    Else
        If StrComp(z, y) < 0 Then
            strMed3 = y
        ElseIf StrComp(x, z) < 0 Then
            strMed3 = x
        Else
            strMed3 = z
        End If
    End If
End Function




'' �e�X�g���\�b�h
Sub TestDicSort()

    Dim output As String
    Dim dic As Object
    Dim key As Variant

    Set dic = CreateObject("Scripting.Dictionary")

    dic("g") = "gggg"
    dic("9") = "999"
    dic("��") = "��������"
    dic("4") = "444"
    dic("��") = "��������"
    dic("(") = "(((("
    dic("a") = "aaaa"

    output = "##before" & vbNewLine

    For Each key In dic
        output = output & key & ":" & dic(key) & vbNewLine
    Next key

    Call DicSort(dic)

    output = output + vbNewLine & "##after" & vbNewLine
    For Each key In dic
        output = output & key & ":" & dic(key) & vbNewLine
    Next key

    Debug.Print output
End Sub
