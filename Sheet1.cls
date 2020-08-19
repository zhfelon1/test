VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


'·������'
Const AudioName As String = "\game\audio"
Const imageName As String = "\game\images"
Const BgmName As String = "\game"
Const OutputName As String = "\game\script.rpy"

'�ؼ���KeyWord����'
Const KW_Scene As String = "ת��"
Const KW_Choose As String = "ѡ��"


'������Id'
Const Column_Role As Long = 1
Const Column_Pos As Long = 2
Const Column_Chat As Long = 3
Const Column_Bgm As Long = 4
Const Column_Drawing As Long = 5
Const Column_BackGround As Long = 6
Const Column_Trans As Long = 7



'����'
Const ChrDefaultColor As String = "#c8ffc8"  'todo: Ĭ�ϵĽ�ɫ������ɫ'

Const ChrNamePrefix As String = "Chr_"
Const ImageNamePrefix As String = "Image_"
Const BackGroundPrefix As String = "Scene_"



'��ȡRenpy�Ĺ�������'
Public Function ProjectName()
    ProjectName = Sheet1.Cells(11, 2)
End Function



'����·��'
Public Function ProjectPath()
    ProjectPath = ProjectName()
End Function


'��������'
Public Function ImageScale()
    ImageScale = Sheet1.Cells(12, 2)
End Function


'-------------��ť����------------'

'������ƵĿ¼'
Sub OpenAudioPath()
    Dim path
    path = ProjectPath() & AudioName
    Shell "explorer.exe " & path, vbNormalFocus
End Sub

'����ͼƬĿ¼'
Sub OpenImagePath()
    Dim path
    path = ProjectPath() & imageName
    Shell "explorer.exe " & path, vbNormalFocus
End Sub

'����BgmĿ¼'
Sub OpenBgmPath()
    Dim path
    path = ProjectPath() & BgmName
    Shell "explorer.exe " & path, vbNormalFocus
End Sub

'Ԥ���ű��ļ�'
Sub OpenPyFile()
Attribute OpenPyFile.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim path
    path = ProjectPath() & OutputName
    Shell "notepad.exe " & path, vbNormalFocus
End Sub


'����'
Sub OutPutPy()
Attribute OutPutPy.VB_ProcData.VB_Invoke_Func = "t\n14"

    Dim content As String
    Dim sh As Worksheet

    '---- ���������� ----'
    For Each sh In Worksheets
       
        '---- ��һ�����ñ������� ----'
        If Not sh.Name Like "����" Then
        
            Dim arr As Variant
            arr = sh.UsedRange
            If IsArray(arr) Then
                content = content & GenStr(sh.Name, arr)
            End If
        End If
    Next
    
    
    '---- д���ļ� ----'
    Dim fileName As String
    fileName = ProjectPath() & OutputName
    
    Dim objStream As Object
    Set objStream = CreateObject("ADODB.Stream")
    
    With objStream
        .Type = 2
        .Charset = "UTF-8"
        .Open
        .WriteText content
        .SaveToFile fileName, 2
    End With
    objStream.Close
     
    MsgBox ("�����ɹ�")
End Sub


'------------�����߼�------------'

'����һ���ַ����Ƿ���������'
Function FindStrInArr(arr, str As String)

    Dim idx As Integer
    Dim bFind As Boolean
    
    bFind = False
    Debug.Print (LBound(arr))
    If IsArray(arr) Then
        For idx = LBound(arr) To UBound(arr)
        
            If arr(idx) = str Then
                bFind = True
                Exit For
            End If
        Next
    End If
    FindStrInArr = bFind
    
End Function


'ɾ����չ��'
Function DelExt(str As String)
    DelExt = Left(str, InStr(1, str, ".") - 1)
End Function


'����Ƿ�Ϊ�ؼ���'
Function IsKeyWord(str As String)

    If str = KW_Scene Or str = KW_Choose Then
        IsKeyWord = True
    Else
        IsKeyWord = False
    End If
 
End Function


'���ɽ�ɫ��Ϣ'
Function GenRoleList(arr As Variant)

    '---- ��ʼ��ȡ��ɫ��Ϣ��д�� ----'
    Dim rowCount As Long
    Dim cell As String
    Dim roleArr(1000) As String
    Dim roleArrIdx As Integer
    
    roleArrIdx = 0
    
    For rowCount = LBound(arr) To UBound(arr)
    
        '---- ��һ�б����У������� ----'
        If rowCount <> 1 Then
        
            '---- ��һ��Ϊ��ɫ�У�ȥ�ز��������� ----'
            cell = arr(rowCount, Column_Role)
            
            If cell <> "" Then
            
                '---- �ų��ؼ��ֲ�ȥ�� ----'
                If IsKeyWord(cell) = False And FindStrInArr(roleArr, cell) = False Then
                    roleArr(roleArrIdx) = cell
                    roleArrIdx = roleArrIdx + 1
                End If
            End If
        End If
    Next
    
    
    '---- ���ɽ�ɫ�����ַ��� ----'
    Dim idx As Long
    Dim roleStr As String
    For idx = LBound(roleArr) To UBound(roleArr)
        If roleArr(idx) <> "" Then
            roleStr = ChrNamePrefix & roleArr(idx)
            GenRoleList = GenRoleList & ("define " & roleStr & " = Character('" & roleArr(idx) & "', color=""" & ChrDefaultColor & """)" & vbCrLf)
        End If
    Next
    
End Function


'����imageList'
Function GenImageList(arr As Variant)

    '---- ��ʼ��ȡ��ɫ��Ϣ��д�� ----'
    Dim rowCount As Long
    Dim cell As String
    Dim imageArr(1000) As String
    Dim imageArrIdx As Integer
    
    imageArrIdx = 0
    
    For rowCount = LBound(arr) To UBound(arr)
    
        '---- ��һ�б����У������� ----'
        If rowCount <> 1 Then
        
            '---- ��һ��Ϊ��ɫ�У�ȥ�ز��������� ----'
            cell = arr(rowCount, Column_Drawing)
            
            If cell <> "" Then
            
                '---- ȥ�� ----'
                If FindStrInArr(imageArr, cell) = False Then
                    imageArr(imageArrIdx) = cell
                    imageArrIdx = imageArrIdx + 1
                End If
            End If
        End If
    Next
    
    
    '---- ���ɽ�ɫ�����ַ��� ----'
    Dim idx As Long
    Dim imageName As String
    
    For idx = LBound(imageArr) To UBound(imageArr)
        If imageArr(idx) <> "" Then
            imageName = imageArr(idx)
            imageName = DelExt(imageName)
            imageName = ImageNamePrefix & imageName
            
            GenImageList = GenImageList & "image " & imageName & " = im.FactorScale(""" & imageArr(idx) & """, " & ImageScale & ")" & vbCrLf
        End If
    Next
    
End Function


'����backGround'
Function GenBackGroundList(arr As Variant)

    '---- ��ʼ��ȡ��ɫ��Ϣ��д�� ----'
    Dim rowCount As Long
    Dim cell As String
    Dim imageArr(1000) As String
    Dim imageArrIdx As Integer
    
    imageArrIdx = 0
    
    For rowCount = LBound(arr) To UBound(arr)
    
        '---- ��һ�б����У������� ----'
        If rowCount <> 1 Then
        
            '---- ��һ��Ϊ��ɫ�У�ȥ�ز��������� ----'
            cell = arr(rowCount, Column_BackGround)
            
            If cell <> "" Then
            
                '---- ȥ�� ----'
                If FindStrInArr(imageArr, cell) = False Then
                    imageArr(imageArrIdx) = cell
                    imageArrIdx = imageArrIdx + 1
                End If
            End If
        End If
    Next
    
    
    '---- ���ɽ�ɫ�����ַ��� ----'
    Dim idx As Long
    Dim imageName As String
    
    For idx = LBound(imageArr) To UBound(imageArr)
        If imageArr(idx) <> "" Then
            imageName = imageArr(idx)
            imageName = DelExt(imageName)
            imageName = BackGroundPrefix & imageName
            
            GenBackGroundList = GenBackGroundList & "image " & imageName & " = im.Scale(""" & imageArr(idx) & """, config.screen_width, config.screen_height)" & vbCrLf
        End If
    Next
    
End Function



'---- ������� ----'
Function ClearDraw(draw As String)
    If draw <> "" Then
        ClearDraw = ClearDraw & "    hide " & draw
        ClearDraw = ClearDraw & vbCrLf & vbCrLf
    End If
End Function

Function GetPos(pos As String)
    If pos = "��" Then
        GetPos = " at left"
    ElseIf pos = "��" Then
        GetPos = " at right"
    End If
End Function



'�����籾��'
Function GenStr(sheetName As String, arr As Variant)
   
    '���ɽ�ɫͷ�ļ���Ϣ'
    GenStr = GenStr & GenRoleList(arr) & vbCrLf
 
    GenStr = GenStr & GenImageList(arr) & vbCrLf
    
    GenStr = GenStr & GenBackGroundList(arr) & vbCrLf
    
    'label start:'
    GenStr = GenStr & vbCrLf & vbCrLf
    GenStr = GenStr & "label start:"
    GenStr = GenStr & vbCrLf
    
    '������������������Ϣ'
    Dim rowCount As Long
    
    
    Dim pos As String
    
    Dim hideDraw As String
    
   
    Dim bgPrevName As String
    Dim bgName As String
    
    Dim drawName As String
    Dim prevLeftDraw As String
    Dim prevRightDraw As String
    Dim prevMidDraw As String
    
    Dim prevRole As String
    Dim prevPos As String
     
    Dim cellRole As String
    Dim cellBackGround As String
    Dim cellBgm As String
    Dim cellDraw As String
    Dim cellChat As String
    Dim cellEffect As String
    
    
    For rowCount = LBound(arr) To UBound(arr)
    
        If rowCount <> 1 Then
            
            '��ɫ��'
            cellRole = arr(rowCount, Column_Role)
            
            '��ɫ����Ϊ�պ͹ؼ��ִ�������'
            If cellRole <> "" Or prevRole <> "" Then
            
                If cellRole <> "" Then
                    prevRole = cellRole
                End If
                  
                '����Bgm'
                cellBgm = arr(rowCount, Column_Bgm)
                If cellBgm <> "" Then
                    GenStr = GenStr & "    play music """ & cellBgm & """ fadeout 1.0 fadein 1.0"
                    GenStr = GenStr & vbCrLf & vbCrLf
                End If
                
                
                
                 '��������'
                cellBackGround = arr(rowCount, Column_BackGround)
                If cellBackGround <> "" Then

                    bgName = DelExt(cellBackGround)
                    bgName = BackGroundPrefix & bgName
            
                    
                     '����scene image, ʹ�ô��ڴ�С����'
                    GenStr = GenStr & "    scene " & bgName
                    GenStr = GenStr & vbCrLf & vbCrLf
                    
                     '������һ�ű���'
                    hideDraw = ClearDraw(bgPrevName)
                    If hideDraw <> "" Then
                        GenStr = GenStr & hideDraw
                        bgPrevName = ""
                    End If
                    
                    bgPrevName = bgName
                End If
                
              
                '��������'
                pos = arr(rowCount, Column_Pos)
                cellDraw = arr(rowCount, Column_Drawing)
                
                If pos <> "" Then
                    prevPos = pos
                End If
                
                If cellDraw <> "" Then
            
                    drawName = DelExt(cellDraw)
                    drawName = ImageNamePrefix & drawName

                    If prevPos = "��" Then
                        
                        hideDraw = ClearDraw(prevLeftDraw)
                        If hideDraw <> "" Then
                            GenStr = GenStr & hideDraw
                            prevLeftDraw = ""
                        End If
                        
                        prevLeftDraw = drawName
                    
                    ElseIf prevPos = "��" Then
                    
                        hideDraw = ClearDraw(prevRightDraw)
                        If hideDraw <> "" Then
                            GenStr = GenStr & hideDraw
                            prevRightDraw = ""
                        End If
                                                    
                        prevRightDraw = drawName
                    Else
                        
                        hideDraw = ClearDraw(prevLeftDraw)
                        If hideDraw <> "" Then
                            GenStr = GenStr & hideDraw
                            prevLeftDraw = ""
                        End If
                        
                        hideDraw = ClearDraw(prevRightDraw)
                        If hideDraw <> "" Then
                            GenStr = GenStr & hideDraw
                            prevRightDraw = ""
                        End If
                        
                        hideDraw = ClearDraw(prevMidDraw)
                        If hideDraw <> "" Then
                            GenStr = GenStr & hideDraw
                            prevMidDraw = ""
                        End If
                        prevMidDraw = drawName
                        
                    End If
                    
                    GenStr = GenStr & "    show " & drawName & GetPos(prevPos)
                    GenStr = GenStr & vbCrLf & vbCrLf
                    
                End If
                
                '������Ч'
                cellEffect = arr(rowCount, Column_Trans)
                If cellEffect <> "" Then
                    GenStr = GenStr & "    with " & cellEffect
                    GenStr = GenStr & vbCrLf & vbCrLf
                End If
                
                
                
                '�����Ի�'
                cellChat = arr(rowCount, Column_Chat)
               
                If cellChat <> "" And cellChat <> " " Then
                
                    cellChat = Replace(cellChat, """", "\""")
                    If prevRole = KW_Scene Then
                        GenStr = GenStr & "    " & """" & cellChat & """"
                    Else
                        GenStr = GenStr & "    " & ChrNamePrefix & prevRole & " """ & cellChat & """"
                    End If
                    
                    GenStr = GenStr & vbCrLf & vbCrLf
                End If
                
            End If
            
        End If
    Next
    
 
End Function




