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


'路径配置'
Const AudioName As String = "\game\audio"
Const imageName As String = "\game\images"
Const BgmName As String = "\game"
Const OutputName As String = "\game\script.rpy"

'关键字KeyWord定义'
Const KW_Scene As String = "转场"
Const KW_Choose As String = "选项"


'列内容Id'
Const Column_Role As Long = 1
Const Column_Pos As Long = 2
Const Column_Chat As Long = 3
Const Column_Bgm As Long = 4
Const Column_Drawing As Long = 5
Const Column_BackGround As Long = 6
Const Column_Trans As Long = 7



'其他'
Const ChrDefaultColor As String = "#c8ffc8"  'todo: 默认的角色名字颜色'

Const ChrNamePrefix As String = "Chr_"
Const ImageNamePrefix As String = "Image_"
Const BackGroundPrefix As String = "Scene_"



'读取Renpy的工程名称'
Public Function ProjectName()
    ProjectName = Sheet1.Cells(11, 2)
End Function



'工程路径'
Public Function ProjectPath()
    ProjectPath = ProjectName()
End Function


'立绘缩放'
Public Function ImageScale()
    ImageScale = Sheet1.Cells(12, 2)
End Function


'-------------按钮功能------------'

'开启音频目录'
Sub OpenAudioPath()
    Dim path
    path = ProjectPath() & AudioName
    Shell "explorer.exe " & path, vbNormalFocus
End Sub

'开启图片目录'
Sub OpenImagePath()
    Dim path
    path = ProjectPath() & imageName
    Shell "explorer.exe " & path, vbNormalFocus
End Sub

'开启Bgm目录'
Sub OpenBgmPath()
    Dim path
    path = ProjectPath() & BgmName
    Shell "explorer.exe " & path, vbNormalFocus
End Sub

'预览脚本文件'
Sub OpenPyFile()
Attribute OpenPyFile.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim path
    path = ProjectPath() & OutputName
    Shell "notepad.exe " & path, vbNormalFocus
End Sub


'导出'
Sub OutPutPy()
Attribute OutPutPy.VB_ProcData.VB_Invoke_Func = "t\n14"

    Dim content As String
    Dim sh As Worksheet

    '---- 遍历工作表 ----'
    For Each sh In Worksheets
       
        '---- 第一张配置表不处理 ----'
        If Not sh.Name Like "配置" Then
        
            Dim arr As Variant
            arr = sh.UsedRange
            If IsArray(arr) Then
                content = content & GenStr(sh.Name, arr)
            End If
        End If
    Next
    
    
    '---- 写入文件 ----'
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
     
    MsgBox ("导出成功")
End Sub


'------------导出逻辑------------'

'查找一个字符串是否在数组里'
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


'删除扩展名'
Function DelExt(str As String)
    DelExt = Left(str, InStr(1, str, ".") - 1)
End Function


'检查是否为关键字'
Function IsKeyWord(str As String)

    If str = KW_Scene Or str = KW_Choose Then
        IsKeyWord = True
    Else
        IsKeyWord = False
    End If
 
End Function


'生成角色信息'
Function GenRoleList(arr As Variant)

    '---- 开始读取角色信息并写入 ----'
    Dim rowCount As Long
    Dim cell As String
    Dim roleArr(1000) As String
    Dim roleArrIdx As Integer
    
    roleArrIdx = 0
    
    For rowCount = LBound(arr) To UBound(arr)
    
        '---- 第一行保留行，不处理 ----'
        If rowCount <> 1 Then
        
            '---- 第一列为角色列，去重并放入数组 ----'
            cell = arr(rowCount, Column_Role)
            
            If cell <> "" Then
            
                '---- 排除关键字并去重 ----'
                If IsKeyWord(cell) = False And FindStrInArr(roleArr, cell) = False Then
                    roleArr(roleArrIdx) = cell
                    roleArrIdx = roleArrIdx + 1
                End If
            End If
        End If
    Next
    
    
    '---- 生成角色配置字符串 ----'
    Dim idx As Long
    Dim roleStr As String
    For idx = LBound(roleArr) To UBound(roleArr)
        If roleArr(idx) <> "" Then
            roleStr = ChrNamePrefix & roleArr(idx)
            GenRoleList = GenRoleList & ("define " & roleStr & " = Character('" & roleArr(idx) & "', color=""" & ChrDefaultColor & """)" & vbCrLf)
        End If
    Next
    
End Function


'生成imageList'
Function GenImageList(arr As Variant)

    '---- 开始读取角色信息并写入 ----'
    Dim rowCount As Long
    Dim cell As String
    Dim imageArr(1000) As String
    Dim imageArrIdx As Integer
    
    imageArrIdx = 0
    
    For rowCount = LBound(arr) To UBound(arr)
    
        '---- 第一行保留行，不处理 ----'
        If rowCount <> 1 Then
        
            '---- 第一列为角色列，去重并放入数组 ----'
            cell = arr(rowCount, Column_Drawing)
            
            If cell <> "" Then
            
                '---- 去重 ----'
                If FindStrInArr(imageArr, cell) = False Then
                    imageArr(imageArrIdx) = cell
                    imageArrIdx = imageArrIdx + 1
                End If
            End If
        End If
    Next
    
    
    '---- 生成角色配置字符串 ----'
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


'生成backGround'
Function GenBackGroundList(arr As Variant)

    '---- 开始读取角色信息并写入 ----'
    Dim rowCount As Long
    Dim cell As String
    Dim imageArr(1000) As String
    Dim imageArrIdx As Integer
    
    imageArrIdx = 0
    
    For rowCount = LBound(arr) To UBound(arr)
    
        '---- 第一行保留行，不处理 ----'
        If rowCount <> 1 Then
        
            '---- 第一列为角色列，去重并放入数组 ----'
            cell = arr(rowCount, Column_BackGround)
            
            If cell <> "" Then
            
                '---- 去重 ----'
                If FindStrInArr(imageArr, cell) = False Then
                    imageArr(imageArrIdx) = cell
                    imageArrIdx = imageArrIdx + 1
                End If
            End If
        End If
    Next
    
    
    '---- 生成角色配置字符串 ----'
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



'---- 清除立绘 ----'
Function ClearDraw(draw As String)
    If draw <> "" Then
        ClearDraw = ClearDraw & "    hide " & draw
        ClearDraw = ClearDraw & vbCrLf & vbCrLf
    End If
End Function

Function GetPos(pos As String)
    If pos = "左" Then
        GetPos = " at left"
    ElseIf pos = "右" Then
        GetPos = " at right"
    End If
End Function



'处理剧本表'
Function GenStr(sheetName As String, arr As Variant)
   
    '生成角色头文件信息'
    GenStr = GenStr & GenRoleList(arr) & vbCrLf
 
    GenStr = GenStr & GenImageList(arr) & vbCrLf
    
    GenStr = GenStr & GenBackGroundList(arr) & vbCrLf
    
    'label start:'
    GenStr = GenStr & vbCrLf & vbCrLf
    GenStr = GenStr & "label start:"
    GenStr = GenStr & vbCrLf
    
    '遍历表，生成其他信息'
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
            
            '角色名'
            cellRole = arr(rowCount, Column_Role)
            
            '角色名不为空和关键字处理此行'
            If cellRole <> "" Or prevRole <> "" Then
            
                If cellRole <> "" Then
                    prevRole = cellRole
                End If
                  
                '处理Bgm'
                cellBgm = arr(rowCount, Column_Bgm)
                If cellBgm <> "" Then
                    GenStr = GenStr & "    play music """ & cellBgm & """ fadeout 1.0 fadein 1.0"
                    GenStr = GenStr & vbCrLf & vbCrLf
                End If
                
                
                
                 '处理背景'
                cellBackGround = arr(rowCount, Column_BackGround)
                If cellBackGround <> "" Then

                    bgName = DelExt(cellBackGround)
                    bgName = BackGroundPrefix & bgName
            
                    
                     '声明scene image, 使用窗口大小缩放'
                    GenStr = GenStr & "    scene " & bgName
                    GenStr = GenStr & vbCrLf & vbCrLf
                    
                     '销毁上一张背景'
                    hideDraw = ClearDraw(bgPrevName)
                    If hideDraw <> "" Then
                        GenStr = GenStr & hideDraw
                        bgPrevName = ""
                    End If
                    
                    bgPrevName = bgName
                End If
                
              
                '处理立绘'
                pos = arr(rowCount, Column_Pos)
                cellDraw = arr(rowCount, Column_Drawing)
                
                If pos <> "" Then
                    prevPos = pos
                End If
                
                If cellDraw <> "" Then
            
                    drawName = DelExt(cellDraw)
                    drawName = ImageNamePrefix & drawName

                    If prevPos = "左" Then
                        
                        hideDraw = ClearDraw(prevLeftDraw)
                        If hideDraw <> "" Then
                            GenStr = GenStr & hideDraw
                            prevLeftDraw = ""
                        End If
                        
                        prevLeftDraw = drawName
                    
                    ElseIf prevPos = "右" Then
                    
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
                
                '处理特效'
                cellEffect = arr(rowCount, Column_Trans)
                If cellEffect <> "" Then
                    GenStr = GenStr & "    with " & cellEffect
                    GenStr = GenStr & vbCrLf & vbCrLf
                End If
                
                
                
                '处理对话'
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





