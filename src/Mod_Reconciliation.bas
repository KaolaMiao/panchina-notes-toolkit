Attribute VB_Name = "Mod_Reconciliation"
Option Explicit

' ==========================================================================================
' 模块名称: 动态数据核对工具 (Mod_Data_Validation_V22)
' 功能描述: 双重稽核校验
' 更新内容: 支持输入【多组关键词】(逗号分隔)，只要匹配其中任意一个即命中
' ==========================================================================================

' --- Master 配置 ---
Private Const SHEET_NAME_MASTER As String = "Master"
Private Const COL_M_FILE As Integer = 1        ' A列: 文件名
Private Const COL_M_ITEM As Integer = 4        ' D列: ItemName
Private Const COL_M_METRIC As Integer = 5      ' E列: 指标类型
Private Const COL_M_VALUE As Integer = 6       ' F列: 数值
Private Const COL_M_ACCOUNT As Integer = 8     ' H列: 科目名称
Private Const COL_NAME_FINAL_AUDIT As String = "审定数"

Public Sub Run_Data_Reconciliation()
    Dim wsTB As Worksheet, wsMaster As Worksheet, wsReport As Worksheet
    
    ' 字典定义
    Dim dictExplicitTotal As Object, dictDetailSum As Object, dictHasTotalRow As Object
    
    Dim arrMaster As Variant
    Dim arrOutput() As Variant
    Dim r As Long, c As Long, outRow As Long, outCol As Long
    Dim lastRowM As Long
    Dim key As String, fileKey As String, subjectKey As String, itemName As String
    Dim valMasterTotalRow As Double, valMasterDetailSum As Double, valTB As Double
    Dim diffInternal As Double, diffExternal As Double
    Dim cellText As String
    
    ' 交互变量
    Dim rngColSelection As Range, rngRowSelection As Range
    Dim strParentFileMaster As String
    Dim startCol As Long, endCol As Long
    Dim isFirstColParent As Boolean
    Dim headerRow As Long
    Dim strSubjectName As String
    
    ' --- 【新增变量】 多关键词处理 ---
    Dim strTargetMetric As String
    Dim arrKeywords As Variant
    Dim vKey As Variant
    Dim isMatch As Boolean
    
    ' 1. 初始化
    Set wsTB = ActiveSheet
    On Error Resume Next
    Set wsMaster = ActiveWorkbook.Sheets(SHEET_NAME_MASTER)
    On Error GoTo 0
    If wsMaster Is Nothing Then MsgBox "未找到 Master 表！", vbCritical: Exit Sub
    
    ' ==========================================================================
    ' 2. 用户交互 (选区配置)
    ' ==========================================================================
    On Error Resume Next
    Set rngColSelection = Application.InputBox("步骤 1/4: 框选表头列范围", "列范围", Type:=8)
    If rngColSelection Is Nothing Then Exit Sub
    
    headerRow = rngColSelection.Row
    Set wsTB = rngColSelection.Parent
    startCol = rngColSelection.Column
    endCol = startCol + rngColSelection.Columns.count - 1
    
    Set rngRowSelection = Application.InputBox("步骤 2/4: 框选科目列(C列)数据区", "行范围", Type:=8)
    If rngRowSelection Is Nothing Then Exit Sub
    Set rngRowSelection = Application.Intersect(rngRowSelection, wsTB.UsedRange)
    If rngRowSelection Is Nothing Then Exit Sub
    
    ' 判断首列映射
    Dim firstColHeader As String
    firstColHeader = Trim(CStr(wsTB.Cells(headerRow, startCol).Value))
    If firstColHeader <> COL_NAME_FINAL_AUDIT Then
        strParentFileMaster = InputBox("步骤 3/4: 确认母公司文件名", "映射", firstColHeader)
        If StrPtr(strParentFileMaster) = 0 Then Exit Sub
        isFirstColParent = True
    Else
        isFirstColParent = False
    End If

    ' --- 【新增交互】 询问要核对的指标 (支持多组) ---
    strTargetMetric = InputBox("步骤 4/4: 请输入Master表中要核对的指标关键词" & vbCrLf & vbCrLf & _
                               "● 规则：输入多个词用逗号隔开，匹配任意一个即可。" & vbCrLf & _
                               "● 示例：输入 '期末,本年' 将同时抓取 '期末数' 和 '本年累计数'。" & vbCrLf & _
                               "默认值：期末数", "多指标设定", "期末数")
    If StrPtr(strTargetMetric) = 0 Or strTargetMetric = "" Then Exit Sub
    
    ' 预处理：将中文逗号替换为英文逗号，并分割成数组
    strTargetMetric = Replace(strTargetMetric, "，", ",")
    arrKeywords = Split(strTargetMetric, ",")
    
    ' --- 【断点 1】 配置完成确认 ---
    If MsgBox("配置已完成！" & vbCrLf & "核对范围: [" & strTargetMetric & "]" & vbCrLf & _
              "即将开始读取 Master 数据..." & vbCrLf & vbCrLf & _
              "点击【确定】继续", vbOKCancel + vbInformation, "阶段 1/3: 准备就绪") = vbCancel Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.StatusBar = "正在读取 Master 数据..."
    
    ' ==========================================================================
    ' 3. Master 数据聚合 (多关键词匹配逻辑)
    ' ==========================================================================
    Set dictExplicitTotal = CreateObject("Scripting.Dictionary")
    Set dictDetailSum = CreateObject("Scripting.Dictionary")
    Set dictHasTotalRow = CreateObject("Scripting.Dictionary")
    
    lastRowM = wsMaster.Cells(wsMaster.Rows.count, COL_M_FILE).End(xlUp).Row
    arrMaster = wsMaster.Range(wsMaster.Cells(1, 1), wsMaster.Cells(lastRowM, 8)).Value
    
    For r = 2 To UBound(arrMaster, 1)
        Dim currentMetric As String
        currentMetric = Trim(CStr(arrMaster(r, COL_M_METRIC)))
        
        ' ======================================================================
        ' 【修改点】 循环检查所有关键词，只要命中一个 isMatch 就为 True
        ' ======================================================================
        isMatch = False
        For Each vKey In arrKeywords
            If Len(Trim(vKey)) > 0 Then
                If InStr(1, currentMetric, Trim(vKey), vbTextCompare) > 0 Then
                    isMatch = True
                    Exit For ' 只要匹配到一个，就不需要继续检查后面的词了
                End If
            End If
        Next vKey
        
        If isMatch Then
            fileKey = Trim(CStr(arrMaster(r, COL_M_FILE)))
            subjectKey = CleanKey(CStr(arrMaster(r, COL_M_ACCOUNT)))
            itemName = Replace(Trim(CStr(arrMaster(r, COL_M_ITEM))), " ", "")
            
            Dim curVal As Double
            curVal = val(arrMaster(r, COL_M_VALUE))
            
            key = fileKey & "|" & subjectKey
            Dim totalKey As String
            totalKey = "TOTAL_SUM|" & subjectKey
            
            If InStr(itemName, "合计") > 0 Then
                ' A. 合计行
                AddVal dictExplicitTotal, key, curVal
                AddVal dictExplicitTotal, totalKey, curVal
                If Not dictHasTotalRow.exists(key) Then dictHasTotalRow.Add key, True
                If Not dictHasTotalRow.exists(totalKey) Then dictHasTotalRow.Add totalKey, True
            Else
                ' B. 明细行
                AddVal dictDetailSum, key, curVal
                AddVal dictDetailSum, totalKey, curVal
            End If
        End If
    Next r
    
    ' --- 【断点 2】 Master加载完成确认 ---
    Application.ScreenUpdating = True
    If MsgBox("Master 数据加载完毕！" & vbCrLf & "即将执行双向校验..." & vbCrLf & vbCrLf & _
              "点击【确定】开始核对", vbOKCancel + vbInformation, "阶段 2/3: 数据加载完成") = vbCancel Then
        Application.StatusBar = False
        Exit Sub
    End If
    Application.ScreenUpdating = False
    
    ' ==========================================================================
    ' 4. 生成校验矩阵 (逻辑不变)
    ' ==========================================================================
    Application.StatusBar = "正在进行比对分析..."
    
    Dim totalOutputCols As Long
    totalOutputCols = (endCol - startCol + 1) + 1
    ReDim arrOutput(1 To rngRowSelection.Cells.count, 1 To totalOutputCols)
    outRow = 0
    
    Dim cellSub As Range, currentRow As Long, currentTargetCol As Long
    
    For Each cellSub In rngRowSelection
        If cellSub.Font.Bold = True And Trim(cellSub.Value) <> "" Then
            outRow = outRow + 1
            strSubjectName = Trim(cellSub.Value)
            subjectKey = CleanKey(strSubjectName)
            currentRow = cellSub.Row
            
            arrOutput(outRow, 1) = strSubjectName
            outCol = 1
            
            For currentTargetCol = startCol To endCol
                outCol = outCol + 1
                
                Dim colHeader As String
                colHeader = Trim(CStr(wsTB.Cells(headerRow, currentTargetCol).Value))
                
                ' 跳过逻辑
                If InStr(colHeader, "审计") > 0 Or InStr(colHeader, "调整") > 0 Then
                    arrOutput(outRow, outCol) = ""
                Else
                    valTB = val(wsTB.Cells(currentRow, currentTargetCol).Value)
                    
                    If currentTargetCol = endCol Then
                        key = "TOTAL_SUM|" & subjectKey
                    ElseIf currentTargetCol = startCol And isFirstColParent Then
                        key = strParentFileMaster & "|" & subjectKey
                    Else
                        key = colHeader & "|" & subjectKey
                    End If
                    
                    cellText = ""
                    
                    ' 1. 获取基准值
                    valMasterTotalRow = 0
                    Dim hasTotal As Boolean
                    hasTotal = dictHasTotalRow.exists(key)
                    If hasTotal Then valMasterTotalRow = dictExplicitTotal(key)
                    
                    ' 2. 获取明细汇总
                    valMasterDetailSum = 0
                    If dictDetailSum.exists(key) Then valMasterDetailSum = dictDetailSum(key)
                    
                    ' 3. 双向比对
                    If hasTotal Then
                        diffInternal = Round(valMasterTotalRow - valMasterDetailSum, 2)
                        If Abs(diffInternal) > 0.01 Then
                            cellText = cellText & "源数据明细差:" & diffInternal & "; "
                        End If
                        
                        diffExternal = Round(valTB - valMasterTotalRow, 2)
                        If Abs(diffExternal) > 0.01 Then
                            cellText = cellText & "TB差异:" & diffExternal & "; "
                        End If
                    Else
                        diffExternal = Round(valTB - valMasterDetailSum, 2)
                        If Abs(diffExternal) > 0.01 Then
                            cellText = cellText & "TB差异(无合计行):" & diffExternal & "; "
                        End If
                    End If
                    
                    If cellText = "" Then
                        arrOutput(outRow, outCol) = "OK"
                    Else
                        If Right(cellText, 2) = "; " Then cellText = Left(cellText, Len(cellText) - 2)
                        arrOutput(outRow, outCol) = cellText
                    End If
                End If
            Next currentTargetCol
        End If
    Next cellSub
    
    ' --- 【断点 3】 诊断完成确认 ---
    Application.ScreenUpdating = True
    If MsgBox("核对计算已完成！" & vbCrLf & "即将生成报告。" & vbCrLf & vbCrLf & _
              "点击【确定】生成", vbOKCancel + vbInformation, "阶段 3/3: 计算完成") = vbCancel Then
        Application.StatusBar = False
        Exit Sub
    End If
    Application.ScreenUpdating = False
    
    ' ==========================================================================
    ' 5. 输出结果
    ' ==========================================================================
    Application.StatusBar = "正在生成报告..."
    
    If outRow = 0 Then MsgBox "未发现加粗科目！", vbExclamation: Exit Sub
    
    Dim reportName As String
    reportName = "诊断_" & Format(Now, "HHMMSS")
    Set wsReport = Worksheets.Add(After:=Worksheets(Worksheets.count))
    wsReport.Name = reportName
    
    ' 顶部说明
    With wsReport.Cells(1, 1)
        .Value = "【结果解读】 匹配规则[" & strTargetMetric & "] ■ 红色底纹: 源数据内部(合计行vs明细和)不平 | ■ 黄色底纹: TB数与Master合计行不符 | ■ 绿色文字: 校验通过"
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = RGB(80, 80, 80)
    End With
    wsReport.Range(wsReport.Cells(1, 1), wsReport.Cells(1, totalOutputCols)).Merge
    wsReport.Range(wsReport.Cells(1, 1), wsReport.Cells(1, totalOutputCols)).HorizontalAlignment = xlLeft
    
    ' 表头
    wsReport.Cells(2, 1).Value = "科目名称"
    Dim hIdx As Long: hIdx = 1
    For c = startCol To endCol
        hIdx = hIdx + 1
        wsReport.Cells(2, hIdx).Value = wsTB.Cells(headerRow, c).Value
    Next c
    
    ' 样式
    wsReport.Range(wsReport.Cells(2, 1), wsReport.Cells(2, totalOutputCols)).Interior.Color = RGB(55, 86, 115)
    wsReport.Range(wsReport.Cells(2, 1), wsReport.Cells(2, totalOutputCols)).Font.Color = vbWhite
    wsReport.Range(wsReport.Cells(2, 1), wsReport.Cells(2, totalOutputCols)).Font.Bold = True
    
    ' 数据
    wsReport.Range("A3").Resize(outRow, totalOutputCols).Value = arrOutput
    wsReport.Range(wsReport.Cells(2, 2), wsReport.Cells(outRow + 2, totalOutputCols)).HorizontalAlignment = xlCenter
    
    ' ==========================================================================
    ' 6. 智能上色
    ' ==========================================================================
    Application.StatusBar = "正在应用样式..."
    Dim rngCells As Range, cCell As Range
    Dim txtVal As String
    
    On Error Resume Next
    Set rngCells = wsReport.Range("B3").Resize(outRow, totalOutputCols - 1)
    On Error GoTo 0
    
    For Each cCell In rngCells
        txtVal = CStr(cCell.Value)
        If txtVal = "OK" Then
            cCell.Font.Color = RGB(0, 128, 0)
        ElseIf InStr(txtVal, "源数据") > 0 Then
            cCell.Interior.Color = RGB(255, 199, 206)
            cCell.Font.Color = RGB(156, 0, 6)
        ElseIf InStr(txtVal, "TB差异") > 0 Then
            cCell.Interior.Color = RGB(255, 235, 156)
            cCell.Font.Color = RGB(156, 87, 0)
        End If
    Next cCell
    
    wsReport.Columns.AutoFit
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    MsgBox "全流程结束！双重稽核报告已生成。", vbInformation, "完成"
End Sub

Private Sub AddVal(dict As Object, k As String, v As Double)
    If dict.exists(k) Then
        dict(k) = dict(k) + v
    Else
        dict(k) = v
    End If
End Sub

Private Function CleanKey(s As String) As String
    CleanKey = UCase(Trim(Replace(s, " ", "")))
End Function

