Attribute VB_Name = "Module2"
Sub RunScenarios()
    
    ' Change calculation to manua for faster macro
    ' execution and turn off screen updating for speed as well
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    ' Clear results of previous scenarios
    With Worksheets("s_res")
        .Activate
        .Range(Cells(10, 1), Cells(100000, 25)).ClearContents
   End With
   
   ' Start scenario counter to number each row of results
   scen_count = 0
   
    With Worksheets("s_def")
        .Activate
        
        ' Record input values pre-scenarios to be
        ' restored once scenarios are finished
        Dim orig_inp(0 To 9) As Double
        For i = 0 To 9 Step 1
            If .Cells(6 + i, 2) <> "" Then
                irange = .Cells(6 + i, 3) & "!" & .Cells(6 + i, 4)
                orig_inp(i) = Range(irange).Value
            End If
        Next i
       
        ' Loop over each input variable (10 loops)
        ' Skip if the variable is blank
        For one = .Cells(6, 5) To .Cells(6, 6) Step .Cells(6, 7)
            For two = .Cells(7, 5) To .Cells(7, 6) Step .Cells(7, 7)
                For three = .Cells(8, 5) To .Cells(8, 6) Step .Cells(8, 7)
                    For four = .Cells(9, 5) To .Cells(9, 6) Step .Cells(9, 7)
                        For five = .Cells(10, 5) To .Cells(10, 6) Step .Cells(10, 7)
                            For six = .Cells(11, 5) To .Cells(11, 6) Step .Cells(11, 7)
                                For seven = .Cells(12, 5) To .Cells(12, 6) Step .Cells(12, 7)
                                    For eight = .Cells(13, 5) To .Cells(13, 6) Step .Cells(13, 7)
                                        For nine = .Cells(14, 5) To .Cells(14, 6) Step .Cells(14, 7)
                                            For ten = .Cells(15, 5) To .Cells(15, 6) Step .Cells(15, 7)
                                                        'Debug.Print one & "," & two & "," & three & "," & four
                                          
                                                        ' Write inputs to cells
                                                        If .Cells(6, 2) <> "" Then
                                                            irange1 = .Cells(6, 3) & "!" & .Cells(6, 4)
                                                            Range(irange1).Value = one
                                                        End If
                                                        If .Cells(7, 2) <> "" Then
                                                            irange2 = .Cells(7, 3) & "!" & .Cells(7, 4)
                                                            Range(irange2).Value = two
                                                        End If
                                                        If .Cells(8, 2) <> "" Then
                                                            irange3 = .Cells(8, 3) & "!" & .Cells(8, 4)
                                                            Range(irange3).Value = three
                                                        End If
                                                        If .Cells(9, 2) <> "" Then
                                                            irange4 = .Cells(9, 3) & "!" & .Cells(9, 4)
                                                            Range(irange4).Value = four
                                                        End If
                                                        If .Cells(10, 2) <> "" Then
                                                            irange5 = .Cells(10, 3) & "!" & .Cells(10, 4)
                                                            Range(irange5).Value = five
                                                        End If
                                                        If .Cells(11, 2) <> "" Then
                                                            irange6 = .Cells(11, 3) & "!" & .Cells(11, 4)
                                                            Range(irange6).Value = six
                                                        End If
                                                        If .Cells(12, 2) <> "" Then
                                                            irange7 = .Cells(12, 3) & "!" & .Cells(12, 4)
                                                            Range(irange7).Value = seven
                                                        End If
                                                        If .Cells(13, 2) <> "" Then
                                                            irange8 = .Cells(13, 3) & "!" & .Cells(13, 4)
                                                            Range(irange8).Value = eight
                                                        End If
                                                        If .Cells(14, 2) <> "" Then
                                                            irange9 = .Cells(14, 3) & "!" & .Cells(14, 4)
                                                            Range(irange9).Value = nine
                                                        End If
                                                        If .Cells(15, 2) <> "" Then
                                                            irange10 = .Cells(15, 3) & "!" & .Cells(15, 4)
                                                            Range(irange10).Value = ten
                                                        End If
                
                                                        Calculate

                                                        ' Enter formulas to make certain that UDFs calculate
                                                        With Worksheets("s_res")
                                                            .Activate
                                                            .Cells(6, 2).Select
                                                            ActiveCell.FormulaR1C1 = "=IFERROR(befe(R[-2]C,R[-1]C),)"
                                                            ActiveCell.Copy
                                                            .Range("u6:c6").Select
                                                            Selection.PasteSpecial Paste:=xlPasteFormulas
                                                        End With
                                                        Calculate

                                                        ' Copy and paste input and output results for each scenario
                                                        With Worksheets("s_res")
                                                            .Activate
                                                            Range(.Cells(6, 2), .Cells(6, 21)).Copy
                                                            .Cells(scen_count + 10, 2).Select
                                                            Selection.PasteSpecial Paste:=xlPasteValues
                                                            .Cells(scen_count + 10, 1).Value = scen_count + 1
                                                         End With
                                                         Calculate

                                                        ' Put origininal input values (pre-scenarios) back
                                                        ' into their respective cells
                                                        For j = 0 To 9 Step 1
                                                            With Worksheets("s_def")
                                                                .Activate
                                                                If .Cells(6 + j, 2).Value <> "" Then
                                                                    irange = .Cells(6 + j, 3) & "!" & .Cells(6 + j, 4)
                                                                    Range(irange).Value = orig_inp(j)
                                                                End If
                                                            End With
                                                        Next j
                                                        Calculate
                                                        scen_count = scen_count + 1
                                                                                
                                                If .Cells(15, 2) = "" Then
                                                    Exit For
                                                End If
                                                Next ten
                                            If .Cells(14, 2) = "" Then
                                                Exit For
                                            End If
                                            Next nine
                                        If .Cells(13, 2) = "" Then
                                            Exit For
                                        End If
                                        Next eight
                                    If .Cells(12, 2) = "" Then
                                        Exit For
                                    End If
                                    Next seven
                                If .Cells(11, 2) = "" Then
                                    Exit For
                                End If
                                Next six
                            If .Cells(10, 2) = "" Then
                                Exit For
                            End If
                            Next five
                        If .Cells(9, 2) = "" Then
                            Exit For
                        End If
                        Next four
                    If .Cells(8, 2) = "" Then
                        Exit For
                    End If
                    Next three
                If .Cells(7, 2) = "" Then
                    Exit For
                End If
                Next two
           If .Cells(6, 2) = "" Then
                Exit For
            End If
           Next one
                                    
    End With

' Put calculation back to automatic
' and allow screen to efresh

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

End Sub

Function befe(sheet, cell)
    
    If sheet <> 0 And cell <> 0 Then
        befe = Range(sheet & "!" & cell).Value
    Else
        GoTo errormsg
    End If

errormsg:
    'befe = CVErr(xlErrValue)

End Function
