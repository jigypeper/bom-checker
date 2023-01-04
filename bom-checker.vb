AddReference "Microsoft.Office.Interop.Excel.dll"
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Dim estr As New List(Of String)

Dim oFileDlg As Inventor.FileDialog = Nothing
            InventorVb.Application.CreateFileDialog(oFileDlg)
            oFileDlg.InitialDirectory = "C:\MCAD\Workspace\" 'OLC\"
            oFileDlg.Filter = "CSV Files (*.csv)|*.csv"
            oFileDlg.DialogTitle = "Select a BOM Sheet"
            'oFileDlg.InitialDirectory = ThisDoc.Path
            oFileDlg.CancelError = True
            'On Error Resume Next
            oFileDlg.ShowOpen()
            If Err.Number <> 0 Then
                'exit if file not selected
                Return
            ElseIf oFileDlg.FileName <> "" Then
                Dim myCSV As String = oFileDlg.FileName
				'Dim filepath As String = Cstr(myCSV)
'                'define Excel Application object
'                excelApp = CreateObject("Excel.Application")
'                'workbook exists, open it
'                excelWorkbook = excelApp.Workbooks.Open(myXLS)
'                oSheetCount = excelWorkbook.Sheets.Count
'                mySheet = "BOM"

'                Dim oSheetName As String = "BOM"
'                Dim oExcelApp As New Microsoft.Office.Interop.Excel.ApplicationClass
'                oExcelApp.DisplayAlerts = False
'                Dim oWB As Workbook = oExcelApp.Workbooks.Open(myXLS)
'                Dim oWS As Worksheet = oWB.Sheets.Item(oSheetName)
'                Dim oLastRowUsed As Integer = oWS.UsedRange.Rows.Count
'                'MsgBox("Last row being used is Row#:  " & oLastRowUsed)
'                Dim oCells As Range = oWS.Cells
				'MsgBox(myCSV)
                Dim oData As New Dictionary(Of String, String)
                Dim oDataM As New Dictionary(Of Integer, Integer)
                Dim oRow As Integer
                'First number is Row index, second number is Column Index
'                For oRow = 2 To oLastRowUsed
'                    Dim x As Integer = GoExcel.CellValue(myXLS, "BOM", "A" & oRow)
'                    Dim y As Integer = CInt(GoExcel.CellValue(myXLS, "BOM", "B" & oRow))
'                    oData.Add(x, y)

'                Next
				
				' Open the CSV file
				'Dim reader As StreamReader = myCSV
				'reader(myCSV)
				Using reader As New StreamReader(myCSV)
				    ' Read the header line and discard it
				    reader.ReadLine()

				    ' Read the rest of the file
				    While Not reader.EndOfStream
				        ' Read a line of text
				        Dim Line As String = reader.ReadLine()

				        ' Split the line into fields
				        Dim fields As String() = Line.Split(","c)

				        ' Add the fields to the dictionary
				        oData.Add(fields(0), fields(1))
				    End While
				End Using

                'For Each kvp As KeyValuePair(Of String, Integer) In oData
                '  MsgBox(kvp.Key & " -- " & kvp.Value)
                'Next
				
				

                Dim oAsm As AssemblyDocument = ThisApplication.ActiveDocument
                Dim oAsmCompDef As AssemblyComponentDefinition = oAsm.ComponentDefinition
                Dim comps As New List(Of ComponentOccurrence)


                ' Get the active assembly.
                Dim oAsmDoc As AssemblyDocument
                oAsmDoc = ThisApplication.ActiveDocument

                ' Get the definition of the assembly.
                Dim oAsmDef As AssemblyComponentDefinition
                oAsmDef = oAsmDoc.ComponentDefinition

                Dim nrOfOccs As Integer = 0


                ' Get the occurrences that represent this document.
                Dim oOccs As ComponentOccurrencesEnumerator
                'oOccs = oAsmDef.Occurrences.AllReferencedOccurrences(oDoc)

                ' Print the occurrences to the Immediate window.
                Dim oOcc As ComponentOccurrence


                For Each kvp As KeyValuePair(Of String, String) In oData

                    For Each oDocCheck As Document In oAsmDoc.AllReferencedDocuments
                        Try
                            If oDocCheck.PropertySets.Item("Design Tracking Properties").Item("Stock Number").Value = kvp.Key Then
                                nrOfOccs = oAsmDoc.ComponentDefinition.Occurrences.AllReferencedOccurrences(oDocCheck).Count

                                For Each comp As ComponentOccurrence In oAsm.ComponentDefinition.Occurrences
                                    oFileName5 = comp.Definition.Document.displayname
                                    If InStr(oFileName5, kvp.Key) <> 0 Then
                                        comps.Add(comp)
                                    End If
                                Next
                            End If
                        Catch
                        End Try
                    Next
                    Dim x2 As Integer = kvp.Key
                    Dim y2 As Integer = nrOfOccs
                    oDataM.Add(x2, y2)
                    'MsgBox("Number of occurrences of: " & kvp.Key & " = " & nrOfOccs)
                Next

                'For Each kvp As KeyValuePair(Of Integer, Integer) In oDataM
                'MsgBox(kvp.Key & " -- " & kvp.Value)
                'Next	
                'MsgBox(comps.Count)


                Dim assemblyDef As AssemblyComponentDefinition = ThisDoc.Document.ComponentDefinition

                Try


                    'Activate a writeable View Rep (master view rep is not writeable)

                    assemblyDef.RepresentationsManager.DesignViewRepresentations.Item("BOM Check").Activate()

                Catch

                    'Assume error means this View Rep does not exist, so create it

                    oViewRep = assemblyDef.RepresentationsManager.DesignViewRepresentations.Add("BOM Check")
                    oViewRep.Activate

                End Try


'                If (oData.ContainsKey(30525649) And oDataM.ContainsKey(30525649)) And (oData(30525649) = oDataM(30525649)) Then
'                    MsgBox(oData(30525649) & " VS " & oDataM(30525649))
'                Else
'                    MsgBox(oData(30525649) & " <> " & oDataM(30525649))
'                End If

'                If (oData.ContainsKey(30444523) And oDataM.ContainsKey(30444523)) And (oData(30444523) = oDataM(30444523)) Then
'                    MsgBox(oData(30444523) & " VS " & oDataM(30444523))
'                Else
'                    MsgBox(oData(30444523) & " <> " & oDataM(30444523))
'                End If

                For Each oDocCheck As Document In oAsmDoc.AllReferencedDocuments
                    For Each oworkplane In assemblyDef.WorkPlanes
                        oworkplane.Visible = False
                    Next
                    For Each oworkaxis In assemblyDef.WorkAxes
                        oworkaxis.Visible = False
                    Next
                    For Each oworkpoint In assemblyDef.WorkPoints
                        oworkpoint.Visible = False
                    Next
                Next

                Dim occ As Inventor.ComponentOccurrence

                For Each occ In assemblyDef.Occurrences.AllLeafOccurrences
                    Dim refDoc As PartDocument = occ.Definition.Document

                    customPropSet = refDoc.PropertySets.Item("Design Tracking Properties")

                    SAPProp = customPropSet.Item("Stock Number")

                    SAPVal = SAPProp.Value

                    '		If SAPVal.ToString.Contains("30571635") Then
                    '			MsgBox("Exists")
                    '		End If

                    If SAPVal <> Nothing And IsNumeric(SAPVal) Then
                        'MsgBox(SAPVal)
                        If (oData.ContainsKey(SAPVal) And oDataM.ContainsKey(SAPVal)) Then
                            If (oData(SAPVal) = oDataM(SAPVal)) Then
                                occ.Visible = False
                            Else
                                occ.Visible = True
								estr.Add(CStr(SAPVal) & ": " & CStr(oData(SAPVal)) & " <> " & CStr(oDataM(SAPVal)))
                            End If
                        End If


                    End If



                Next

                For Each occ In assemblyDef.Occurrences
                    Dim oANAME As String = occ.Name
                    iPropSAP = iProperties.Value(oANAME, "Project", "Stock Number")

                    If iPropSAP <> Nothing And IsNumeric(iPropSAP) Then
                        If (oData.ContainsKey(iPropSAP) And oDataM.ContainsKey(iPropSAP)) Then
                            If (oData(iPropSAP) = oDataM(iPropSAP)) Then
                                occ.Visible = False
                            Else
                                occ.Visible = True
								estr.Add(CStr(iPropSAP) & ": " & CStr(oData(iPropSAP)) & " <> " & CStr(oDataM(iPropSAP))) 
                            End If
                        End If
                    End If

                Next

                'MsgBox(oData.Item(30571635) & "--" & oDataM.Item(30571635))
				
				'AsmComment = oAsmDoc.PropertySets.Item("Inventor Summary Information").Item("Comments").Value.ToString
				'AsmComment.Equals(""), case for all other assemblies minus legacy
				'AsmComment.Contains("GC Series"), case for GC Series
				'AsmComment.Contains("Legacy"), case for legacy systems (s80,s90, s200, etc.)
				
				Try
	                If oAsmDoc.PropertySets.Item("Inventor Summary Information").Item("Comments").Value.ToString.Contains("GC Series") Then
	                    assemblyDef.RepresentationsManager.LevelOfDetailRepresentations.Item("All Content Center Suppressed").Activate()
	                Else
	                    assemblyDef.RepresentationsManager.LevelOfDetailRepresentations.Item("iLogic").Activate()
	                End If
				Catch
					MsgBox("Not a GC or Legacy System", , "Bom Check")
				End Try
				
			End If
			
			Dim result As List(Of String) = estr.Distinct().ToList
			
			For i As Integer = 0 To result.Count - 1
				If i = 0 Then
					oText = result(i)
				Else
					oText = oText & vbLf & result(i)
				End If
			Next
			
			'MsgBox(oText)
			'Create and write to a text file
			oWrite = System.IO.File.CreateText(ThisDoc.PathAndFileName(False) & ".txt")

			'Write Array out to String
			oWrite.WriteLine("SAP No: EBOM QTY <> BOM QTY:")
			For Each Item As String In result
				oWrite.WriteLine(Item)
			Next

			oWrite.Close()

			'ThisDoc.Launch(ThisDoc.PathAndFileName(False) & ".txt")'open the file
			
			MsgBox("The following Items have inconsistent quantities:" & vbCrLf & "SAP No: EBOM QTY <> BOM QTY:" & vbCrLf & oText, ,"BOM Check")

                'close the workbook And the Excel Application
'                excelWorkbook.Close
'            excelApp.Quit
			
			ThisDoc.Launch(ThisDoc.PathAndFileName(False) & ".txt")'open the file