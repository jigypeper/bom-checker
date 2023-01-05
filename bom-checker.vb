Imports System.IO

' initialize empty list for data
Dim estr As New List(Of String)

Dim oFileDlg As Inventor.FileDialog = Nothing
            InventorVb.Application.CreateFileDialog(oFileDlg)
            oFileDlg.InitialDirectory = "C:\MCAD\Workspace\" 'OLC\"
            oFileDlg.Filter = "CSV Files (*.csv)|*.csv"
            oFileDlg.DialogTitle = "Select a BOM Sheet"
            oFileDlg.CancelError = True
            'On Error Resume Next
            oFileDlg.ShowOpen()
            If Err.Number <> 0 Then
                'exit if file not selected
                Return
            ElseIf oFileDlg.FileName <> "" Then
                Dim myCSV As String = oFileDlg.FileName
				
                Dim oData As New Dictionary(Of String, String)
                Dim oDataM As New Dictionary(Of Integer, Integer)
                Dim oRow As Integer
               
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

                ' Get the active assembly.
                Dim oAsmDoc As AssemblyDocument
                oAsmDoc = ThisApplication.ActiveDocument


                Dim nrOfOccs As Integer = 0


                ' Get the occurrences that represent this document.
                Dim oOccs As ComponentOccurrencesEnumerator

                ' Print the occurrences to the Immediate window.
                Dim oOcc As ComponentOccurrence


                For Each kvp As KeyValuePair(Of String, String) In oData

                    For Each oDocCheck As Document In oAsmDoc.AllReferencedDocuments
                        Try
                            If oDocCheck.PropertySets.Item("Design Tracking Properties").Item("Stock Number").Value = kvp.Key Then
                                nrOfOccs = oAsmDoc.ComponentDefinition.Occurrences.AllReferencedOccurrences(oDocCheck).Count
                            End If
                        Catch
                        End Try
                    Next
                    Dim x2 As Integer = kvp.Key
                    Dim y2 As Integer = nrOfOccs
                    oDataM.Add(x2, y2)

                Next


                Dim assemblyDef As AssemblyComponentDefinition = ThisDoc.Document.ComponentDefinition

                Try


                    ' Activate a writeable View Rep (master view rep is not writeable)

                    assemblyDef.RepresentationsManager.DesignViewRepresentations.Item("BOM Check").Activate()

                Catch

                    ' Assume error means this View Rep does not exist, so create it

                    oViewRep = assemblyDef.RepresentationsManager.DesignViewRepresentations.Add("BOM Check")
                    oViewRep.Activate

                End Try
                
                ' turn off work features in the model
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

                ' turn off visibility if quantities in dictionaries are equal (part level)
                For Each occ In assemblyDef.Occurrences.AllLeafOccurrences
                    Dim refDoc As PartDocument = occ.Definition.Document

                    customPropSet = refDoc.PropertySets.Item("Design Tracking Properties")

                    SAPProp = customPropSet.Item("Stock Number")

                    SAPVal = SAPProp.Value

                    If SAPVal <> Nothing And IsNumeric(SAPVal) Then

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
                ' turn off visibility if quantities in dictionaries are equal (assembly level)
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

				' try to activate level of detail
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
			
            ' create list of results, remove duplicates
			Dim result As List(Of String) = estr.Distinct().ToList
			
            ' concatenate results to variable for msgbox display
			For i As Integer = 0 To result.Count - 1
				If i = 0 Then
					oText = result(i)
				Else
					oText = oText & vbLf & result(i)
				End If
			Next
			
			' Create and write to a text file
			oWrite = System.IO.File.CreateText(ThisDoc.PathAndFileName(False) & ".txt")

			' Write list out to file
			oWrite.WriteLine("SAP No: EBOM QTY <> BOM QTY:")
			For Each Item As String In result
				oWrite.WriteLine(Item)
			Next

            ' close the file
			oWrite.Close()
			
            ' display data to user
			MsgBox("The following Items have inconsistent quantities:" & vbCrLf & "SAP No: EBOM QTY <> BOM QTY:" & vbCrLf & oText, ,"BOM Check")

			' open the file
			ThisDoc.Launch(ThisDoc.PathAndFileName(False) & ".txt")