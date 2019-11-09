' ******************************************************************************************************************
' MODULE_NAME: mcParseVBARequireHeaders
' MODULE_VERSION: 2019-11-09
' MODULE_DESCRIPTION: Parses a BAS module's code (provided as a string) returns an instance of clsVBARequireHeader
' populated with the information.
' MODULE_HISTORY:
' 2019-11-09: Just getting something on the board. This is a first attempt and has not been tested. The class
' referenced here has also not yet been created.
' ******************************************************************************************************************
Option Explicit

Public Function ParseVBARequireHeaders(ByVal vbaModuleCode As String) As clsVBARequireHeader

	Dim result As New clsVBARequireHeader
	Dim lines() As String
	Dim line As String
	Dim r As Long
	Dim i As Long
	Dim currentHeader As String
	lines = Split(vbaModuleCode, vbCRLF)

	For r = LBound(lines) to UBound(lines)
		Do
			line = lines(r)
			If Left$(line, 1) = "'" Then
				' Prepare line for further processing
				line = Trim$(Mid$(line, 2))
			ElseIf UCase$(Left$(line, 4)) <> "REM " Then 
				' Prepare line for further processing
				line = Trim$(Mid$(line, 5))
			Else
				' Reached the end of the header (header must be at the top)
				Exit For
			End If
			If Left$(line, 7) = "MODULE_" Then
				line = Mid$(line, 8)
				i = Instr(line, ":")
				If i > 1 Then
					' Contains the name of the header directive
					currentHeader = Left$(line, i)
					' Line now contains the raw value
					line = LTrim$(Mid$(line, i+1))
					' Process the header line, and preserve header name
					' continuation on the next line if allowed
					If Not ProcessVBARequireHeader(result, currentHeader, line) Then currentHeader = ""
				Else
					' Line does not conform to the spec, ignore it
					Exit Do
				End If
			ElseIf currentHeader = "" Then
				' Not a VBA-Require header and not continuing a previous header,
				' ignore this line and continue processing with the next line
				Exit Do
			ElseIf line = "" Or Left$(line, 3) = "***" OR Left$(line, 3) = "---" Then
				' Blank or divider comment, don't allow the next line to add to the
				' previous value.
				currentHeader = ""
			End If
		Loop While False
	Next

	Set ParseVBARequireHeaders = result

End Function

' Returns True if the header was processed successfully AND the header supports multi-line values.
Private Function ProcessVBARequireHeader(ByRef result As clsVBARequireHeader, ByVal currentHeader As String, ByVal value As String) As Boolean
	Dim result As Boolean
	result = True
	Select Case currentHeader

		' Headers that only support a single value will overwrite with each new line
		Case "VERSION": result.Version = value: result = False
		Case "URL": result.URL = value: result = False
		Case "NAME": result.Name = value: result = False
		Case "HOMEPAGE": result.HomePage = value: result = False
		case "LICENSE": result.License = value: result = False

		' These header may have multi-line values, but as a single value. The word-wrapping locations
		' are preserved.
		Case "COPYRIGHT": AddStringLine(result.Copyright, value)
		Case "COMPATIBILITY": AddStringLine(result.Compatibility, value)
		Case "AUTHOR": AddStringLine(result.Author, value)
		Case "DESCRIPTION": AddStringLine(result.Description, value)
		Case "NOTES": AddStringLine(result.Notes, value)
		Case "USAGE": AddStringLine(result.Usage, value)

		' These headers should always have a single line with a comma-delimited set of values
		' FUTURE: Perhaps support multiple lines so long lists can be word-wrapped?
		Case "SCOPE_METHODS_NEEDED": result.ScopeMethodsNeeded = SplitCommaDelimitedString(value): result = False
		Case "SCOPE_VARIABLES_NEEDED": result.ScopeVariablesNeeded = SplitCommaDelimitedString(value): result = False
		Case "SCOPE_RANGES_NEEDED": result.ScopeRangesNeeded = SplitCommaDelimitedString(value): result = False

		' This should be stored as a 2D array, where the first dimension is 0=URL, 1=optional version,
		' and the second is the index for the dependency. Each dependency must have its own line.
		' This will require ReDim Preserve, which can only change the *rightmost* dimension.
		Case "DEPENDENCY":
			' TBD
		
		' This should be stored as a 2D array, where the first dimension is 0=version number, 1=note,
		' 2=URL, and the second dimension is the index for the version. See above re: ReDim Preserve.
		' A new version must *start* on its own line, but may have additional lines that should be added
		' to the previous Note. Version numbers should be validated as either x.y... or yyyy-mm-dd.
		Case "HISTORY"
			' TBD

		Case Else
			result = False

	End Select

End Function

Private Sub AddStringLine(ByRef s As String, ByVal value As String)
	If Len(s) > 0 Then s = s & vbCRLF
	s = s & value
End Sub

Private Function SplitCommaDelimitedString(ByVal s As String) As String()
	Dim result() As String
	Dim i As Long
	result = Split(s, ",")
	For i = LBound(result) To UBound(result)
		result(i) = Trim$(result)
	Next
	SplitCommaDelimitedString = result
End Function
