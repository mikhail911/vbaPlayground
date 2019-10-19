'#Language "WWB-COM"

'''''''''''''''''''''''''''''''''''''
'									'
'	E-Field Numerical Integration	'
'	  (c) 2019 by Michal Stojke    	'
'          license: MIT				'
'		   version: 0.9		     	'
'									'
'''''''''''''''''''''''''''''''''''''

Option Explicit

Sub Main
	Debug.Clear
	Dim Methods$(2), Directions$(3), Points$(4)
	Dim IntegrationStart As Double, IntegrationEnd As Double, IntegrationPoints As Double
	Dim IntegrationMethod As String, IntegrationDirection As String, drct As String, mthd As String
	Dim DebugSuffix As String
	Dim DebugCheckBox As Integer, DebugSuffixBox As Integer
	Dim nxmax As Long, nymax As Long, nzmax As Long, nxmin As Long, nymin As Long, nzmin As Long
	Dim datafile As String

	' Check maximum available values for x, y, z
	nxmax = Mesh.GetX(Mesh.GetNx-1)
    nymax = Mesh.GetY(Mesh.GetNy-1)
    nzmax = Mesh.GetZ(Mesh.GetNz-1)
	nxmin = Mesh.GetX(0)
	nymin = Mesh.GetY(0)
	nzmin = Mesh.GetZ(0)

	' Define available options
	Methods$(1) = "Trapezoidal rule"
	Methods$(2) = "Riemann sum (midpoint rule)"
	Directions$(1) = "x"
	Directions$(2) = "y"
	Directions$(3) = "z"
	Points$(1) = "1000"
	Points$(2) = "2000"
	Points$(3) = "500"
	Points$(4) = "100"

	' Define default values
	IntegrationStart = 0
	IntegrationEnd = 1000
	IntegrationPoints = 1000

	Begin Dialog UserDialog 630,322 ' %GRID:10,7,1,1
		Text 10,14,610,18,"E-Field Numerical Integration",.Text1,2
		Text 10,287,270,14,"(C) 2019 by Michal Stojke",.Text2,2
		GroupBox 10,49,610,224,"Options",.GroupBox1
		Text 30,70,200,14,"Method:",.Text3,2
		Text 30,100,200,14,"Direction of integration:",.Text4,2
		Text 30,130,200,14,"Start:",.Text5,2
		Text 30,160,200,14,"End:",.Text6,2
		Text 30,190,200,14,"Points:",.Text7,2
		PushButton 380,280,90,21,"&Calculate"
		CancelButton 500,280,90,21

		' User imput section
		TextBox 330,130,250,16,.IntegrationStart
		TextBox 330,160,250,16,.IntegrationEnd
		DropListBox 330,70,250,16,Methods(),.IntegrationMethod
		DropListBox 330,100,250,16,Directions(),.IntegrationDirection
		TextBox 330,190,250,16,.IntegrationPoints
		CheckBox 30+30,220,200,14,"Create debug.txt file",.DebugCheckBox
		CheckBox 340+20,220,250,14,"Include suffix to debug.txt",.DebugSuffixBox
		Text 30,250,200,14,"Suffix:",.Text8,2
		TextBox 330,250,250,16,.DebugSuffix
	End Dialog

	Dim dlg As UserDialog
	dlg.IntegrationStart = CStr(IntegrationStart)
	dlg.IntegrationEnd = CStr(IntegrationEnd)
	dlg.IntegrationPoints = CStr(IntegrationPoints)
	dlg.DebugSuffix = CStr(DebugSuffix)
	If (Dialog(dlg) = 0) Then Exit All

	IntegrationStart = Evaluate(dlg.IntegrationStart)
	IntegrationEnd = Evaluate(dlg.IntegrationEnd)
	IntegrationPoints = Evaluate(dlg.IntegrationPoints)
	IntegrationMethod = Evaluate(dlg.IntegrationMethod)
	IntegrationDirection = Evaluate(dlg.IntegrationDirection)
	DebugCheckBox = Evaluate(dlg.DebugCheckBox)
	DebugSuffix = dlg.DebugSuffix
	DebugSuffixBox = dlg.DebugSuffixBox
	Debug.Print(DebugSuffix)

	' Create and open file for debugging
	' During testing may occur some errors, but when exported to global macros works fine
	If Len(DebugSuffix) > 0 And DebugSuffixBox = 1 Then
		datafile = "debug_"+DebugSuffix+".txt"
		Open datafile For Output As #1
	ElseIf DebugCheckBox = 1 Then
		datafile = "debug.txt"
		Open datafile For Output As #1
	End If

	' Check for wierd data at input
	If IntegrationDirection = CStr(0) Then
		If IntegrationStart < nxmin Or IntegrationEnd > nxmax Then
			MsgBox "Start or end value out of range"
			Exit All
		End If
	ElseIf IntegrationDirection = CStr(1) Then
		If IntegrationStart < nymin Or IntegrationEnd > nymax Then
			MsgBox "Start or end value out of range"
			Exit All
		End If
	Else
		If IntegrationStart < nzmin Or IntegrationEnd > nzmax Then
			MsgBox "Start or end value out of range"
			Exit All
		End If
	End If
	If IntegrationEnd < IntegrationStart Then
		MsgBox "End value can't be smaller than start value, sorry."
		Exit All
	End If

	Dim i As Double

	If IntegrationMethod = CStr(0) Then
		mthd = "Trapezoidal rule"
		Dim vxre As Double, vyre As Double, vzre As Double, vxim As Double, vyim As Double, vzim As Double
		Dim s As Double, sIm As Double, h As Double, f As Double, f1 As Double, f2 As Double, yp As Double, zp As Double
		Dim xi As Double, yi As Double, f11 As Double, f22 As Double

		' Debug file header
		If DebugCheckBox = 1 Then Print #1, "x"; vbTab; vbTab; "y"; vbTab; vbTab; "z"; vbTab; vbTab; "VxRe"; vbTab; vbTab; "VyRe"; vbTab; vbTab; _
		          							"VzRe"; vbTab; vbTab; "VxIm"; vbTab; vbTab; "VyIm"; vbTab; vbTab; "VzIm"; vbCrLf


		If IntegrationDirection = CStr(0) Then
			drct = "X"

			yp = 0
			zp = 0

			'SelectTreeItem "2D/3D Results\E-Field\"
			f = GetFieldVector(IntegrationStart, yp, zp, vxre, vyre, vzre, vxim, vyim, vzim)
			f1 = vxre
			f11 = vxim
			f = GetFieldVector(IntegrationEnd, yp, zp, vxre, vyre, vzre, vxim, vyim, vzim)
			f2 = vxre
			f22 = vxim
			s = (f1 - f2)/2
			sIm = (f1 - f2)/2
			h = (IntegrationEnd - IntegrationStart)/IntegrationPoints

			For i = IntegrationStart To IntegrationEnd
				xi = IntegrationStart + i * h
				yi = GetFieldVector(xi, yp, zp, vxre, vyre, vzre, vxim, vyim, vzim)
				If DebugCheckBox = 1 Then Print #1, CStr(Format(xi,"0.00E+00")); vbTab ; CStr(Format(yp,"0.00E+00")); vbTab; CStr(Format(zp,"0.00E+00")); _
				          							vbTab; CStr(Format(vxre,"0.00E+00")); vbTab; CStr(Format(vyre,"0.00E+00")); vbTab; _
						 							CStr(Format(vzre,"0.00E+00")); vbTab; CStr(Format(vxim,"0.00E+00")); vbTab; _
						  							CStr(Format(vyim,"0.00E+00")); vbTab; CStr(Format(vzim,"0.00E+00"))
				s = s + vxre
				sIm = sIm + vxim
			Next
			s = h * s * Units.GetGeometryUnitToSI()
			sIm = h * sIm * Units.GetGeometryUnitToSI()
		ElseIf IntegrationDirection = CStr(1) Then
			drct = "Y"
		Else
			drct = "Z"
		End If

	ElseIf IntegrationMethod = CStr(1) Then
		mthd = "Riemann sum (midpoint rule)"
		Dim delta As Double

		yp = 0
		zp = 0

		If IntegrationDirection = CStr(0) Then
			drct = "X"

			s = 0
			sIm = 0
			For i = IntegrationStart To IntegrationEnd
				delta = i + (((i+1)-i)/2)
				f = GetFieldVector(delta, yp, zp, vxre, vyre, vzre, vxim, vyim, vzim)
				If DebugCheckBox = 1 Then Print #1, CStr(Format(delta,"0.00E+00")); vbTab ; CStr(Format(yp,"0.00E+00")); vbTab; CStr(Format(zp,"0.00E+00")); _
				          							vbTab; CStr(Format(vxre,"0.00E+00")); vbTab; CStr(Format(vyre,"0.00E+00")); vbTab; _
						 							CStr(Format(vzre,"0.00E+00")); vbTab; CStr(Format(vxim,"0.00E+00")); vbTab; _
						  							CStr(Format(vyim,"0.00E+00")); vbTab; CStr(Format(vzim,"0.00E+00"))
				s = s + vxre *((i+1)-i) * Units.GetGeometryUnitToSI()
				sIm = sIm + vxim *((i+1)-i) * Units.GetGeometryUnitToSI()
			Next
		ElseIf IntegrationDirection = CStr(1) Then
			drct = "Y"
		Else
			drct = "Z"
		End If
	End If

	If DebugCheckBox = 1 Then Print #1, vbCrLf; "Results: Re: "; s; " Im:"; sIm
	Close #1
	If DebugCheckBox = 1 Then Shell("notepad.exe " + datafile, 1)
	MsgBox vbCrLf + mthd + " integration from " + CStr(IntegrationStart) + " to " + CStr(IntegrationEnd) +  _
		   " in " + drct + " direction, with " + CStr(IntegrationPoints) + " points." + vbCrLf + _
		   vbCrLf + "Results:" + vbCrLf + vbCrLf + "V Real: " + CStr(s) + " [V]" + vbCrLf + _
		  "V Imag: " + CStr(sIm) + " [V]"
End Sub
