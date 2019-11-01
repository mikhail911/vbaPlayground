'#Language "WWB-COM"

Option Explicit

' Generate E-Fields from 0.8 To 18 GHz

Sub Main ()
	Debug.Clear
	Dim y, zval As Double
	Dim x, datafile, freq, zvals As String

	Dim STP As Double
	STP = 100

	' Delete existing calculations
	With Resulttree
		.Name "Farfields"
 		.Delete
	End With

	' Create file for output
	datafile = ""+GetProjectPath("Project")+"\phase_center_export.txt"
	Open datafile For Output As #1

	' Write header to file
	Print #1, "frequency [GHz]"; vbTab; vbTab; "z [mm]"
	Print #1, "---------------------------------------"

	For y = 800 To 18000 STEP STP
		x = Str(y / 1000) ' Convert MHz to GHz (desired for project)
		x = Replace(x, ",", ".", 1, -1)

		SelectTreeItem("Farfields\farfield (f=" & x & ") [1]")

		With FarfieldPlot
		     .Plottype ("3d")
		     .Step (5)
		     .SetColorByValue (True)
		     .DrawStepLines (False)
		     .DrawIsoLongitudeLatitudeLines (False)
		     .EnablePhaseCenterCalculation (True)
		     .SetPhaseCenterAngularLimit (30)
		     .SetPhaseCenterComponent ("boresight")
		     .SetPhaseCenterPlane ("e-plane")
		End With
		
		freq = Replace(Str(Format(y, "0.00E+00")), ",", ".", 1, -1)
		zval = FarfieldPlot.GetPhaseCenterResult("z", "eplane")
		zvals = Replace(Str(Format(zval, "0.00E+00")), ",", ".", 1, -1)
		Print #1, freq; vbTab; vbTab; zvals
	Next[y]
	Close #1
End Sub
