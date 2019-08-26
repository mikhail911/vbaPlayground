' Generate E-Fields from 0 To 100 MHz

Sub Main ()
	For i = 0 To 100 STEP 1
		' Delete previous results
		Monitor.Delete "e-field (f=" & i & ")"

		' Create field monitor for desired frequencies
		With Monitor
	     .Reset
	     .Name "e-field (f=" & i & ")"
	     .Dimension "Volume"
	     .Domain "Frequency"
	     .FieldType "Efield"
	     .MonitorValue "" & i & ""
	     .UseSubvolume "False"
	     .Coordinates "Structure"
	     .SetSubvolume "0", "3.0e+04", "-1.5", "1.5", "-1.5", "203"
	     .SetSubvolumeOffset "0.0", "0.0", "0.0", "0.0", "0.0", "0.0"
	     .Create
	     End With
	Next[i]

	' Calculate E-Fields
	Solver.Start
	Dim dIntReal As Double, dIntImag As Double
	For i = 0 To 100 STEP 1
		EvaluateFieldAlongCurve.FitCurveToGridForPlot(False)
		EvaluateFieldAlongCurve.FitCurveToGridForIntegration(True)
		EvaluateFieldAlongCurve.EvaluateOnSurface(False)

		SelectTreeItem("2D/3D Results\E-Field\e-field (f=" & i & ") [pw]")

		EvaluateFieldAlongCurve.IntegrateField("wire1", "x", dIntReal, dIntImag)
		EvaluateFieldAlongCurve.PlotField("wire1", "x")
	Next
End Sub
