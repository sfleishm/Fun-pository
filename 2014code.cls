VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub hw2easy2014():

  ' Set an initial variable for holding the brand name
  Dim Brand_Name As String

  ' Set an initial variable for holding volume
  
  Dim Brand_Total As Double
  Brand_Total = 0

  ' Keep track of each entry
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all everything
  For i = 2 To 705714

    ' Check if we are still within the same brand, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Brand name
      Brand_Name = Cells(i, 1).Value

      ' Add to the Brand Total
      Brand_Total = Brand_Total + Cells(i, 7).Value

      ' Print the Brand in the Summary Table
      Range("i" & Summary_Table_Row).Value = Brand_Name

      ' Print the Volume total to the Summary Table
      Range("j" & Summary_Table_Row).Value = Brand_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      Brand_Total = 0

    ' If the cell immediately following a row is the same brand...
    Else
      ' Add to the Brand Total
      Brand_Total = Brand_Total + Cells(i, 7).Value
    End If
  Next i
End Sub




