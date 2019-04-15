Attribute VB_Name = "M_SheetsVisible"
Option Explicit

Private Sub SheetsVisibleSample()
    'シート3枚のブックを用意。
    Dim wb As Excel.Workbook
    Set wb = Workbooks.Add(XlWBATemplate.xlWBATWorksheet)
    wb.Worksheets.Add
    wb.Worksheets.Add
    
    Stop
    
    '1, 3番目のシートの集合を取得。
    Dim shts As Excel.Sheets
    Set shts = wb.Worksheets.Item(VBA.Array(1, 3))
    
    '複数を通常非表示はOK(1, 3番目のシートが非表示になる)。
    shts.Visible = XlSheetVisibility.xlSheetHidden
    
    '以下はNG
    '表示したい...
    shts.Visible = XlSheetVisibility.xlSheetVisible
    '手動操作で表示できなくしたい...
    shts.Visible = XlSheetVisibility.xlSheetVeryHidden
    
End Sub
