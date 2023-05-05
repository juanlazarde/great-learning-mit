Option Explicit

Sub ImportExcelData()
    ' ImportExcelData Macro
    ' Description: Reads each row of an Excel file, using column B as the headline and column C as the content, and inserts a new heading and content into the Word document. First row has headers.
    
    On Error GoTo ErrorHandler
    
    ' Declare variables
    Dim ScreenUpdating As Boolean
    Dim ExcelApp As Object
    Dim ExcelWorkbook As Object
    Dim ExcelWorksheet As Object
    Dim RowIndex As Long
    Dim Rows As Long
    Dim Heading As String
    Dim Content As String
    Dim ExcelFilePath As String
    
    ' Excel file to import. If "", then it will open a dialog box to find it.
    ExcelFilePath = ""

    ' Optimize for speed
    ScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False
    
    ' Open file dialog box to select Excel file
    If ExcelFilePath = "" Then
        With Application.FileDialog(msoFileDialogFilePicker)
            .Title = "Select Excel File with FAQ"
            .Filters.Clear
            .Filters.Add "Excel Files", "*.xlsx", 1
            .Filters.Add "CSV Files", "*.csv", 2
            If .Show = -1 Then ExcelFilePath = .SelectedItems(1) Else Exit Sub
        End With
    End If
    
    ' Open the Excel file
    Set ExcelApp = CreateObject("Excel.Application")
    Set ExcelWorkbook = ExcelApp.Workbooks.Open(ExcelFilePath, ReadOnly:=True, UpdateLinks:=False, IgnoreReadOnlyRecommended:=True)
    
    ' Loop through each sheet in the Excel file
    For Each ExcelWorksheet In ExcelWorkbook.Worksheets
        ' Find used cells in the sheet
        For Rows = UBound(ExcelWorksheet.UsedRange.Formula) To 1 Step -1
            If Trim(ExcelWorksheet.Cells(Rows, 1).Value) <> "" Then Exit For
        Next Rows
        
        ' Loop through each row in the sheet
        For RowIndex = 2 To Rows
            ' Read the heading and content from the row
            Heading = ExcelWorksheet.Cells(RowIndex, 2).Value
            Content = ExcelWorksheet.Cells(RowIndex, 3).Value
            
            ' Replace any "." with ". ", "," with ", ", and any double spaces or more in the text
            Content = Replace(Content, ".", ". ")
            Content = Replace(Content, ",", ", ")
            Content = Replace(Content, "  ", " ")
            Content = Replace(Content, "  ", " ")
            
            ' Insert the heading and content into the Word document
            With Selection
                .Style = ActiveDocument.Styles("Heading 1")
                .TypeText Text:=Heading
                .TypeParagraph
                .Style = ActiveDocument.Styles("Normal")
                .TypeText Text:=Content
                .TypeParagraph
            End With
        Next RowIndex
    Next ExcelWorksheet
    
    ' Clean up and close the Excel file
    ExcelWorkbook.Close SaveChanges:=False
    ExcelApp.Quit
    Set ExcelWorksheet = Nothing
    Set ExcelWorkbook = Nothing
    Set ExcelApp = Nothing
    
    ' Insert Table of Contents or Update it if it exists
    Dim TOC As TablesOfContents 'Object
    Set TOC = ActiveDocument.TablesOfContents
    
    If TOC.Count = 0 Then _
        TOC.Add _
            Range:=ActiveDocument.Range(0, 0), _
            UseFields:=False, _
            UseHeadingStyles:=True, _
            LowerHeadingLevel:=3, _
            UpperHeadingLevel:=1, _
            IncludePageNumbers:=False, _
            UseHyperlinks:=True
    TOC(1).Update
    Set TOC = Nothing
    TocFormatting
        
    ' Add a closing message
    Selection.TypeText Text:="Have fun learning!"
    
    ' Restore settings
    Application.ScreenUpdating = ScreenUpdating
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    If Not ExcelWorkbook Is Nothing Then ExcelWorkbook.Close SaveChanges:=False
    If Not ExcelApp Is Nothing Then ExcelApp.Quit
    Set ExcelWorksheet = Nothing
    Set ExcelWorkbook = Nothing
    Set ExcelApp = Nothing
    Set TOC = Nothing
    Application.ScreenUpdating = ScreenUpdating
    
End Sub

Private Sub TocFormatting()
    With ActiveDocument.Styles("TOC 1")
        .AutomaticallyUpdate = True
        .BaseStyle = "Normal"
        .NextParagraphStyle = "Normal"
        .LinkToListTemplate ListTemplate:= _
            ListGalleries(wdNumberGallery).ListTemplates(1), ListLevelNumber:=1
        With .Font
            .Name = "+Body"
            .Size = 6
        End With
        With .ParagraphFormat
            .LeftIndent = InchesToPoints(0)
            .RightIndent = InchesToPoints(0)
            .SpaceBefore = 0
            .SpaceBeforeAuto = False
            .SpaceAfter = 0
            .SpaceAfterAuto = False
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphLeft
            .LineUnitBefore = 0
            .LineUnitAfter = 0
        End With
    End With
    With ListGalleries(wdNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "%1."
        .TrailingCharacter = wdTrailingTab
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = InchesToPoints(0.25)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = InchesToPoints(0.5)
        .TabPosition = wdUndefined
        .ResetOnHigher = 0
        .StartAt = 1
        .Font.Size = 6
        .LinkedStyle = "TOC 1"
    End With
End Sub

Private Sub Document_Close()
    Setup Uninstall:=True
End Sub

Private Sub Document_Open()
    Setup
End Sub

Sub Setup(Optional ByVal Uninstall As Boolean = False)
    ' SetupButton Macro
    ' Description: Adds a button to Word's Ribbon to run a specified macro, if the button doesn't already exist.
    
    ' Set the name of the macro to run when the button is clicked
    Const RibbonName = "MachineLearning"
    Const MacroName = "ImportExcelData"
    Const Caption = "FAQ Excel to Word"
    Const TooltipText = "Import AI/ML No Code Excel FAQ and convert to Word"
    
    On Error GoTo ErrorHandler
    
    ' Get or create the Ribbon command bar and make it visible
    Dim Ribbon As CommandBar
    For Each Ribbon In Application.CommandBars
        If Ribbon.Type = msoBarTypeNormal And Ribbon.Name = RibbonName Then Exit For
    Next Ribbon
    If Ribbon Is Nothing Then Set Ribbon = CommandBars.Add(RibbonName)
    If Not Ribbon.Visible Then Ribbon.Visible = True
    
    ' Delete the Ribbon and exit if the Uninstall flag is set
    If Uninstall Then
        If Not Ribbon Is Nothing Then Ribbon.Delete 'CommandBars(RibbonName).Delete
        Exit Sub
    End If
    
    ' Check if the button already exists on the Ribbon
    Dim RibbonButton As CommandBarButton
    For Each RibbonButton In Ribbon.Controls
        If RibbonButton.OnAction = MacroName And TypeOf RibbonButton Is CommandBarButton Then
            MsgBox "Button already exists on the Ribbon." & vbCr & "Delete the button, the ribbon, or ignore this message."
            Exit Sub
        End If
    Next RibbonButton
    
    ' If the button doesn't already exist, add it to the Ribbon
    Set RibbonButton = Ribbon.Controls.Add(Type:=msoControlButton, Before:=1)
    With RibbonButton
        .Caption = Caption
        .OnAction = MacroName
        .Style = msoButtonCaption
        .TooltipText = TooltipText
        .DescriptionText = "Description text"
    End With
    
    ' Display a message box to indicate that the button was added to the Ribbon
    MsgBox Caption & " button added to the Add-ins menu."
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical, "Error"
    
End Sub
