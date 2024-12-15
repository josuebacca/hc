Attribute VB_Name = "Module2"
' vcf1cmpt.txt
'
' Copyright (c) 1992-1995, Visual Components Inc.
'
'Option Explicit
Public Function SSAddColPageBreak%(SS As Object, ByVal nCol%)

   On Error GoTo SSAddColPageBreakError
   SS.AddColPageBreak nCol
   SSAddColPageBreak = 0
   Exit Function

SSAddColPageBreakError:
   SSAddColPageBreak = Err.Number

End Function

Public Function SSAddPageBreak%(SS As Object)

   On Error GoTo SSAddPageBreakError
   SS.AddPageBreak
   SSAddPageBreak = 0
   Exit Function

SSAddPageBreakError:
   SSAddPageBreak = Err.Number

End Function

Public Function SSAddRowPageBreak(SS As Object, ByVal nRow%)

   On Error GoTo SSAddRowPageBreakError
   SS.AddRowPageBreak nRow
   SSAddRowPageBreak = 0
   Exit Function

SSAddRowPageBreakError:
   SSAddRowPageBreak = Err.Number

End Function

Public Function SSAddSelection%(SS As Object, ByVal nR1%, ByVal nC1%, ByVal nR2%, ByVal nC2%)

   On Error GoTo SSAddSelectionError
   SS.AddSelection nR1, nC1, nR2, nC2
   SSAddSelection = 0
   Exit Function

SSAddSelectionError:
   SSAddSelection = Err.Number

End Function

Public Function SSAttach%(SS As Object, ByVal pTitle$)

   On Error GoTo SSAttachError
   SS.Attach pTitle
   SSAttach = 0
   Exit Function

SSAttachError:
   SSAttach = Err.Number

End Function

Public Function SSAttachToSS%(SS As Object, hSrcSS As Long)

   On Error GoTo SSAttachToSSError
   SS.AttachToSS hSrcSS
   SSAttachToSS = 0
   Exit Function

SSAttachToSSError:
   SSAttachToSS = Err.Number

End Function

Public Function SSCalculationDlg%(SS As Object)

   On Error GoTo SSCalculationDlgError
   SS.CalculationDlg
   SSCalculationDlg = 0
   Exit Function

SSCalculationDlgError:
   SSCalculationDlg = Err.Number

End Function

Public Function SSCancelEdit%(SS As Object)

   On Error GoTo SSCancelEditError
   SS.CancelEdit
   SSCancelEdit = 0
   Exit Function

SSCancelEditError:
   SSCancelEdit = Err.Number

End Function

Public Function SSEditPasteValues%(SS As Object)

   On Error GoTo SSEditPasteValuesError
   SS.EditPasteValues
   SSEditPasteValues = 0
   Exit Function

SSEditPasteValuesError:
   SSEditPasteValues = Err.Number

End Function

Public Function SSCheckRecalc%(SS As Object)

   On Error GoTo SSCheckRecalcError
   SS.CheckRecalc
   SSCheckRecalc = 0
   Exit Function

SSCheckRecalcError:
   SSCheckRecalc = Err.Number

End Function

Public Function SSClearClipboard%(SS As Object)

   On Error GoTo SSClearClipboardError
   SS.ClearClipboard
   SSClearClipboard = 0
   Exit Function

SSClearClipboardError:
   SSClearClipboard = Err.Number

End Function

Public Function SSClearRange%(SS As Object, ByVal nR1%, ByVal nC1%, ByVal nR2%, ByVal nC2%, ByVal nClearType%)

   On Error GoTo SSClearRangeError
   SS.ClearRange nR1, nC1, nR2, nC2, nClearType
   SSClearRange = 0
   Exit Function

SSClearRangeError:
   SSClearRange = Err.Number

End Function

Public Function SSColorPaletteDlg%(SS As Object)

   On Error GoTo SSColorPaletteDlgError
   SS.ColorPaletteDlg
   SSColorPaletteDlg = 0
   Exit Function

SSColorPaletteDlgError:
   SSColorPaletteDlg = Err.Number

End Function

Public Function SSColWidthDlg%(SS As Object)

   On Error GoTo SSColWidthDlgError
   SS.ColWidthDlg
   SSColWidthDlg = 0
   Exit Function

SSColWidthDlgError:
   SSColWidthDlg = Err.Number

End Function

Public Function SSCopyAll%(SS As Object, ByVal hSrcSS&)

   On Error GoTo SSCopyAllError
   SS.CopyAll hSrcSS
   SSCopyAll = 0
   Exit Function

SSCopyAllError:
   SSCopyAll = Err.Number

End Function

Public Function SSCopyRange%(SS As Object, ByVal nDstR1%, ByVal nDstC1%, ByVal nDstR2%, ByVal nDstC2%, ByVal hSrcSS&, ByVal nSrcR1%, ByVal nSrcC1%, ByVal nSrcR2%, ByVal nSrcC2%)

   On Error GoTo SSCopyRangeError
   SS.CopyRange nDstR1, nDstC1, nDstR2, nDstC2, hSrcSS, nSrcR1, nSrcC1, nSrcR2, nSrcC2
   SSCopyRange = 0
   Exit Function

SSCopyRangeError:
   SSCopyRange = Err.Number

End Function

Public Function SSDefinedNameDlg%(SS As Object)

   On Error GoTo SSDefinedNameDlgError
   SS.DefinedNameDlg
   SSDefinedNameDlg = 0
   Exit Function

SSDefinedNameDlgError:
   SSDefinedNameDlg = Err.Number

End Function

Public Function SSDeleteDefinedName%(SS As Object, ByVal pName$)

   On Error GoTo SSDeleteDefinedNameError
   SS.DeleteDefinedName pName
   SSDeleteDefinedName = 0
   Exit Function

SSDeleteDefinedNameError:
   SSDeleteDefinedName = Err.Number

End Function

Public Function SSDeleteRange%(SS As Object, ByVal nR1%, ByVal nC1%, ByVal nR2%, ByVal nC2%, ByVal nShiftType%)

   On Error GoTo SSDeleteRangeError
   SS.DeleteRange nR1, nC1, nR2, nC2, nShiftType
   SSDeleteRange = 0
   Exit Function

SSDeleteRangeError:
   SSDeleteRange = Err.Number

End Function

Public Function SSDeleteTable%(SS As Object)

   On Error GoTo SSDeleteTableError
   SS.DeleteTable
   SSDeleteTable = 0
   Exit Function

SSDeleteTableError:
   SSDeleteTable = Err.Number

End Function

Public Function SSDraw%(SS As Object, ByVal hDC%, ByVal X%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal nRow%, ByVal nCol%, nRows%, nCols%, ByVal nFixedRow%, ByVal nFixedCol%, ByVal nFixedRows%, ByVal nFixedCols%)

   On Error GoTo SSDrawError
   SS.Draw hDC, X, Y, cx, cy, nRow, nCol, nRows, nCols, nFixedRow, nFixedCol, nFixedRows, nFixedCols
   SSDraw = 0
   Exit Function

SSDrawError:
   SSDraw = Err.Number

End Function

Public Function SSEditClear%(SS As Object, ByVal nClearType%)

   On Error GoTo SSEditClearError
   SS.EditClear nClearType
   SSEditClear = 0
   Exit Function

SSEditClearError:
   SSEditClear = Err.Number

End Function

Public Function SSEditCopy%(SS As Object)

   On Error GoTo SSEditCopyError
   SS.EditCopy
   SSEditCopy = 0
   Exit Function

SSEditCopyError:
   SSEditCopy = Err.Number

End Function

Public Function SSEditCopyDown%(SS As Object)

   On Error GoTo SSEditCopyDownError
   SS.EditCopyDown
   SSEditCopyDown = 0
   Exit Function

SSEditCopyDownError:
   SSEditCopyDown = Err.Number

End Function

Public Function SSEditCopyRight%(SS As Object)

   On Error GoTo SSEditCopyRightError
   SS.EditCopyRight
   SSEditCopyRight = 0
   Exit Function

SSEditCopyRightError:
   SSEditCopyRight = Err.Number

End Function

Public Function SSEditCut%(SS As Object)

   On Error GoTo SSEditCutError
   SS.EditCut
   SSEditCut = 0
   Exit Function

SSEditCutError:
   SSEditCut = Err.Number

End Function

Public Function SSEditDelete%(SS As Object, ByVal nShiftType%)

   On Error GoTo SSEditDeleteError
   SS.EditDelete nShiftType
   SSEditDelete = 0
   Exit Function

SSEditDeleteError:
   SSEditDelete = Err.Number

End Function

Public Function SSEditInsert%(SS As Object, ByVal nShiftType%)

   On Error GoTo SSEditInsertError
   SS.EditInsert nShiftType
   SSEditInsert = 0
   Exit Function

SSEditInsertError:
   SSEditInsert = Err.Number

End Function

Public Function SSEditPaste%(SS As Object)

   On Error GoTo SSEditPasteError
   SS.EditPaste
   SSEditPaste = 0
   Exit Function

SSEditPasteError:
   SSEditPaste = Err.Number

End Function

Public Function SSEndEdit%(SS As Object)

   On Error GoTo SSEndEditError
   SS.EndEdit
   SSEndEdit = 0
   Exit Function

SSEndEditError:
   SSEndEdit = Err.Number

End Function

Public Function SSErrorNumberToText%(SS As Object, ByVal nError%, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSErrorNumberToTextError
   pBuf = SS.ErrorNumberToText(nError)
   SSErrorNumberToText = 0
   Exit Function

SSErrorNumberToTextError:
   SSErrorNumberToText = Err.Number

End Function

Public Function SSFilePageSetupDlg%(SS As Object)

   On Error GoTo SSFilePageSetupDlgError
   SS.FilePageSetupDlg
   SSFilePageSetupDlg = 0
   Exit Function

SSFilePageSetupDlgError:
   SSFilePageSetupDlg = Err.Number

End Function

Public Function SSFilePrint%(SS As Object, ByVal bShowPrintDlg%)

   On Error GoTo SSFilePrintError
   SS.FilePrint bShowPrintDlg
   SSFilePrint = 0
   Exit Function

SSFilePrintError:
   SSFilePrint = Err.Number

End Function

Public Function SSFilePrintSetupDlg%(SS As Object)

   On Error GoTo SSFilePrintSetupDlgError
   SS.FilePrintSetupDlg
   SSFilePrintSetupDlg = 0
   Exit Function

SSFilePrintSetupDlgError:
   SSFilePrintSetupDlg = Err.Number

End Function

Public Function SSFormatAlignmentDlg%(SS As Object)

   On Error GoTo SSFormatAlignmentDlgError
   SS.FormatAlignmentDlg
   SSFormatAlignmentDlg = 0
   Exit Function

SSFormatAlignmentDlgError:
   SSFormatAlignmentDlg = Err.Number

End Function

Public Function SSFormatBorderDlg%(SS As Object)

   On Error GoTo SSFormatBorderDlgError
   SS.FormatBorderDlg
   SSFormatBorderDlg = 0
   Exit Function

SSFormatBorderDlgError:
   SSFormatBorderDlg = Err.Number

End Function

Public Function SSFormatCurrency0%(SS As Object)

   On Error GoTo SSFormatCurrency0Error
   SS.FormatCurrency0
   SSFormatCurrency0 = 0
   Exit Function

SSFormatCurrency0Error:
   SSFormatCurrency0 = Err.Number

End Function

Public Function SSFormatCurrency2%(SS As Object)

   On Error GoTo SSFormatCurrency2Error
   SS.FormatCurrency2
   SSFormatCurrency2 = 0
   Exit Function

SSFormatCurrency2Error:
   SSFormatCurrency2 = Err.Number

End Function

Public Function SSFormatDefaultFontDlg%(SS As Object)

   On Error GoTo SSFormatDefaultFontDlgError
   SS.FormatDefaultFontDlg
   SSFormatDefaultFontDlg = 0
   Exit Function

SSFormatDefaultFontDlgError:
   SSFormatDefaultFontDlg = Err.Number

End Function

Public Function SSFormatFixed%(SS As Object)

   On Error GoTo SSFormatFixedError
   SS.FormatFixed
   SSFormatFixed = 0
   Exit Function

SSFormatFixedError:
   SSFormatFixed = Err.Number

End Function

Public Function SSFormatFixed2%(SS As Object)

   On Error GoTo SSFormatFixed2Error
   SS.FormatFixed2
   SSFormatFixed2 = 0
   Exit Function

SSFormatFixed2Error:
   SSFormatFixed2 = Err.Number

End Function

Public Function SSFormatFontDlg%(SS As Object)

   On Error GoTo SSFormatFontDlgError
   SS.FormatFontDlg
   SSFormatFontDlg = 0
   Exit Function

SSFormatFontDlgError:
   SSFormatFontDlg = Err.Number

End Function

Public Function SSFormatFraction%(SS As Object)

   On Error GoTo SSFormatFractionError
   SS.FormatFraction
   SSFormatFraction = 0
   Exit Function

SSFormatFractionError:
   SSFormatFraction = Err.Number

End Function

Public Function SSFormatGeneral%(SS As Object)

   On Error GoTo SSFormatGeneralError
   SS.FormatGeneral
   SSFormatGeneral = 0
   Exit Function

SSFormatGeneralError:
   SSFormatGeneral = Err.Number

End Function

Public Function SSFormatHmmampm%(SS As Object)

   On Error GoTo SSFormatHmmampmError
   SS.FormatHmmampm
   SSFormatHmmampm = 0
   Exit Function

SSFormatHmmampmError:
   SSFormatHmmampm = Err.Number

End Function

Public Function SSFormatMdyy%(SS As Object)

   On Error GoTo SSFormatMdyyError
   SS.FormatMdyy
   SSFormatMdyy = 0
   Exit Function

SSFormatMdyyError:
   SSFormatMdyy = Err.Number

End Function

Public Function SSFormatNumberDlg%(SS As Object)

   On Error GoTo SSFormatNumberDlgError
   SS.FormatNumberDlg
   SSFormatNumberDlg = 0
   Exit Function

SSFormatNumberDlgError:
   SSFormatNumberDlg = Err.Number

End Function

Public Function SSFormatPatternDlg%(SS As Object)

   On Error GoTo SSFormatPatternDlgError
   SS.FormatPatternDlg
   SSFormatPatternDlg = 0
   Exit Function

SSFormatPatternDlgError:
   SSFormatPatternDlg = Err.Number

End Function

Public Function SSFormatPercent%(SS As Object)

   On Error GoTo SSFormatPercentError
   SS.FormatPercent
   SSFormatPercent = 0
   Exit Function

SSFormatPercentError:
   SSFormatPercent = Err.Number

End Function

Public Function SSFormatRCNr%(SS As Object, ByVal nRow%, ByVal nCol%, ByVal bDoAbsolute%, ByRef pBuf$, ByVal nBufSize%)

   On Error GoTo SSFormatRCNrError
   pBuf = SS.FormatRCNr(nRow, nCol, bDoAbsolute)
   SSFormatRCNr = 0
   Exit Function

SSFormatRCNrError:
   SSFormatRCNr = Err.Number

End Function

Public Function SSFormatScientific%(SS As Object)

   On Error GoTo SSFormatScientificError
   SS.FormatScientific
   SSFormatScientific = 0
   Exit Function

SSFormatScientificError:
   SSFormatScientific = Err.Number

End Function

Public Function SSGetActiveCell%(SS As Object, pRow%, pCol%)

   On Error GoTo SSGetActiveCellError
   SS.GetActiveCell pRow, pCol
   SSGetActiveCell = 0
   Exit Function

SSGetActiveCellError:
   SSGetActiveCell = Err.Number

End Function

Public Function SSGetAlignment%(SS As Object, pHorizontal%, pWordWrap%, pVertical%, pOrientation%)

   On Error GoTo SSGetAlignmentError
   SS.GetAlignment pHorizontal, pWordWrap, pVertical, pOrientation
   SSGetAlignment = 0
   Exit Function

SSGetAlignmentError:
   SSGetAlignment = Err.Number

End Function

Public Function SSGetAllowArrows%(SS As Object, pAllowArrows%)

   On Error GoTo SSGetAllowArrowsError
   pAllowArrows = SS.AllowArrows
   SSGetAllowArrows = 0
   Exit Function

SSGetAllowArrowsError:
   SSGetAllowArrows = Err.Number

End Function

Public Function SSGetAllowDelete%(SS As Object, pAllowDelete%)

   On Error GoTo SSGetAllowDeleteError
   pAllowDelete = SS.AllowDelete
   SSGetAllowDelete = 0
   Exit Function

SSGetAllowDeleteError:
   SSGetAllowDelete = Err.Number

End Function

Public Function SSGetAllowEditHeaders%(SS As Object, pAllowEditHeaders%)

   On Error GoTo SSGetAllowEditHeadersError
   pAllowEditHeaders = SS.AllowEditHeaders
   SSGetAllowEditHeaders = 0
   Exit Function

SSGetAllowEditHeadersError:
   SSGetAllowEditHeaders = Err.Number

End Function

Public Function SSGetAllowFillRange%(SS As Object, pAllowFillRange%)

   On Error GoTo SSGetAllowFillRangeError
   pAllowFillRange = SS.AllowFillRange
   SSGetAllowFillRange = 0
   Exit Function

SSGetAllowFillRangeError:
   SSGetAllowFillRange = Err.Number

End Function

Public Function SSGetAllowInCellEditing%(SS As Object, pAllowInCellEditing%)

   On Error GoTo SSGetAllowInCellEditingError
   pAllowInCellEditing = SS.AllowInCellEditing
   SSGetAllowInCellEditing = 0
   Exit Function

SSGetAllowInCellEditingError:
   SSGetAllowInCellEditing = Err.Number

End Function

Public Function SSGetAllowMoveRange%(SS As Object, pAllowMoveRange%)

   On Error GoTo SSGetAllowMoveRangeError
   pAllowMoveRange = SS.AllowMoveRange
   SSGetAllowMoveRange = 0
   Exit Function

SSGetAllowMoveRangeError:
   SSGetAllowMoveRange = Err.Number

End Function

Public Function SSGetAllowObjSelections%(SS As Object, pAllowObjSelections%)

   On Error GoTo SSGetAllowObjSelectionsError
   pAllowObjSelections = SS.AllowObjSelections
   SSGetAllowObjSelections = 0
   Exit Function

SSGetAllowObjSelectionsError:
   SSGetAllowObjSelections = Err.Number

End Function

Public Function SSGetAllowResize%(SS As Object, pAllowResize%)

   On Error GoTo SSGetAllowResizeError
   pAllowResize = SS.AllowResize
   SSGetAllowResize = 0
   Exit Function

SSGetAllowResizeError:
   SSGetAllowResize = Err.Number

End Function

Public Function SSGetAllowSelections%(SS As Object, pAllowSelections%)

   On Error GoTo SSGetAllowSelectionsError
   pAllowSelection = SS.AllowSelections
   SSGetAllowSelections = 0
   Exit Function

SSGetAllowSelectionsError:
   SSGetAllowSelections = Err.Number

End Function

Public Function SSGetAllowTabs%(SS As Object, pAllowTabs%)

   On Error GoTo SSGetAllowTabsError
   pAllowTabs = SS.AllowTabs
   SSGetAllowTabs = 0
   Exit Function

SSGetAllowTabsError:
   SSGetAllowTabs = Err.Number

End Function

Public Function SSGetAllowFormulas%(SS As Object, pAllowFormulas%)

   On Error GoTo SSGetAllowFormulasError
   pAllowFormulas = SS.AllowFormulas
   SSGetAllowFormulas = 0
   Exit Function

SSGetAllowFormulasError:
   SSGetAllowFormulas = Err.Number

End Function

Public Function SSGetAppName%(SS As Object, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetAppNameError
   pBuf = SS.APPNAME
   SSGetAppName = 0
   Exit Function

SSGetAppNameError:
   SSGetAppName = Err.Number

End Function

Public Function SSGetAutoRecalc%(SS As Object, pAutoRecalc%)

   On Error GoTo SSGetAutoRecalcError
   pAutoRecalc = SS.AutoRecalc
   SSGetAutoRecalc = 0
   Exit Function

SSGetAutoRecalcError:
   SSGetAutoRecalc = Err.Number

End Function

Public Function SSGetBackColor%(SS As Object, pBackColor&)

   On Error GoTo SSGetBackColorError
   pBackColor = SS.BackColor
   SSGetBackColor = 0
   Exit Function

SSGetBackColorError:
   SSGetBackColor = Err.Number

End Function

Public Function SSGetBorder%(SS As Object, pLeft%, pRight%, pTop%, pBottom%, pShade%, pcrLeft&, pcrRight&, pcrTop&, pcrBottom&)

   On Error GoTo SSGetBorderError
   SS.GetBorder pLeft, pRight, pTop, pBottom, pShade, pcrLeft, pcrRight, pcrTop, pcrBottom
   SSGetBorder = 0
   Exit Function

SSGetBorderError:
   SSGetBorder = Err.Number

End Function

Public Function SSGetColText%(SS As Object, ByVal nCol%, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetColTextError
   pBuf = SS.ColText(nCol)
   SSGetColText = 0
   Exit Function

SSGetColTextError:
   SSGetColText = Err.Number

End Function

Public Function SSGetColWidth%(SS As Object, ByVal nCol%, pWidth%)

   On Error GoTo SSGetColWidthError
   pWidth = SS.ColWidth(nCol)
   SSGetColWidth = 0
   Exit Function

SSGetColWidthError:
   SSGetColWidth = Err.Number

End Function

Public Function SSGetDefaultFont%(SS As Object, ByVal pBuf$, ByVal nBufSize%, pSize%)

   On Error GoTo SSGetDefaultFontError
   SS.GetDefaultFont pBuf, pSize
   SSGetDefaultFont = 0
   Exit Function

SSGetDefaultFontError:
   SSGetDefaultFont = Err.Number

End Function

Public Function SSGetDefinedName%(SS As Object, ByVal pName$, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetDefinedNameError
   pBuf = SS.GetDefinedName(pName)
   SSGetDefinedName = 0
   Exit Function

SSGetDefinedNameError:
   SSGetDefinedName = Err.Number

End Function

Public Function SSGetEnableProtection%(SS As Object, pEnableProtection%)

   On Error GoTo SSGetEnableProtectionError
   pEnableProtection = SS.EnableProtection
   SSGetEnableProtection = 0
   Exit Function

SSGetEnableProtectionError:
   SSGetEnableProtection = Err.Number

End Function

Public Function SSGetEnterMovesDown%(SS As Object, pEnterMovesDown%)

   On Error GoTo SSGetEnterMovesDownError
   pEnterMovesDown = SS.EnterMovesDown
   SSGetEnterMovesDown = 0
   Exit Function

SSGetEnterMovesDownError:
   SSGetEnterMovesDown = Err.Number

End Function

Public Function SSGetEntry%(SS As Object, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetEntryError
   pBuf = SS.Entry
   SSGetEntry = 0
   Exit Function

SSGetEntryError:
   SSGetEntry = Err.Number

End Function

Public Function SSGetEntryRC%(SS As Object, ByVal nRow%, ByVal nCol%, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetEntryRCError
   pBuf = SS.EntryRC(nRow, nCol)
   SSGetEntryRC = 0
   Exit Function

SSGetEntryRCError:
   SSGetEntryRC = Err.Number

End Function

Public Function SSGetExtraColor%(SS As Object, pExtraColor&)

   On Error GoTo SSGetExtraColorError
   pExtraColor = SS.ExtraColor
   SSGetExtraColor = 0
   Exit Function

SSGetExtraColorError:
   SSGetExtraColor = Err.Number

End Function

Public Function SSGetFixedCols%(SS As Object, pC1%, pCols%)

   On Error GoTo SSGetFixedColsError
   pCols = SS.FixedCols(pC1)
   SSGetFixedCols = 0
   Exit Function

SSGetFixedColsError:
   SSGetFixedCols = Err.Number

End Function

Public Function SSGetFixedRows%(SS As Object, pR1%, pRows%)

   On Error GoTo SSGetFixedRowsError
   pRows = SS.FixedRows(pR1)
   SSGetFixedRows = 0
   Exit Function

SSGetFixedRowsError:
   SSGetFixedRows = Err.Number

End Function

Public Function SSGetFont%(SS As Object, ByVal pBuf$, ByVal nBufSize%, pSize%, pBold%, pItalic%, pUnderline%, pStrikeout%, pcrColor&, pOutline%, pShadow%)

   On Error GoTo SSGetFontError
   SS.GetFont pBuf, pSize, pBold, pItalic, pUnderline, pStrikeout, pcrColor, pOutline, pShadow
   SSGetFont = 0
   Exit Function

SSGetFontError:
   SSGetFont = Err.Number

End Function

Public Function SSGetFormattedText%(SS As Object, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetFormattedTextError
   pBuf = SS.FormattedText
   SSGetFormattedText = 0
   Exit Function

SSGetFormattedTextError:
   SSGetFormattedText = Err.Number

End Function

Public Function SSGetFormattedTextRC%(SS As Object, ByVal nRow%, ByVal nCol%, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetFormattedTextRCError
   pBuf = SS.FormattedTextRC(nRow, nCol)
   SSGetFormattedTextRC = 0
   Exit Function

SSGetFormattedTextRCError:
   SSGetFormattedTextRC = Err.Number

End Function

Public Function SSGetFormula%(SS As Object, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetFormulaError
   pBuf = SS.Formula
   SSGetFormula = 0
   Exit Function

SSGetFormulaError:
   SSGetFormula = Err.Number

End Function

Public Function SSGetFormulaRC%(SS As Object, ByVal nRow%, ByVal nCol%, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetFormulaRCError
   pBuf = SS.FormulaRC(nRow, nCol)
   SSGetFormulaRC = 0
   Exit Function

SSGetFormulaRCError:
   SSGetFormulaRC = Err.Number

End Function

Public Function SSGetHdrHeight%(SS As Object, pHeight%)

   On Error GoTo SSGetHdrHeightError
   pHeight = SS.HdrHeight
   SSGetHdrHeight = 0
   Exit Function

SSGetHdrHeightError:
   SSGetHdrHeight = Err.Number

End Function

Public Function SSGetHdrSelection%(SS As Object, pTopLeftHdr%, pRowHdr%, pColHdr%)

   On Error GoTo SSGetHdrSelectionError
   SS.GetHdrSelection pTopLeftHdr, pRowHdr, pColHdr
   SSGetHdrSelection = 0
   Exit Function

SSGetHdrSelectionError:
   SSGetHdrSelection = Err.Number

End Function

Public Function SSGetHdrWidth%(SS As Object, pWidth%)

   On Error GoTo SSGetHdrWidthError
   pWidth = SS.HdrWidth
   SSGetHdrWidth = 0
   Exit Function

SSGetHdrWidthError:
   SSGetHdrWidth = Err.Number

End Function

Public Function SSGetIteration%(SS As Object, pIteration%, pMaxIterations%, pMaxChange#)

   On Error GoTo SSGetIterationError
   SS.GetIteration pIteration, pMaxIterations, pMaxChange
   SSGetIteration = 0
   Exit Function

SSGetIterationError:
   SSGetIteration = Err.Number

End Function

Public Function SSGetLastCol%(SS As Object, pLastCol%)

   On Error GoTo SSGetLastColError
   pLastCol = SS.LastCol
   SSGetLastCol = 0
   Exit Function

SSGetLastColError:
   SSGetLastCol = Err.Number

End Function

Public Function SSGetLastColForRow%(SS As Object, ByVal nRow%, pLastColForRow%)

   On Error GoTo SSGetLastColForRowError
   pLastColForRow = SS.LastColForRow(nRow)
   SSGetLastColForRow = 0
   Exit Function

SSGetLastColForRowError:
   SSGetLastColForRow = Err.Number

End Function

Public Function SSGetLastRow%(SS As Object, pLastRow%)

   On Error GoTo SSGetLastRowError
   pLastRow = SS.LastRow
   SSGetLastRow = 0
   Exit Function

SSGetLastRowError:
   SSGetLastRow = Err.Number

End Function

Public Function SSGetLeftCol%(SS As Object, pLeftCol%)

   On Error GoTo SSGetLeftColError
   pLeftCol = SS.LeftCol
   SSGetLeftCol = 0
   Exit Function

SSGetLeftColError:
   SSGetLeftCol = Err.Number

End Function

Public Function SSGetLineStyle%(SS As Object, pStyle%, pcrColor&, pWeight%)

   On Error GoTo SSGetLineStyleError
   SS.LineStyle pStyle, pcrColor, pWeight
   SSGetLineStyle = 0
   Exit Function

SSGetLineStyleError:
   SSGetLineStyle = Err.Number

End Function

Public Function SSGetLogicalRC%(SS As Object, ByVal nRow%, ByVal nCol%, pIsTrue%)

   On Error GoTo SSGetLogicalRCError
   pIsTrue = SS.LogicalRC(nRow, nCol)
   SSGetLogicalRC = 0
   Exit Function

SSGetLogicalRCError:
   SSGetLogicalRC = Err.Number

End Function

Public Function SSGetMinCol%(SS As Object, pMinCol%)

   On Error GoTo SSGetMinColError
   pMinCol = SS.MinCol
   SSGetMinCol = 0
   Exit Function

SSGetMinColError:
   SSGetMinCol = Err.Number

End Function

Public Function SSGetMinRow%(SS As Object, pMinRow%)

   On Error GoTo SSGetMinRowError
   pMinRow = SS.MinRow
   SSGetMinRow = 0
   Exit Function

SSGetMinRowError:
   SSGetMinRow = Err.Number

End Function

Public Function SSGetMaxCol%(SS As Object, pMaxCol%)

   On Error GoTo SSGetMaxColError
   pMaxCol = SS.MaxCol
   SSGetMaxCol = 0
   Exit Function

SSGetMaxColError:
   SSGetMaxCol = Err.Number

End Function

Public Function SSGetMaxRow%(SS As Object, pMaxRow%)

   On Error GoTo SSGetMaxRowError
   pMaxRow = SS.MaxRow
   SSGetMaxRow = 0
   Exit Function

SSGetMaxRowError:
   SSGetMaxRow = Err.Number

End Function

Public Function SSGetMode%(SS As Object, pMode%)

   On Error GoTo SSGetModeError
   pMode = SS.Mode
   SSGetMode = 0
   Exit Function

SSGetModeError:
   SSGetMode = Err.Number

End Function

Public Function SSGetNumber%(SS As Object, pNumber#)

   On Error GoTo SSGetNumberError
   pNumber = SS.Number
   SSGetNumber = 0
   Exit Function

SSGetNumberError:
   SSGetNumber = Err.Number

End Function

Public Function SSGetNumberFormat%(SS As Object, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetNumberFormatError
   pBuf = SS.NumberFormat
   SSGetNumberFormat = 0
   Exit Function

SSGetNumberFormatError:
   SSGetNumberFormat = Err.Number

End Function

Public Function SSGetNumberRC%(SS As Object, ByVal nRow%, ByVal nCol%, pNumber#)

   On Error GoTo SSGetNumberRCError
   pNumber = SS.NumberRC(nRow, nCol)
   SSGetNumberRC = 0
   Exit Function

SSGetNumberRCError:
   SSGetNumberRC = Err.Number

End Function

Public Function SSGetPattern%(SS As Object, pPattern%, pcrFG&, pcrBG&)

   On Error GoTo SSGetPatternError
   SS.GetPattern pPattern, pcrFG, pcrBG
   SSGetPattern = 0
   Exit Function

SSGetPatternError:
   SSGetPattern = Err.Number

End Function

Public Function SSGetPolyEditMode%(SS As Object, pPolyEditMode%)

   On Error GoTo SSGetPolyEditModeError
   pPolyEditMode = SS.PolyEditMode
   SSGetPolyEditMode = 0
   Exit Function

SSGetPolyEditModeError:
   SSGetPolyEditMode = Err.Number

End Function

Public Function SSGetPrintArea%(SS As Object, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetPrintAreaError
   pBuf = SS.PrintArea
   SSGetPrintArea = 0
   Exit Function

SSGetPrintAreaError:
   SSGetPrintArea = Err.Number

End Function

Public Function SSGetPrintBottomMargin%(SS As Object, pPrintBottomMargin#)

   On Error GoTo SSGetPrintBottomMarginError
   pPrintBottomMargin = SS.PrintBottomMargin
   SSGetPrintBottomMargin = 0
   Exit Function

SSGetPrintBottomMarginError:
   SSGetPrintBottomMargin = Err.Number

End Function

Public Function SSGetPrintColHeading%(SS As Object, pPrintColHeading%)

   On Error GoTo SSGetPrintColHeadingError
   pPrintColHeading = SS.PrintColHeading
   SSGetPrintColHeading = 0
   Exit Function

SSGetPrintColHeadingError:
   SSGetPrintColHeading = Err.Number

End Function

Public Function SSGetPrintFooter%(SS As Object, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetPrintFooterError
   pBuf = SS.PrintFooter
   SSGetPrintFooter = 0
   Exit Function

SSGetPrintFooterError:
   SSGetPrintFooter = Err.Number

End Function

Public Function SSGetPrintGridLines%(SS As Object, pPrintGridLines%)

   On Error GoTo SSGetPrintGridLinesError
   pPrintGridLines = SS.PrintGridLines
   SSGetPrintGridLines = 0
   Exit Function

SSGetPrintGridLinesError:
   SSGetPrintGridLines = Err.Number

End Function

Public Function SSGetPrintHCenter%(SS As Object, pPrintHCenter%)

   On Error GoTo SSGetPrintHCenterError
   pPrintHCenter = SS.PrintHCenter
   SSGetPrintHCenter = 0
   Exit Function

SSGetPrintHCenterError:
   SSGetPrintHCenter = Err.Number

End Function

Public Function SSGetPrintHeader%(SS As Object, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetPrintHeaderError
   pBuf = SS.PrintHeader
   SSGetPrintHeader = 0
   Exit Function

SSGetPrintHeaderError:
   SSGetPrintHeader = Err.Number

End Function

Public Function SSGetPrintLandscape%(SS As Object, pLandscape%)

   On Error GoTo SSGetPrintLandscapeError
   pLandscape = SS.PrintLandscape
   SSGetPrintLandscape = 0
   Exit Function

SSGetPrintLandscapeError:
   SSGetPrintLandscape = Err.Number

End Function

Public Function SSGetPrintLeftMargin%(SS As Object, pPrintLeftMargin#)

   On Error GoTo SSGetPrintLeftMarginError
   pPrintLeftMargin = SS.PrintLeftMargin
   SSGetPrintLeftMargin = 0
   Exit Function

SSGetPrintLeftMarginError:
   SSGetPrintLeftMargin = Err.Number

End Function

Public Function SSGetPrintLeftToRight%(SS As Object, pPrintLeftToRight%)

   On Error GoTo SSGetPrintLeftToRightError
   pPrintLeftToRight = SS.PrintLeftToRight
   SSGetPrintLeftToRight = 0
   Exit Function

SSGetPrintLeftToRightError:
   SSGetPrintLeftToRight = Err.Number

End Function

Public Function SSGetPrintNoColor%(SS As Object, pNoColor%)

   On Error GoTo SSGetPrintNoColorError
   pNoColor = SS.PrintNoColor
   SSGetPrintNoColor = 0
   Exit Function

SSGetPrintNoColorError:
   SSGetPrintNoColor = Err.Number

End Function

Public Function SSGetPrintRightMargin%(SS As Object, pPrintRightMargin#)

   On Error GoTo SSGetPrintRightMarginError
   pPrintRightMargin = SS.PrintRightMargin
   SSGetPrintRightMargin = 0
   Exit Function

SSGetPrintRightMarginError:
   SSGetPrintRightMargin = Err.Number

End Function

Public Function SSGetPrintRowHeading%(SS As Object, pPrintRowHeading%)

   On Error GoTo SSGetPrintRowHeadingError
   pPrintRowHeading = SS.PrintRowHeading
   SSGetPrintRowHeading = 0
   Exit Function

SSGetPrintRowHeadingError:
   SSGetPrintRowHeading = Err.Number

End Function

Public Function SSGetPrintScale%(SS As Object, pScale%, pFitToPage%, pVPages%, pHPages%)

   On Error GoTo SSGetPrintScaleError
   SS.GetPrintScale pScale, pFitToPage, pVPages, pHPages
   SSGetPrintScale = 0
   Exit Function

SSGetPrintScaleError:
   SSGetPrintScale = Err.Number

End Function

Public Function SSGetPrintTitles%(SS As Object, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetPrintTitlesError
   pBuf = SS.PrintTitles
   SSGetPrintTitles = 0
   Exit Function

SSGetPrintTitlesError:
   SSGetPrintTitles = Err.Number

End Function

Public Function SSGetPrintTopMargin%(SS As Object, pPrintTopMargin#)

   On Error GoTo SSGetPrintTopMarginError
   pPrintTopMargin = SS.PrintTopMargin
   SSGetPrintTopMargin = 0
   Exit Function

SSGetPrintTopMarginError:
   SSGetPrintTopMargin = Err.Number

End Function

Public Function SSGetPrintVCenter%(SS As Object, pPrintVCenter%)

   On Error GoTo SSGetPrintVCenterError
   pPrintVCenter = SS.PrintVCenter
   SSGetPrintVCenter = 0
   Exit Function

SSGetPrintVCenterError:
   SSGetPrintVCenter = Err.Number

End Function

Public Function SSGetProtection%(SS As Object, pLocked%, pHidden%)

   On Error GoTo SSGetProtectionError
   SS.GetProtection pLocked, pHidden
   SSGetProtection = 0
   Exit Function

SSGetProtectionError:
   SSGetProtection = Err.Number

End Function

Public Function SSGetRepaint%(SS As Object, pRepaint%)

   On Error GoTo SSGetRepaintError
   pRepaint = SS.Repaint
   SSGetRepaint = 0
   Exit Function

SSGetRepaintError:
   SSGetRepaint = Err.Number

End Function

Public Function SSGetRowHeight%(SS As Object, ByVal nRow%, pHeight%)

   On Error GoTo SSGetRowHeightError
   pHeight = SS.RowHeight(nRow)
   SSGetRowHeight = 0
   Exit Function

SSGetRowHeightError:
   SSGetRowHeight = Err.Number

End Function

Public Function SSGetRowMode%(SS As Object, pRowMode%)

   On Error GoTo SSGetRowModeError
   pRowMode = SS.RowMode
   SSGetRowMode = 0
   Exit Function

SSGetRowModeError:
   SSGetRowMode = Err.Number

End Function

Public Function SSGetRowText%(SS As Object, ByVal nRow%, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetRowTextError
   pBuf = SS.RowText(nRow)
   SSGetRowText = 0
   Exit Function

SSGetRowTextError:
   SSGetRowText = Err.Number

End Function

Public Function SSGetScale%(SS As Object, pScale%)

   On Error GoTo SSGetScaleError
   pScale = SS.ViewScale
   SSGetScale = 0
   Exit Function

SSGetScaleError:
   SSGetScale = Err.Number

End Function

Public Function SSGetSelection%(SS As Object, ByVal nSelection%, pR1%, pC1%, pR2%, pC2%)

   On Error GoTo SSGetSelectionError
   SS.GetSelection nSelection, pR1, pC1, pR2, pC2
   SSGetSelection = 0
   Exit Function

SSGetSelectionError:
   SSGetSelection = Err.Number

End Function

Public Function SSGetSelectionCount%(SS As Object, pCount%)

   On Error GoTo SSGetSelectionCountError
   pCount = SS.SelectionCount
   SSGetSelectionCount = 0
   Exit Function

SSGetSelectionCountError:
   SSGetSelectionCount = Err.Number

End Function

Public Function SSGetSelectionRef%(SS As Object, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetSelectionRefError
   pBuf = SS.SelectionRef
   SSGetSelectionRef = 0
   Exit Function

SSGetSelectionRefError:
   SSGetSelectionRef = Err.Number

End Function

Public Function SSGetShowColHeading%(SS As Object, pShowColHeading%)

   On Error GoTo SSGetShowColHeadingError
   pShowColHeading = SS.ShowColHeading
   SSGetShowColHeading = 0
   Exit Function

SSGetShowColHeadingError:
   SSGetShowColHeading = Err.Number

End Function

Public Function SSGetShowFormulas%(SS As Object, pShowFormulas%)

   On Error GoTo SSGetShowFormulasError
   pShowFormulas = SS.ShowFormulas
   SSGetShowFormulas = 0
   Exit Function

SSGetShowFormulasError:
   SSGetShowFormulas = Err.Number

End Function

Public Function SSGetShowGridLines%(SS As Object, pShowGridLines%)

   On Error GoTo SSGetShowGridLinesError
   pShowGridLines = SS.ShowGridLines
   SSGetShowGridLines = 0
   Exit Function

SSGetShowGridLinesError:
   SSGetShowGridLines = Err.Number

End Function

Public Function SSGetShowHScrollBar%(SS As Object, pShowHScrollBar%)

   On Error GoTo SSGetShowHScrollBarError
   pShowHScrollBar = SS.ShowHScrollBar
   SSGetShowHScrollBar = 0
   Exit Function

SSGetShowHScrollBarError:
   SSGetShowHScrollBar = Err.Number

End Function

Public Function SSGetShowRowHeading%(SS As Object, pShowRowHeading%)

   On Error GoTo SSGetShowRowHeadingError
   pShowRowHeading = SS.ShowRowHeading
   SSGetShowRowHeading = 0
   Exit Function

SSGetShowRowHeadingError:
   SSGetShowRowHeading = Err.Number

End Function

Public Function SSGetShowSelections%(SS As Object, pShowSelections%)

   On Error GoTo SSGetShowSelectionsError
   pShowSelections = SS.ShowSelections
   SSGetShowSelections = 0
   Exit Function

SSGetShowSelectionsError:
   SSGetShowSelections = Err.Number

End Function

Public Function SSGetShowVScrollBar%(SS As Object, pShowVScrollBar%)

   On Error GoTo SSGetShowVScrollBarError
   pShowVScrollBar = SS.ShowVScrollBar
   SSGetShowVScrollBar = 0
   Exit Function

SSGetShowVScrollBarError:
   SSGetShowVScrollBar = Err.Number

End Function

Public Function SSGetShowZeroValues%(SS As Object, pShowZeroValues%)

   On Error GoTo SSGetShowZeroValuesError
   pShowZeroValues = SS.ShowZeroValues
   SSGetShowZeroValues = 0
   Exit Function

SSGetShowZeroValuesError:
   SSGetShowZeroValues = Err.Number

End Function

Public Function SSGetTabbedText%(SS As Object, ByVal nR1%, ByVal nC1%, ByVal nR2%, ByVal nC2%, ByVal bValuesOnly%, phText%)

   On Error GoTo SSGetTabbedTextError
   SS.GetTabbedText nR1, nC1, nR2, nC2, bValuesOnly, phText
   SSGetTabbedText = 0
   Exit Function

SSGetTabbedTextError:
   SSGetTabbedText = Err.Number

End Function

Public Function SSGetText%(SS As Object, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetTextError
   pBuf = SS.Text
   SSGetText = 0
   Exit Function

SSGetTextError:
   SSGetText = Err.Number

End Function

Public Function SSGetTextRC%(SS As Object, ByVal nRow%, ByVal nCol%, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetTextRCError
   pBuf = SS.TextRC(nRow, nCol)
   SSGetTextRC = 0
   Exit Function

SSGetTextRCError:
   SSGetTextRC = Err.Number

End Function

Public Function SSGetTitle%(SS As Object, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetTitleError
   pBuf = SS.Title
   SSGetTitle = 0
   Exit Function

SSGetTitleError:
   SSGetTitle = Err.Number

End Function

Public Function SSGetTopLeftText%(SS As Object, ByRef pBuf$, ByVal nBufSize%)

   On Error GoTo SSGetTopLeftTextError
   pBuf = SS.TopLeftText
   SSGetTopLeftText = 0
   Exit Function

SSGetTopLeftTextError:
   SSGetTopLeftText = Err.Number

End Function

Public Function SSGetTopRow%(SS As Object, pTopRow%)

   On Error GoTo SSGetTopRowError
   pTopRow = SS.TopRow
   SSGetTopRow = 0
   Exit Function

SSGetTopRowError:
   SSGetTopRow = Err.Number

End Function

Public Function SSGetTypeRC%(SS As Object, ByVal nRow%, ByVal nCol%, pType%)

   On Error GoTo SSGetTypeRCError
   pType = SS.TypeRC(nRow, nCol)
   SSGetTypeRC = 0
   Exit Function

SSGetTypeRCError:
   SSGetTypeRC = Err.Number

End Function

Public Function SSGotoDlg%(SS As Object)

   On Error GoTo SSGotoDlgError
   SS.GotoDlg
   SSGotoDlg = 0
   Exit Function

SSGotoDlgError:
   SSGotoDlg = Err.Number

End Function

Public Sub SSHeapMin(SS As Object)

   SS.HeapMin

End Sub

Public Function SSInitTable%(SS As Object)

   On Error GoTo SSInitTableError
   SS.InitTable
   SSInitTable = 0
   Exit Function

SSInitTableError:
   SSInitTable = Err.Number

End Function

Public Function SSInsertRange%(SS As Object, ByVal nR1%, ByVal nC1%, ByVal nR2%, ByVal nC2%, ByVal nShiftType%)

   On Error GoTo SSInsertRangeError
   SS.InsertRange nR1, nC1, nR2, nC2, nShiftType
   SSInsertRange = 0
   Exit Function

SSInsertRangeError:
   SSInsertRange = Err.Number

End Function

Public Function SSLaunchAppDesigner%(SS As Object)

   On Error GoTo SSLaunchAppDesignerError
   SS.LaunchAppDesigner
   SSLaunchAppDesigner = 0
   Exit Function

SSLaunchAppDesignerError:
   SSLaunchAppDesigner = Err.Number

End Function

Public Function SSLineStyleDlg%(SS As Object)

   On Error GoTo SSLineStyleDlgError
   SS.LineStyleDlg
   SSLineStyleDlg = 0
   Exit Function

SSLineStyleDlgError:
   SSLineStyleDlg = Err.Number

End Function

Public Function SSMoveRange%(SS As Object, ByVal nR1%, ByVal nC1%, ByVal nR2%, ByVal nC2%, ByVal nRowOffset%, ByVal nColOffset%)

   On Error GoTo SSMoveRangeError
   SS.MoveRange nR1, nC1, nR2, nC2, nRowOffset, nColOffset
   SSMoveRange = 0
   Exit Function

SSMoveRangeError:
   SSMoveRange = Err.Number

End Function

Public Function SSNextColPageBreak%(SS As Object, ByVal nCol%, pNextCol%)

   On Error GoTo SSNextColPageBreakError
   pNextCol = SS.NextColPageBreak(nCol)
   SSNextColPageBreak = 0
   Exit Function

SSNextColPageBreakError:
   SSNextColPageBreak = Err.Number

End Function

Public Function SSNextRowPageBreak%(SS As Object, ByVal nRow%, pNextRow%)

   On Error GoTo SSNextRowPageBreakError
   pNextRow = SS.NextRowPageBreak(nRow)
   SSNextRowPageBreak = 0
   Exit Function

SSNextRowPageBreakError:
   SSNextRowPageBreak = Err.Number

End Function

Public Function SSObjAddItem%(SS As Object, ByVal ID&, ByVal pItem$)

   On Error GoTo SSObjAddItemError
   SS.ObjAddItem ID, pItem
   SSObjAddItem = 0
   Exit Function

SSObjAddItemError:
   SSObjAddItem = Err.Number

End Function

Public Function SSObjAddSelection%(SS As Object, ByVal ID&)

   On Error GoTo SSObjAddSelectionError
   SS.ObjAddSelection ID
   SSObjAddSelection = 0
   Exit Function

SSObjAddSelectionError:
   SSObjAddSelection = Err.Number

End Function

Public Function SSObjBringToFront%(SS As Object)

   On Error GoTo SSObjBringToFrontError
   SS.ObjBringToFront
   SSObjBringToFront = 0
   Exit Function

SSObjBringToFrontError:
   SSObjBringToFront = Err.Number

End Function

Public Function SSObjDeleteItem%(SS As Object, ByVal ID&, ByVal nItem%)

   On Error GoTo SSObjDeleteItemError
   SS.ObjDeleteItem ID, nItem
   SSObjDeleteItem = 0
   Exit Function

SSObjDeleteItemError:
   SSObjDeleteItem = Err.Number

End Function

Public Function SSObjFirstID%(SS As Object, pFirstID&)

   On Error GoTo SSObjFirstIDError
   pFirstID = SS.ObjFirstID
   SSObjFirstID = 0
   Exit Function

SSObjFirstIDError:
   SSObjFirstID = Err.Number

End Function

Public Function SSObjGetCell%(SS As Object, ByVal ID&, pHasCell%, pRow%, pCol%)

   On Error GoTo SSObjGetCellError
   SS.ObjGetCell ID, pHasCell, pRow, pCol
   SSObjGetCell = 0
   Exit Function

SSObjGetCellError:
   SSObjGetCell = Err.Number

End Function

Public Function SSObjGetItem%(SS As Object, ByVal ID&, ByVal nItem%, ByRef pBuf$, ByVal nBufSize%)

   On Error GoTo SSObjGetItemError
   pBuf = SS.ObjItem(ID, nItem)
   
   SSObjGetItem = 0
   Exit Function

SSObjGetItemError:
   SSObjGetItem = Err.Number

End Function

Public Function SSObjGetItemCount%(SS As Object, ByVal ID&, pItems%)

   On Error GoTo SSObjGetItemCountError
   SS.ObjGetItemCount ID, pItems
   SSObjGetItemCount = 0
   Exit Function

SSObjGetItemCountError:
   SSObjGetItemCount = Err.Number

End Function

Public Function SSObjGetItems%(SS As Object, ByVal ID&, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSObjGetItemsError
   SS.ObjGetItems ID, pBuf
   SSObjGetItems = 0
   Exit Function

SSObjGetItemsError:
   SSObjGetItems = Err.Number

End Function

Public Function SSObjGetName%(SS As Object, ByVal ID&, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSObjGetNameError
   SS.ObjGetName ID, pBuf
   SSObjGetName = 0
   Exit Function

SSObjGetNameError:
   SSObjGetName = Err.Number

End Function

Public Function SSObjGetPos%(SS As Object, ByVal ID&, pX1!, pY1!, pX2!, pY2!)

   On Error GoTo SSObjGetPosError
   SS.ObjGetPos ID, pX1, pY1, pX2, pY2
   SSObjGetPos = 0
   Exit Function

SSObjGetPosError:
   SSObjGetPos = Err.Number

End Function

Public Function SSObjGetSelection%(SS As Object, ByVal nSelection%, pID&)

   On Error GoTo SSObjGetSelectionError
   SS.ObjGetSelection nSelection, pID
   SSObjGetSelection = 0
   Exit Function

SSObjGetSelectionError:
   SSObjGetSelection = Err.Number

End Function

Public Function SSObjGetSelectionCount%(SS As Object, pCount%)

   On Error GoTo SSObjGetSelectionCountError
   pCount = SS.ObjGetSelectionCount
   SSObjGetSelectionCount = 0
   Exit Function

SSObjGetSelectionCountError:
   SSObjGetSelectionCount = Err.Number

End Function

Public Function SSObjGetText%(SS As Object, ByVal ID&, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSObjGetTextError
   SS.ObjGetText ID, pBuf, nBufSize
   SSObjGetText = 0
   Exit Function

SSObjGetTextError:
   SSObjGetText = Err.Number

End Function

Public Function SSObjGetType%(SS As Object, ByVal ID&, pType%)

   On Error GoTo SSObjGetTypeError
   SS.ObjGetType ID, pType
   SSObjGetType = 0
   Exit Function

SSObjGetTypeError:
   SSObjGetType = Err.Number

End Function

Public Function SSObjGetValue%(SS As Object, ByVal ID&, pValue%)

   On Error GoTo SSObjGetValueError
   SS.ObjGetValue ID, pValue
   SSObjGetValue = 0
   Exit Function

SSObjGetValueError:
   SSObjGetValue = Err.Number

End Function

Public Function SSObjGetVisible%(SS As Object, ByVal ID&, pVisible%)

   On Error GoTo SSObjGetVisibleError
   SS.ObjGetVisible ID, pVisible
   SSObjGetVisible = 0
   Exit Function

SSObjGetVisibleError:
   SSObjGetVisible = Err.Number

End Function

Public Function SSObjInsertItem%(SS As Object, ByVal ID&, ByVal nItem%, ByVal pItem$)

   On Error GoTo SSObjInsertItemError
   SS.ObjInsertItem ID, nItem, pItem
   SSObjInsertItem = 0
   Exit Function

SSObjInsertItemError:
   SSObjInsertItem = Err.Number

End Function

Public Function SSObjNameDlg%(SS As Object)

   On Error GoTo SSObjNameDlgError
   SS.ObjNameDlg
   SSObjNameDlg = 0
   Exit Function

SSObjNameDlgError:
   SSObjNameDlg = Err.Number

End Function

Public Function SSObjNameToID%(SS As Object, ByVal pName$, pID&)

   On Error GoTo SSObjNameToIDError
   pID = SS.ObjNameToID(pName)
   SSObjNameToID = 0
   Exit Function

SSObjNameToIDError:
   SSObjNameToID = Err.Number

End Function

Public Function SSObjNew%(SS As Object, ByVal nType%, ByVal nX1!, ByVal nY1!, ByVal nX2!, ByVal nY2!, pID&)

   On Error GoTo SSObjNewError
   SS.ObjNew nType, nX1, nY1, nX2, nY2, pID
   SSObjNew = 0
   Exit Function

SSObjNewError:
   SSObjNew = Err.Number

End Function

Public Function SSObjNewPicture%(SS As Object, ByVal nX1!, ByVal nY1!, ByVal nX2!, ByVal nY2!, pID&, ByVal hMF%, ByVal nMapMode%, ByVal nWndExtentX%, ByVal nWndExtentY%)

   On Error GoTo SSObjNewPictureError
   SS.ObjNewPicture nX1, nY1, nX2, nY2, pID, hMF, nMapMode, nWndExtentX, nWndExtentY
   SSObjNewPicture = 0
   Exit Function

SSObjNewPictureError:
   SSObjNewPicture = Err.Number

End Function

Public Function SSObjNextID%(SS As Object, ByVal ID&, pNextID&)

   On Error GoTo SSObjNextIDError
   pNextID = SS.ObjNextID(ID)
   SSObjNextID = 0
   Exit Function

SSObjNextIDError:
   SSObjNextID = Err.Number

End Function

Public Function SSObjOptionsDlg%(SS As Object)

   On Error GoTo SSObjOptionsDlgError
   SS.ObjOptionsDlg
   SSObjOptionsDlg = 0
   Exit Function

SSObjOptionsDlgError:
   SSObjOptionsDlg = Err.Number

End Function

Public Function SSObjPosToTwips%(SS As Object, ByVal nX1!, ByVal nY1!, ByVal nX2!, ByVal nY2!, pX&, pY&, pCX&, pCY&, pShown%)

   On Error GoTo SSObjPosToTwipsError
   SS.ObjPosToTwips nX1, nY1, nX2, nY2, pX, pY, pCX, pCY, pShown
   SSObjPosToTwips = 0
   Exit Function

SSObjPosToTwipsError:
   SSObjPosToTwips = Err.Number

End Function

Public Function SSObjSendToBack%(SS As Object)

   On Error GoTo SSObjSendToBackError
   SS.ObjSendToBack
   SSObjSendToBack = 0
   Exit Function

SSObjSendToBackError:
   SSObjSendToBack = Err.Number

End Function

Public Function SSObjSetCell%(SS As Object, ByVal ID&, ByVal bHasCell%, ByVal nRow%, ByVal nCol%)

   On Error GoTo SSObjSetCellError
   SS.ObjSetCell ID, bHasCell, nRow, nCol
   SSObjSetCell = 0
   Exit Function

SSObjSetCellError:
   SSObjSetCell = Err.Number

End Function

Public Function SSObjSetItem%(SS As Object, ByVal ID&, ByVal nItem%, ByVal pItem$)

   On Error GoTo SSObjSetItemError
   SS.ObjSetItem ID, nItem, pItem
   SSObjSetItem = 0
   Exit Function

SSObjSetItemError:
   SSObjSetItem = Err.Number

End Function

Public Function SSObjSetItems%(SS As Object, ByVal ID&, ByVal pItems$)

   On Error GoTo SSObjSetItemsError
   SS.ObjSetItems ID, pItems
   SSObjSetItems = 0
   Exit Function

SSObjSetItemsError:
   SSObjSetItems = Err.Number

End Function

Public Function SSObjSetName%(SS As Object, ByVal ID&, ByVal pName$)

   On Error GoTo SSObjSetNameError
   pName = SS.ObjName(ID)
   SSObjSetName = 0
   Exit Function

SSObjSetNameError:
   SSObjSetName = Err.Number

End Function

Public Function SSObjSetPicture%(SS As Object, ByVal ID&, ByVal hMF%, ByVal nMapMode%, ByVal nWndExtentX%, ByVal nWndExtentY%)

   On Error GoTo SSObjSetPictureError
   SS.ObjSetPicture ID, hMF, nMapMode, nWndExtentX, nWndExtentY
   SSObjSetPicture = 0
   Exit Function

SSObjSetPictureError:
   SSObjSetPicture = Err.Number

End Function

Public Function SSObjSetPos%(SS As Object, ByVal ID&, ByVal nX1!, ByVal nY1!, ByVal nX2!, ByVal nY2!)

   On Error GoTo SSObjSetPosError
   SS.ObjSetPos ID, nX1, nY1, nX2, nY2
   SSObjSetPos = 0
   Exit Function

SSObjSetPosError:
   SSObjSetPos = Err.Number

End Function

Public Function SSObjSetSelection%(SS As Object, ByVal ID&)

   On Error GoTo SSObjSetSelectionError
   SS.ObjSetSelection ID
   SSObjSetSelection = 0
   Exit Function

SSObjSetSelectionError:
   SSObjSetSelection = Err.Number

End Function

Public Function SSObjSetText%(SS As Object, ByVal ID&, ByVal pText$)

   On Error GoTo SSObjSetTextError
   SS.ObjSetText(ID) = pText
   SSObjSetText = 0
   Exit Function

SSObjSetTextError:
   SSObjSetText = Err.Number

End Function

Public Function SSObjSetValue%(SS As Object, ByVal ID&, ByVal nValue%)

   On Error GoTo SSObjSetValueError
   SS.ObjValue(ID) = nValue
   SSObjSetValue = 0
   Exit Function

SSObjSetValueError:
   SSObjSetValue = Err.Number

End Function

Public Function SSObjSetVisible%(SS As Object, ByVal ID&, ByVal bVisible%)

   On Error GoTo SSObjSetVisibleError
   SS.ObjVisible(ID) = bVisible
   SSObjSetVisible = 0
   Exit Function

SSObjSetVisibleError:
   SSObjSetVisible = Err.Number

End Function

Public Function SSOpenFileDlg%(SS As Object, ByVal pTitle$, ByVal hWndParent%, ByVal pBuf$, ByVal nBufSize%)

   On Error GoTo SSOpenFileDlgError
   SS.OpenFileDlg pTitle, hWndParent, pBuf
   SSOpenFileDlg = 0
   Exit Function

SSOpenFileDlgError:
   SSOpenFileDlg = Err.Number

End Function

Public Function SSProtectionDlg%(SS As Object)

   On Error GoTo SSProtectionDlgError
   SS.ProtectionDlg
   SSProtectionDlg = 0
   Exit Function

SSProtectionDlgError:
   SSProtectionDlg = Err.Number

End Function

Public Function SSRangeToTwips%(SS As Object, ByVal nRow1%, ByVal nCol1%, ByVal nRow2%, ByVal nCol2%, pX&, pY&, pCX&, pCY&, pShown%)

   On Error GoTo SSRangeToTwipsError
   SS.RangeToTwips nRow1, nCol1, nRow2, nCol2, pX, pY, pCX, pCY, pShown
   SSRangeToTwips = 0
   Exit Function

SSRangeToTwipsError:
   SSRangeToTwips = Err.Number

End Function

Public Function SSRead%(SS As Object, ByVal pPathName$, pFileType%)

   On Error GoTo SSReadError
   SS.Read pPathName, pFileType
   SSRead = 0
   Exit Function

SSReadError:
   SSRead = Err.Number

End Function

Public Function SSRecalc%(SS As Object)

   On Error GoTo SSRecalcError
   SS.Recalc
   SSRecalc = 0
   Exit Function

SSRecalcError:
   SSRecalc = Err.Number

End Function

Public Function SSRemoveColPageBreak%(SS As Object, ByVal nCol%)

   On Error GoTo SSRemoveColPageBreakError
   SS.RemoveColPageBreak (nCol)
   SSRemoveColPageBreak = 0
   Exit Function

SSRemoveColPageBreakError:
   SSRemoveColPageBreak = Err.Number

End Function

Public Function SSRemovePageBreak%(SS As Object)

   On Error GoTo SSRemovePageBreakError
   SS.RemovePageBreak
   SSRemovePageBreak = 0
   Exit Function

SSRemovePageBreakError:
   SSRemovePageBreak = Err.Number

End Function

Public Function SSRemoveRowPageBreak%(SS As Object, ByVal nRow%)

   On Error GoTo SSRemoveRowPageBreakError
   SS.RemoveRowPageBreak (nRow)
   SSRemoveRowPageBreak = 0
   Exit Function

SSRemoveRowPageBreakError:
   SSRemoveRowPageBreak = Err.Number

End Function

Public Function SSRowHeightDlg%(SS As Object)

   On Error GoTo SSRowHeightDlgError
   SS.RowHeightDlg
   SSRowHeightDlg = 0
   Exit Function

SSRowHeightDlgError:
   SSRowHeightDlg = Err.Number

End Function

Public Function SSSaveFileDlg%(SS As Object, ByVal pTitle$, ByVal pBuf$, ByVal nBufSize%, pFileType%)

   On Error GoTo SSSaveFileDlgError
   SS.SaveFileDlg pTitle, pBuf, pFileType
   SSSaveFileDlg = 0
   Exit Function

SSSaveFileDlgError:
   SSSaveFileDlg = Err.Number

End Function

Public Function SSSaveWindowInfo%(SS As Object)

   On Error GoTo SSSaveWindowInfoError
   SS.SaveWindowInfo
   SSSaveWindowInfo = 0
   Exit Function

SSSaveWindowInfoError:
   SSSaveWindowInfo = Err.Number

End Function

Public Function SSSetActiveCell%(SS As Object, ByVal nRow%, ByVal nCol%)

   On Error GoTo SSSetActiveCellError
   SS.SetActiveCell nRow, nCol
   SSSetActiveCell = 0
   Exit Function

SSSetActiveCellError:
   SSSetActiveCell = Err.Number

End Function

Public Function SSSetAlignment%(SS As Object, ByVal nHorizontal%, ByVal bWordWrap%, ByVal nVertical%, ByVal nOrientation%)

   On Error GoTo SSSetAlignmentError
   SS.SetAlignment nHorizontal, bWordWrap, nVertical, nOrientation
   SSSetAlignment = 0
   Exit Function

SSSetAlignmentError:
   SSSetAlignment = Err.Number

End Function

Public Function SSSetAllowArrows%(SS As Object, ByVal bAllowArrows%)

   On Error GoTo SSSetAllowArrowsError
   SS.AllowArrows = bAllowArrows
   SSSetAllowArrows = 0
   Exit Function

SSSetAllowArrowsError:
   SSSetAllowArrows = Err.Number

End Function

Public Function SSSetAllowDelete%(SS As Object, ByVal bAllowDelete%)

   On Error GoTo SSSetAllowDeleteError
   SS.AllowDelete = bAllowDelete
   SSSetAllowDelete = 0
   Exit Function

SSSetAllowDeleteError:
   SSSetAllowDelete = Err.Number

End Function

Public Function SSSetAllowEditHeaders%(SS As Object, ByVal bAllowEditHeaders%)

   On Error GoTo SSSetAllowEditHeadersError
   SS.AllowEditHeaders = bAllowEditHeaders
   SSSetAllowEditHeaders = 0
   Exit Function

SSSetAllowEditHeadersError:
   SSSetAllowEditHeaders = Err.Number

End Function

Public Function SSSetAllowFillRange%(SS As Object, ByVal bAllowFillRange%)

   On Error GoTo SSSetAllowFillRangeError
   SS.AllowFillRange = bAllowFillRange
   SSSetAllowFillRange = 0
   Exit Function

SSSetAllowFillRangeError:
   SSSetAllowFillRange = Err.Number

End Function

Public Function SSSetAllowInCellEditing%(SS As Object, ByVal bAllowInCellEditing%)

   On Error GoTo SSSetAllowInCellEditingError
   SS.AllowInCellEditing = bAllowInCellEditing
   SSSetAllowInCellEditing = 0
   Exit Function

SSSetAllowInCellEditingError:
   SSSetAllowInCellEditing = Err.Number

End Function

Public Function SSSetAllowMoveRange%(SS As Object, ByVal bAllowMoveRange%)

   On Error GoTo SSSetAllowMoveRangeError
   SS.AllowMoveRange = bAllowMoveRange
   SSSetAllowMoveRange = 0
   Exit Function

SSSetAllowMoveRangeError:
   SSSetAllowMoveRange = Err.Number

End Function

Public Function SSSetAllowObjSelections%(SS As Object, ByVal bAllowObjSelections%)

   On Error GoTo SSSetAllowObjSelectionsError
   SS.AllowObjSelections = bAllowObjSelections
   SSSetAllowObjSelections = 0
   Exit Function

SSSetAllowObjSelectionsError:
   SSSetAllowObjSelections = Err.Number

End Function

Public Function SSSetAllowResize%(SS As Object, ByVal bAllowResize%)

   On Error GoTo SSSetAllowResizeError
   SS.AllowResize = bAllowResize
   SSSetAllowResize = 0
   Exit Function

SSSetAllowResizeError:
   SSSetAllowResize = Err.Number

End Function

Public Function SSSetAllowSelections%(SS As Object, ByVal bAllowSelections%)

   On Error GoTo SSSetAllowSelectionsError
   SS.AllowSelections = bAllowSelections
   SSSetAllowSelections = 0
   Exit Function

SSSetAllowSelectionsError:
   SSSetAllowSelections = Err.Number

End Function

Public Function SSSetAllowTabs%(SS As Object, ByVal bAllowTabs%)

   On Error GoTo SSSetAllowTabsError
   SS.AllowTabs = bAllowTabs
   SSSetAllowTabs = 0
   Exit Function

SSSetAllowTabsError:
   SSSetAllowTabs = Err.Number

End Function

Public Function SSSetAllowFormulas%(SS As Object, ByVal bAllowFormulas%)

   On Error GoTo SSSetAllowFormulasError
   SS.AllowFormulas = bAllowFormulas
   SSSetAllowFormulas = 0
   Exit Function

SSSetAllowFormulasError:
   SSSetAllowFormulas = Err.Number

End Function

Public Function SSSetAutoRecalc%(SS As Object, ByVal bAutoRecalc%)

   On Error GoTo SSSetAutoRecalcError
   SS.AutoRecalc = bAutoRecalc
   SSSetAutoRecalc = 0
   Exit Function

SSSetAutoRecalcError:
   SSSetAutoRecalc = Err.Number

End Function

Public Function SSSetAppName%(SS As Object, ByVal pAppName$)

   On Error GoTo SSSetAppNameError
   SS.APPNAME = pAppName
   SSSetAppName = 0
   Exit Function

SSSetAppNameError:
   SSSetAppName = Err.Number

End Function

Public Function SSSetBackColor%(SS As Object, ByVal crBackColor&)

   On Error GoTo SSSetBackColorError
   SS.BackColor = crBackColor
   SSSetBackColor = 0
   Exit Function

SSSetBackColorError:
   SSSetBackColor = Err.Number

End Function

Public Function SSSetBorder%(SS As Object, ByVal nOutline%, ByVal nLeft%, ByVal nRight%, ByVal nTop%, ByVal nBottom%, ByVal nShade%, ByVal crOutline&, ByVal crLeft&, ByVal crRight&, ByVal crTop&, ByVal crBottom&)

   On Error GoTo SSSetBorderError
   SS.SetBorder nOutline, nLeft, nRight, nTop, nBottom, nShade, crOutline, crLeft, crRight, crTop, crBottom
   SSSetBorder = 0
   Exit Function

SSSetBorderError:
   SSSetBorder = Err.Number

End Function

Public Function SSSetColText%(SS As Object, ByVal nCol%, ByVal pColText$)

   On Error GoTo SSSetColTextError
   SS.ColText(nCol) = pColText
   SSSetColText = 0
   Exit Function

SSSetColTextError:
   SSSetColText = Err.Number

End Function

Public Function SSSetColWidth%(SS As Object, ByVal nC1%, ByVal nC2%, ByVal nWidth%, ByVal bDefColWidth%)

   On Error GoTo SSSetColWidthError
   SS.SetColWidth nC1, nC2, nWidth, bDefColWidth
   SSSetColWidth = 0
   Exit Function

SSSetColWidthError:
   SSSetColWidth = Err.Number

End Function

Public Function SSSetColWidthAuto%(SS As Object, ByVal nR1%, ByVal nC1%, ByVal nR2%, ByVal nC2%, ByVal bSetDefaults%)

   On Error GoTo SSSetColWidthAutoError
   SS.SetColWidthAuto nR1, nC1, nR2, nC2, bSetDefaults
   SSSetColWidthAuto = 0
   Exit Function

SSSetColWidthAutoError:
   SSSetColWidthAuto = Err.Number

End Function

Public Function SSSetDefaultFont%(SS As Object, ByVal pName$, ByVal nSize%)

   On Error GoTo SSSetDefaultFontError
   SS.SetDefaultFont pName, nSize
   SSSetDefaultFont = 0
   Exit Function

SSSetDefaultFontError:
   SSSetDefaultFont = Err.Number

End Function

Public Function SSSetDefinedName%(SS As Object, ByVal pName$, ByVal pFormula$)

   On Error GoTo SSSetDefinedNameError
   SS.SetDefinedName(pName) = pFormula
   SSSetDefinedName = 0
   Exit Function

SSSetDefinedNameError:
   SSSetDefinedName = Err.Number

End Function

Public Function SSSetEnableProtection%(SS As Object, ByVal bEnableProtection%)

   On Error GoTo SSSetEnableProtectionError
   SS.EnableProtection = bEnableProtection
   SSSetEnableProtection = 0
   Exit Function

SSSetEnableProtectionError:
   SSSetEnableProtection = Err.Number

End Function

Public Function SSSetEnterMovesDown%(SS As Object, ByVal bEnterMovesDown%)

   On Error GoTo SSSetEnterMovesDownError
   SS.EnterMovesDown = bEnterMovesDown
   SSSetEnterMovesDown = 0
   Exit Function

SSSetEnterMovesDownError:
   SSSetEnterMovesDown = Err.Number

End Function

Public Function SSSetEntry%(SS As Object, ByVal pEntry$)

   On Error GoTo SSSetEntryError
   SS.Entry = pEntry
   SSSetEntry = 0
   Exit Function

SSSetEntryError:
   SSSetEntry = Err.Number

End Function

Public Function SSSetEntryRC%(SS As Object, ByVal nRow%, ByVal nCol%, ByVal pEntry$)

   On Error GoTo SSSetEntryRCError
   SS.EntryRC(nRow, nCol) = pEntry
   SSSetEntryRC = 0
   Exit Function

SSSetEntryRCError:
   SSSetEntryRC = Err.Number

End Function

Public Function SSSetExtraColor%(SS As Object, ByVal crExtraColor&)

   On Error GoTo SSSetExtraColorError
   SS.ExtraColor = crExtraColor
   SSSetExtraColor = 0
   Exit Function

SSSetExtraColorError:
   SSSetExtraColor = Err.Number

End Function

Public Function SSSetFixedCols%(SS As Object, ByVal nC1%, ByVal nCols%)

   On Error GoTo SSSetFixedColsError
   SS.FixedCol = nC1
   SS.FixedCols = nCols
   SSSetFixedCols = 0
   Exit Function

SSSetFixedColsError:
   SSSetFixedCols = Err.Number

End Function

Public Function SSSetFixedRows%(SS As Object, ByVal nR1%, ByVal nRows%)

   On Error GoTo SSSetFixedRowsError
   SS.FixedRow = nR1
   SS.FixedRows = nRows
   SSSetFixedRows = 0
   Exit Function

SSSetFixedRowsError:
   SSSetFixedRows = Err.Number

End Function

Public Function SSSetFont%(SS As Object, ByVal pName$, ByVal nSize%, ByVal bBold%, ByVal bItalic%, ByVal bUnderline%, ByVal bStrikeout%, ByVal crColor&, ByVal bOutline%, ByVal bShadow%)

   On Error GoTo SSSetFontError
   SS.SetFont pName, nSize, bBold, bItalic, bUnderline, bStrikeout, crColor, bOutline, bShadow
   SSSetFont = 0
   Exit Function

SSSetFontError:
   SSSetFont = Err.Number

End Function

Public Function SSSetFormula%(SS As Object, ByVal pFormula$)

   On Error GoTo SSSetFormulaError
   SS.Formula = pFormula
   SSSetFormula = 0
   Exit Function

SSSetFormulaError:
   SSSetFormula = Err.Number

End Function

Public Function SSSetFormulaRC%(SS As Object, ByVal nRow%, ByVal nCol%, ByVal pFormula$)

   On Error GoTo SSSetFormulaRCError
   SS.FormulaRC(nRow, nCol) = pFormula
   SSSetFormulaRC = 0
   Exit Function

SSSetFormulaRCError:
   SSSetFormulaRC = Err.Number

End Function

Public Function SSSetHdrHeight%(SS As Object, ByVal nHeight%)

   On Error GoTo SSSetHdrHeightError
   SS.HdrHeight = nHeight
   SSSetHdrHeight = 0
   Exit Function

SSSetHdrHeightError:
   SSSetHdrHeight = Err.Number

End Function

Public Function SSSetHdrSelection%(SS As Object, ByVal bTopLeftHdr%, ByVal bRowHdr%, ByVal bColHdr%)

   On Error GoTo SSSetHdrSelectionError
   SS.SetHdrSelection bTopLeftHdr, bRowHdr, bColHdr
   SSSetHdrSelection = 0
   Exit Function

SSSetHdrSelectionError:
   SSSetHdrSelection = Err.Number

End Function

Public Function SSSetHdrWidth%(SS As Object, ByVal nWidth%)

   On Error GoTo SSSetHdrWidthError
   SS.HdrWidth = nWidth
   SSSetHdrWidth = 0
   Exit Function

SSSetHdrWidthError:
   SSSetHdrWidth = Err.Number

End Function

Public Function SSSetIteration%(SS As Object, ByVal bIteration%, ByVal nMaxIterations%, ByVal nMaxChange#)

   On Error GoTo SSSetIterationError
   SS.SetIteration bIteration, nMaxIterations, nMaxChange
   SSSetIteration = 0
   Exit Function

SSSetIterationError:
   SSSetIteration = Err.Number

End Function

Public Function SSSetLeftCol%(SS As Object, ByVal nLeftCol%)

   On Error GoTo SSSetLeftColError
   SS.LeftCol = nLeftCol
   SSSetLeftCol = 0
   Exit Function

SSSetLeftColError:
   SSSetLeftCol = Err.Number

End Function

Public Function SSSetLineStyle%(SS As Object, ByVal nStyle%, ByVal crColor&, ByVal nWeight%)

   On Error GoTo SSSetLineStyleError
   SS.SetLineStyle nStyle, crColor, nWeight
   SSSetLineStyle = 0
   Exit Function

SSSetLineStyleError:
   SSSetLineStyle = Err.Number

End Function

Public Function SSSetLogicalRC%(SS As Object, ByVal nRow%, ByVal nCol%, ByVal bIsTrue%)

   On Error GoTo SSSetLogicalRCError
   SS.LogicalRC(nRow, nCol) = bIsTrue
   SSSetLogicalRC = 0
   Exit Function

SSSetLogicalRCError:
   SSSetLogicalRC = Err.Number

End Function

Public Function SSSetMinCol%(SS As Object, ByVal nMinCol%)

   On Error GoTo SSSetMinColError
   SS.MinCol = nMinCol
   SSSetMinCol = 0
   Exit Function

SSSetMinColError:
   SSSetMinCol = Err.Number

End Function

Public Function SSSetMinRow%(SS As Object, ByVal nMinRow%)

   On Error GoTo SSSetMinRowError
   SS.MinRow = nMinRow
   SSSetMinRow = 0
   Exit Function

SSSetMinRowError:
   SSSetMinRow = Err.Number

End Function

Public Function SSSetMaxCol%(SS As Object, ByVal nMaxCol%)

   On Error GoTo SSSetMaxColError
   SS.MaxCol = nMaxCol
   SSSetMaxCol = 0
   Exit Function

SSSetMaxColError:
   SSSetMaxCol = Err.Number

End Function

Public Function SSSetMaxRow%(SS As Object, ByVal nMaxRow%)

   On Error GoTo SSSetMaxRowError
   SS.MaxRow = nMaxRow
   SSSetMaxRow = 0
   Exit Function

SSSetMaxRowError:
   SSSetMaxRow = Err.Number

End Function

Public Function SSSetMode%(SS As Object, ByVal nMode%)

   On Error GoTo SSSetModeError
   SS.Mode = nMode
   SSSetMode = 0
   Exit Function

SSSetModeError:
   SSSetMode = Err.Number

End Function

Public Function SSSetNumber%(SS As Object, ByVal nNumber#)

   On Error GoTo SSSetNumberError
   SS.Number = nNumber
   SSSetNumber = 0
   Exit Function

SSSetNumberError:
   SSSetNumber = Err.Number

End Function

Public Function SSSetNumberFormat%(SS As Object, ByVal pNumberFormat$)

   On Error GoTo SSSetNumberFormatError
   SS.NumberFormat = pNumberFormat
   SSSetNumberFormat = 0
   Exit Function

SSSetNumberFormatError:
   SSSetNumberFormat = Err.Number

End Function

Public Function SSSetNumberRC%(SS As Object, ByVal nRow%, ByVal nCol%, ByVal nNumber#)

   On Error GoTo SSSetNumberRCError
   SS.NumberRC(nRow, nCol) = nNumber
   SSSetNumberRC = 0
   Exit Function

SSSetNumberRCError:
   SSSetNumberRC = Err.Number

End Function

Public Function SSSetPattern%(SS As Object, ByVal nPattern%, ByVal crFG&, ByVal crBG&)

   On Error GoTo SSSetPatternError
   SS.SetPattern nPattern, crFG, crBG
   SSSetPattern = 0
   Exit Function

SSSetPatternError:
   SSSetPattern = Err.Number

End Function

Public Function SSSetPolyEditMode%(SS As Object, ByVal nPolyEditMode%)

   On Error GoTo SSSetPolyEditModeError
   SS.PolyEditMode = nPolyEditMode
   SSSetPolyEditMode = 0
   Exit Function

SSSetPolyEditModeError:
   SSSetPolyEditMode = Err.Number

End Function

Public Function SSSetPrintArea%(SS As Object, ByVal pFormula$)

   On Error GoTo SSSetPrintAreaError
   SS.PrintArea = pFormula
   SSSetPrintArea = 0
   Exit Function

SSSetPrintAreaError:
   SSSetPrintArea = Err.Number

End Function

Public Function SSSetPrintAreaFromSelection%(SS As Object)

   On Error GoTo SSSetPrintAreaFromSelectionError
   SS.SetPrintAreaFromSelection
   SSSetPrintAreaFromSelection = 0
   Exit Function

SSSetPrintAreaFromSelectionError:
   SSSetPrintAreaFromSelection = Err.Number

End Function

Public Function SSSetPrintBottomMargin%(SS As Object, ByVal nMargin#)

   On Error GoTo SSSetPrintBottomMarginError
   SS.PrintBottomMargin = nMargin
   SSSetPrintBottomMargin = 0
   Exit Function

SSSetPrintBottomMarginError:
   SSSetPrintBottomMargin = Err.Number

End Function

Public Function SSSetPrintColHeading%(SS As Object, ByVal bColHeading%)

   On Error GoTo SSSetPrintColHeadingError
   SS.PrintColHeading = bColHeading
   SSSetPrintColHeading = 0
   Exit Function

SSSetPrintColHeadingError:
   SSSetPrintColHeading = Err.Number

End Function

Public Function SSSetPrintFooter%(SS As Object, ByVal pPrintFooter$)

   On Error GoTo SSSetPrintFooterError
   SS.PrintFooter = pPrintFooter
   SSSetPrintFooter = 0
   Exit Function

SSSetPrintFooterError:
   SSSetPrintFooter = Err.Number

End Function

Public Function SSSetPrintGridLines%(SS As Object, ByVal bGridLines%)

   On Error GoTo SSSetPrintGridLinesError
   SS.PrintGridLines = bGridLines
   SSSetPrintGridLines = 0
   Exit Function

SSSetPrintGridLinesError:
   SSSetPrintGridLines = Err.Number

End Function

Public Function SSSetPrintHCenter%(SS As Object, ByVal bHCenter%)

   On Error GoTo SSSetPrintHCenterError
   SS.PrintHCenter = bHCenter
   SSSetPrintHCenter = 0
   Exit Function

SSSetPrintHCenterError:
   SSSetPrintHCenter = Err.Number

End Function

Public Function SSSetPrintHeader%(SS As Object, ByVal pPrintHeader$)

   On Error GoTo SSSetPrintHeaderError
   SS.PrintHeader = pPrintHeader
   SSSetPrintHeader = 0
   Exit Function

SSSetPrintHeaderError:
   SSSetPrintHeader = Err.Number

End Function

Public Function SSSetPrintLandscape%(SS As Object, ByVal bLandscape%)

   On Error GoTo SSSetPrintLandscapeError
   SS.PrintLandscape = bLandscape
   SSSetPrintLandscape = 0
   Exit Function

SSSetPrintLandscapeError:
   SSSetPrintLandscape = Err.Number

End Function

Public Function SSSetPrintLeftMargin%(SS As Object, ByVal nMargin#)

   On Error GoTo SSSetPrintLeftMarginError
   SS.PrintLeftMargin = nMargin
   SSSetPrintLeftMargin = 0
   Exit Function

SSSetPrintLeftMarginError:
   SSSetPrintLeftMargin = Err.Number

End Function

Public Function SSSetPrintLeftToRight%(SS As Object, ByVal bLeftToRight%)

   On Error GoTo SSSetPrintLeftToRightError
   SS.PrintLeftToRight = bLeftToRight
   SSSetPrintLeftToRight = 0
   Exit Function

SSSetPrintLeftToRightError:
   SSSetPrintLeftToRight = Err.Number

End Function

Public Function SSSetPrintNoColor%(SS As Object, ByVal bNoColor%)

   On Error GoTo SSSetPrintNoColorError
   SS.PrintNoColor = bNoColor
   SSSetPrintNoColor = 0
   Exit Function

SSSetPrintNoColorError:
   SSSetPrintNoColor = Err.Number

End Function

Public Function SSSetPrintRightMargin%(SS As Object, ByVal nMargin#)

   On Error GoTo SSSetPrintRightMarginError
   SS.PrintRightMargin = nMargin
   SSSetPrintRightMargin = 0
   Exit Function

SSSetPrintRightMarginError:
   SSSetPrintRightMargin = Err.Number

End Function

Public Function SSSetPrintRowHeading%(SS As Object, ByVal bRowHeading%)

   On Error GoTo SSSetPrintRowHeadingError
   SS.PrintRowHeading = bRowHeading
   SSSetPrintRowHeading = 0
   Exit Function

SSSetPrintRowHeadingError:
   SSSetPrintRowHeading = Err.Number

End Function

Public Function SSSetPrintScale%(SS As Object, ByVal nScale%, ByVal bFitToPage%, ByVal nVPages%, ByVal nHPages%)

   On Error GoTo SSSetPrintScaleError
   SS.SetPrintScale nScale, bFitToPage, nVPages, nHPages
   SSSetPrintScale = 0
   Exit Function

SSSetPrintScaleError:
   SSSetPrintScale = Err.Number

End Function

Public Function SSSetPrintTitles%(SS As Object, ByVal pFormula$)

   On Error GoTo SSSetPrintTitlesError
   SS.PrintTitles = pFormula
   SSSetPrintTitles = 0
   Exit Function

SSSetPrintTitlesError:
   SSSetPrintTitles = Err.Number

End Function

Public Function SSSetPrintTitlesFromSelection%(SS As Object)

   On Error GoTo SSSetPrintTitlesFromSelectionError
   SS.SetPrintTitlesFromSelection
   SSSetPrintTitlesFromSelection = 0
   Exit Function

SSSetPrintTitlesFromSelectionError:
   SSSetPrintTitlesFromSelection = Err.Number

End Function

Public Function SSSetPrintTopMargin%(SS As Object, ByVal nMargin#)

   On Error GoTo SSSetPrintTopMarginError
   SS.PrintTopMargin = nMargin
   SSSetPrintTopMargin = 0
   Exit Function

SSSetPrintTopMarginError:
   SSSetPrintTopMargin = Err.Number

End Function

Public Function SSSetPrintVCenter%(SS As Object, ByVal bVCenter%)

   On Error GoTo SSSetPrintVCenterError
   SS.PrintVCenter = bVCenter
   SSSetPrintVCenter = 0
   Exit Function

SSSetPrintVCenterError:
   SSSetPrintVCenter = Err.Number

End Function

Public Function SSSetProtection%(SS As Object, ByVal bLocked%, ByVal bHidden%)

   On Error GoTo SSSetProtectionError
   SS.SetProtection bLocked, bHidden
   SSSetProtection = 0
   Exit Function

SSSetProtectionError:
   SSSetProtection = Err.Number

End Function

Public Function SSSetRepaint%(SS As Object, ByVal bRepaint%)

   On Error GoTo SSSetRepaintError
   SS.Repaint = bRepaint
   SSSetRepaint = 0
   Exit Function

SSSetRepaintError:
   SSSetRepaint = Err.Number

End Function

Public Function SSSetRowHeight%(SS As Object, ByVal nR1%, ByVal nR2%, ByVal nHeight%, ByVal bDefRowHeight%)

   On Error GoTo SSSetRowHeightError
   SS.SetRowHeight nR1, nR2, nHeight, bDefRowHeight
   SSSetRowHeight = 0
   Exit Function

SSSetRowHeightError:
   SSSetRowHeight = Err.Number

End Function

Public Function SSSetRowHeightAuto%(SS As Object, ByVal nR1%, ByVal nC1%, ByVal nR2%, ByVal nC2%, ByVal bSetDefaults%)

   On Error GoTo SSSetRowHeightAutoError
   SS.SetRowHeightAuto nR1, nC1, nR2, nC2, bSetDefaults
   SSSetRowHeightAuto = 0
   Exit Function

SSSetRowHeightAutoError:
   SSSetRowHeightAuto = Err.Number

End Function

Public Function SSSetRowMode%(SS As Object, ByVal bRowMode%)

   On Error GoTo SSSetRowModeError
   SS.RowMode = bRowMode
   SSSetRowMode = 0
   Exit Function

SSSetRowModeError:
   SSSetRowMode = Err.Number

End Function

Public Function SSSetRowText%(SS As Object, ByVal nRow%, ByVal pRowText$)

   On Error GoTo SSSetRowTextError
   SS.RowText(nRow) = pRowText
   SSSetRowText = 0
   Exit Function

SSSetRowTextError:
   SSSetRowText = Err.Number

End Function

Public Function SSSetScale%(SS As Object, ByVal nScale%)

   On Error GoTo SSSetScaleError
   SS.ViewScale = nScale
   SSSetScale = 0
   Exit Function

SSSetScaleError:
   SSSetScale = Err.Number

End Function

Public Function SSSetSelection%(SS As Object, ByVal nR1%, ByVal nC1%, ByVal nR2%, ByVal nC2%)

   On Error GoTo SSSetSelectionError
   SS.SetSelection nR1, nC1, nR2, nC2
   SSSetSelection = 0
   Exit Function

SSSetSelectionError:
   SSSetSelection = Err.Number

End Function

Public Function SSSetSelectionRef%(SS As Object, ByVal pFormula$)

   On Error GoTo SSSetSelectionRefError
   SS.SelectionRef = pFormula
   SSSetSelectionRef = 0
   Exit Function

SSSetSelectionRefError:
   SSSetSelectionRef = Err.Number

End Function

Public Function SSSetShowColHeading%(SS As Object, ByVal bColHeading%)

   On Error GoTo SSSetShowColHeadingError
   SS.ShowColHeading = bColHeading
   SSSetShowColHeading = 0
   Exit Function

SSSetShowColHeadingError:
   SSSetShowColHeading = Err.Number

End Function

Public Function SSSetShowFormulas%(SS As Object, ByVal bFormulas%)

   On Error GoTo SSSetShowFormulasError
   SS.ShowFormulas = bFormulas
   SSSetShowFormulas = 0
   Exit Function

SSSetShowFormulasError:
   SSSetShowFormulas = Err.Number

End Function

Public Function SSSetShowGridLines%(SS As Object, ByVal bGridLines%)

   On Error GoTo SSSetShowGridLinesError
   SS.ShowGridLines = bGridLines
   SSSetShowGridLines = 0
   Exit Function

SSSetShowGridLinesError:
   SSSetShowGridLines = Err.Number

End Function

Public Function SSSetShowHScrollBar%(SS As Object, ByVal nShowHScrollBar%)

   On Error GoTo SSSetShowHScrollBarError
   SS.ShowHScrollBar = nShowHScrollBar
   SSSetShowHScrollBar = 0
   Exit Function

SSSetShowHScrollBarError:
   SSSetShowHScrollBar = Err.Number

End Function

Public Function SSSetShowRowHeading%(SS As Object, ByVal bRowHeading%)

   On Error GoTo SSSetShowRowHeadingError
   SS.ShowRowHeading = bRowHeading
   SSSetShowRowHeading = 0
   Exit Function

SSSetShowRowHeadingError:
   SSSetShowRowHeading = Err.Number

End Function

Public Function SSSetShowSelections%(SS As Object, ByVal nSelections%)

   On Error GoTo SSSetShowSelectionsError
   SS.ShowSelections = nSelections
   SSSetShowSelections = 0
   Exit Function

SSSetShowSelectionsError:
   SSSetShowSelections = Err.Number

End Function

Public Function SSSetShowVScrollBar%(SS As Object, ByVal nShowVScrollBar%)

   On Error GoTo SSSetShowVScrollBarError
   SS.ShowVScrollBar = nShowVScrollBar
   SSSetShowVScrollBar = 0
   Exit Function

SSSetShowVScrollBarError:
   SSSetShowVScrollBar = Err.Number

End Function

Public Function SSSetShowZeroValues%(SS As Object, ByVal bZeroValues%)

   On Error GoTo SSSetShowZeroValuesError
   SS.ShowZeroValues = bZeroValues
   SSSetShowZeroValues = 0
   Exit Function

SSSetShowZeroValuesError:
   SSSetShowZeroValues = Err.Number

End Function

Public Function SSSetTabbedText%(SS As Object, ByVal nStartRow%, ByVal nStartCol%, pRows%, pCols%, ByVal bValuesOnly%, ByVal pText$)

   On Error GoTo SSSetTabbedTextError
   SS.SetTabbedText nStartRow, nStartCol, pRows, pCols, bValuesOnly, pText
   SSSetTabbedText = 0
   Exit Function

SSSetTabbedTextError:
   SSSetTabbedText = Err.Number

End Function

Public Function SSSetText%(SS As Object, ByVal pText$)

   On Error GoTo SSSetTextError
   SS.Text = pText
   SSSetText = 0
   Exit Function

SSSetTextError:
   SSSetText = Err.Number

End Function

Public Function SSSetTextRC%(SS As Object, ByVal nRow%, ByVal nCol%, ByVal pText$)

   On Error GoTo SSSetTextRCError
   SS.TextRC(nRow, nCol) = pText
   SSSetTextRC = 0
   Exit Function

SSSetTextRCError:
   SSSetTextRC = Err.Number

End Function

Public Function SSSetTitle%(SS As Object, ByVal pTitle$)

   On Error GoTo SSSetTitleError
   SS.Title = pTitle
   SSSetTitle = 0
   Exit Function

SSSetTitleError:
   SSSetTitle = Err.Number

End Function

Public Function SSSetTopLeftText%(SS As Object, ByVal pTopLeftText$)

   On Error GoTo SSSetTopLeftTextError
   SS.TopLeftText = pTopLeftText
   SSSetTopLeftText = 0
   Exit Function

SSSetTopLeftTextError:
   SSSetTopLeftText = Err.Number

End Function

Public Function SSSetTopRow%(SS As Object, ByVal nTopRow%)

   On Error GoTo SSSetTopRowError
   SS.TopRow = nTopRow
   SSSetTopRow = 0
   Exit Function

SSSetTopRowError:
   SSSetTopRow = Err.Number

End Function

Public Function SSShowActiveCell%(SS As Object)

   On Error GoTo SSShowActiveCellError
   SS.ShowActiveCell
   SSShowActiveCell = 0
   Exit Function

SSShowActiveCellError:
   SSShowActiveCell = Err.Number

End Function

Public Function SSSort3%(SS As Object, ByVal nR1%, ByVal nC1%, ByVal nR2%, ByVal nC2%, ByVal bSortByRows%, ByVal nKey1%, ByVal nKey2%, ByVal nKey3%)

   On Error GoTo SSSort3Error
   SS.Sort3 nR1, nC1, nR2, nC2, bSortByRows, nKey1, nKey2, nKey3
   SSSort3 = 0
   Exit Function

SSSort3Error:
   SSSort3 = Err.Number

End Function

Public Function SSSortDlg%(SS As Object)

   On Error GoTo SSSortDlgError
   SS.SortDlg
   SSSortDlg = 0
   Exit Function

SSSortDlgError:
   SSSortDlg = Err.Number

End Function

Public Function SSStartEdit%(SS As Object, ByVal bClear%, ByVal bInCellEditFocus%, ByVal bArrowsExitEditMode%)

   On Error GoTo SSStartEditError
   SS.StartEdit bClear, bInCellEditFocus, bArrowsExitEditMode
   SSStartEdit = 0
   Exit Function

SSStartEditError:
   SSStartEdit = Err.Number

End Function

Public Function SSSwapTables%(SS As Object, ByVal hSS2&)

   On Error GoTo SSSwapTablesError
   SS.SwapTables hSS2
   SSSwapTables = 0
   Exit Function

SSSwapTablesError:
   SSSwapTables = Err.Number

End Function

Public Function SSTransactCommit%(SS As Object)

   On Error GoTo SSTransactCommitError
   SS.TransactCommit
   SSTransactCommit = 0
   Exit Function

SSTransactCommitError:
   SSTransactCommit = Err.Number

End Function

Public Function SSTransactRollBack%(SS As Object)

   On Error GoTo SSTransactRollBackError
   SS.TransactRollback
   SSTransactRollBack = 0
   Exit Function

SSTransactRollBackError:
   SSTransactRollBack = Err.Number

End Function

Public Function SSTransactStart%(SS As Object)

   On Error GoTo SSTransactStartError
   SS.TransactStart
   SSTransactStart = 0
   Exit Function

SSTransactStartError:
   SSTransactStart = Err.Number

End Function

Public Function SSTwipsToRC%(SS As Object, ByVal X&, ByVal Y&, pRow%, pCol%)

   On Error GoTo SSTwipsToRCError
   SS.TwipsToRC X, Y, pRow, pCol
   SSTwipsToRC = 0
   Exit Function

SSTwipsToRCError:
   SSTwipsToRC = Err.Number

End Function

Public Sub SSUpdate(SS As Object)

   'SSUpdate
        
End Sub

Public Function SSVersion%(SS As Object)

   SSVersion = SSVersion

End Function

Public Function SSWrite%(SS As Object, ByVal pPathName$, ByVal nFileType%)

   On Error GoTo SSWriteError
   SS.Write pPathName, nFileType
   SSWrite = 0
   Exit Function

SSWriteError:
   SSWrite = Err.Number

End Function

