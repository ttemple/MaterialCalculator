Public Class frmMain
    Private Function SheetMetalRemnantYieldCalc(ByVal q As SheetMetalQuote) As SheetMetalJob
        Dim ppr As Decimal, rowLength As Decimal, maxRowsStockSheet As Integer
        Dim smJob As New SheetMetalJob

        Try
            '1.  Add frame to part dimensions
            q.dimA += q.FrameDimension
            q.dimB += q.FrameDimension

            rowLength = q.dimB + q.Workholding

            'Error checking
            'Determine if PartDimA exceeds StockDimA
            If q.dimA > q.StockDimA Then
                smJob.ShearLength = 0
                smJob.ShearSheetQty = 0
                smJob.FullSheetQty = 0
                smJob.Yield = 0
                Return smJob

            ElseIf rowLength > q.StockDimB Then
                'Determine if PartDimB exceeds StockDimB
                smJob.ShearLength = 0
                smJob.ShearSheetQty = 0
                smJob.FullSheetQty = 0
                smJob.Yield = 0
                Return smJob
            End If

            '2. Calc yield in parts per row (ppr)across StockDimA
            ppr = Math.Floor(q.StockDimA / q.dimA)
            maxRowsStockSheet = Math.Floor(q.StockDimB / rowLength)

            With smJob
                .Yield = ppr * maxRowsStockSheet
                .ShearSheetQty = 1
                .ShearLength = rowLength
                .FullSheetQty = 1
            End With

            Return smJob
        Catch ex As Exception
            MessageBox.Show("Your entry's have resulted in: " & ex.Message & "." & vbCr & "Try Again!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

    Private Function SheetMetalShearCalc(ByVal q As SheetMetalQuote) As SheetMetalJob
        Dim ppr As Decimal, numRows As Integer 'ppr = PartsPerRow
        Dim rowLength As Decimal, maxRowsStockSheet As Integer
        Dim smJob As New SheetMetalJob

        Try
            '1.  Add frame to part dimensions
            q.dimA += q.FrameDimension
            q.dimB += q.FrameDimension

            rowLength = q.dimB

            'Error checking
            'Determine if PartDimA exceeds StockDimA
            If q.dimA > q.StockDimA Then
                smJob.ShearLength = 0
                smJob.ShearSheetQty = 0
                smJob.FullSheetQty = 0
                smJob.Yield = 0
                Return smJob

            ElseIf rowLength > q.StockDimB Then
                'Determine if PartDimB exceeds StockDimB
                smJob.ShearLength = 0
                smJob.ShearSheetQty = 0
                smJob.FullSheetQty = 0
                smJob.Yield = 0
                Return smJob
            End If

            '2. Calc yield in parts per row (ppr)across StockDimA
            ppr = Math.Floor(q.StockDimA / q.dimA)
            maxRowsStockSheet = Math.Floor(q.StockDimB / rowLength)

            '3. Calc number of rows required in StockDimB to meet Qty
            numRows = Math.Ceiling(q.TotalQty / ppr)

            If maxRowsStockSheet = 1 Then 'If so, StockSheet will yield only one ShearSheet.

                numRows = Math.Ceiling(q.TotalQty / ppr) 'Each StockSheet yields only one row because partDimB+WorkHolding is over half length of StockSheet.
                'Calc yield per shear sheet

                'Assign values to SheetMetalCalc object
                smJob.ShearLength = rowLength
                smJob.ShearSheetQty = numRows
                smJob.FullSheetQty = numRows
                smJob.Yield = numRows * ppr
                Return smJob
            Else 'Calc max shear size, within restraint, to obtain multiple shear sheets of equal size.
                Dim x As Integer 'Counter
                Dim xQty As Integer, xRows As Integer, MaxShearLength As Decimal, xYield As Integer  'Alternate qty's

                xQty = q.TotalQty
                xRows = Math.Ceiling(xQty / ppr) 'Calc number of rows per ShearSheet of alt qty required. PPR does NOT change because stocksize same.
                MaxShearLength = Math.Round((numRows * rowLength) + q.Workholding, 3)

                Do While MaxShearLength > q.ShearSizeDimA 'Goal is to determine max yielding ShearSheet within restraint
                    x += 1 'Increment counter
                    xQty = Math.Ceiling(xQty / 2) 'Dividing total qty in half.
                    xRows = Math.Ceiling(xQty / ppr) 'Calc number of rows per ShearSheet of alt qty required. PPR does NOT change because stocksize same.
                    MaxShearLength = Math.Round((xRows * rowLength) + q.Workholding, 3) 'Calc alternate shear length
                    If xRows = 1 Then Exit Do
                Loop

                '4. Assign values to SheetMetalCalc object
                xYield = xRows * ppr 'Calc yield per shear sheet for alternate shear length
                smJob.ShearLength = MaxShearLength
                smJob.ShearSheetQty = CInt(Math.Ceiling(q.TotalQty / xYield)) 'Calc number of shear sheets required to meet TOTAL qty
                smJob.FullSheetQty = CInt(Math.Ceiling((MaxShearLength * smJob.ShearSheetQty) / q.StockDimB))
                smJob.Yield = xYield * smJob.ShearSheetQty

                Dim ssArea As Decimal = q.StockDimA * smJob.ShearLength
                Dim yieldPerSS As Integer = Math.Ceiling(smJob.Yield / smJob.ShearSheetQty)

                smJob.UnitPercentageShearSheet = Math.Round(((ssArea / yieldPerSS) / 144), 4)
                Return smJob

            End If

        Catch ex As Exception
            MessageBox.Show("Your entry's have resulted in: " & ex.Message & "." & vbCr & "Try Again!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function
    Private Function BarCalc(ByVal q As BarQuote) As BarJob
        'Returns the number of full length bars required to yield qty in BarJob
        Dim b As New BarJob

        With q
            b.PartOAL = .PartOAL + .CleanUpLength + .ExtraOpLength + .PartOffLength

            'Goal is to have equal length machine bars.
            b.PartYieldPerMachineBar = Math.Floor((.BarFeedLength - .BarEndLength) / b.PartOAL)
            b.MachineBarsPerRandom = Math.Floor(.RandomLength / .BarFeedLength)

            'Calculate minimum number of machine bars required to meet order qty.
            b.QtyMachineBarsReq = Math.Ceiling(q.Qty / b.PartYieldPerMachineBar)
            b.QtyRandomsRequired = Math.Ceiling((.BarFeedLength * b.QtyMachineBarsReq) / .RandomLength)

            b.GrossPartYield = b.QtyMachineBarsReq * b.PartYieldPerMachineBar
            b.PartYieldPerRandom = b.MachineBarsPerRandom * b.PartYieldPerMachineBar

            Return b
        End With
    End Function
    Private Function PlateCalc(ByVal q As PlateQuote) As PlateJob
        Dim p As New PlateJob
        With q
            p.DimST = .PartDimST + .ExtraDimST + .StdMatAdd
            p.DimLT = .PartDimLT + .ExtraDimLT + .StdMatAdd
            p.DimL = .PartDimL + .ExtraDimL + .StdMatAdd
        End With
        Return p
    End Function

    Private Function Process_SheetMetal() As SheetMetalPackage
        Dim qA As New SheetMetalQuote, qB As New SheetMetalQuote
        Dim smJobA As New SheetMetalJob, smJobB As New SheetMetalJob
        Dim smPkg As New SheetMetalPackage

        'Set properties of Quote object
        qA.dimA = CDec(txtSheetMetalDimA.Text)
        qA.dimB = CDec(txtSheetMetalDimB.Text)

        'If remnant, qty not necessary.
        If Not chkRemnant.Checked Then qA.TotalQty = CInt(txtSheetMetalQty.Text)

        qA.Workholding = CDec(txtWorkHolding.Text)
        qA.StockDimA = CDec(txtStockDimA.Text)
        qA.StockDimB = CDec(txtStockDimB.Text)
        qA.ShearSizeDimA = CDec(txtShearDimA.Text)
        qA.FrameDimension = CDec(txtFrame.Text)

        'Set properties of Quote object
        qB.dimB = CDec(txtSheetMetalDimA.Text) 'Swap dimA & dimB for alternate calc Add FrameDimension" 
        qB.dimA = CDec(txtSheetMetalDimB.Text)

        If Not chkRemnant.Checked Then qB.TotalQty = CInt(txtSheetMetalQty.Text)

        qB.Workholding = CDec(txtWorkHolding.Text)
        qB.StockDimA = CDec(txtStockDimA.Text)
        qB.StockDimB = CDec(txtStockDimB.Text)
        qB.ShearSizeDimA = CDec(txtShearDimA.Text)
        qB.FrameDimension = CDec(txtFrame.Text)

        If Not chkRemnant.Checked Then
            smJobA = SheetMetalShearCalc(qA)
        Else
            smJobA = SheetMetalRemnantYieldCalc(qA)
        End If


        If smJobA IsNot Nothing Then
            'Show 
            txtShearDimB.Text = "X " & CStr(qA.StockDimA)
            txtWorkHolding.Text = Format(qA.Workholding, "0.000")
            txtStockDimA.Text = Format(qA.StockDimA, "###.000")
            txtStockDimB.Text = Format(qA.StockDimB, "###.000")
            txtFrame.Text = Format(qA.FrameDimension, "0.##0")

            'Show results A
            txtShearSizeA.Text = CStr(smJobA.ShearLength) & " X " & CStr(qA.StockDimA)
            txtShearSheetCountA.Text = CStr(smJobA.ShearSheetQty)
            txtFullSheetCountA.Text = CStr(smJobA.FullSheetQty)
            txtYieldA.Text = Format(smJobA.Yield, "#,##0")
            txtShearSheetPctA.Text = Format(smJobA.UnitPercentageShearSheet, "0.0000")

            txtSheetMetalQty.SelectAll()
            txtSheetMetalQty.Focus()

            lblOption1.Text = "Option 1 - Calculated with Dim A along " & Format(qA.StockDimA, "###") & """ width."
        Else
            txtSheetMetalQty.Focus()
            Return Nothing
            Exit Function
        End If

        If Not chkRemnant.Checked Then
            smJobB = SheetMetalShearCalc(qB)
        Else
            smJobB = SheetMetalRemnantYieldCalc(qB)
        End If

        If smJobB IsNot Nothing Then
            'Show results B
            txtShearSizeB.Text = CStr(smJobB.ShearLength) & " X " & CStr(qB.StockDimA)
            txtShearSheetCountB.Text = CStr(smJobB.ShearSheetQty)
            txtFullSheetCountB.Text = CStr(smJobB.FullSheetQty)
            txtYieldB.Text = Format(smJobB.Yield, "#,##0")
            txtShearSheetPctB.Text = Format(smJobB.UnitPercentageShearSheet, "0.0000")

            txtSheetMetalQty.SelectAll()
            txtSheetMetalQty.Focus()

            lblOption2.Text = "Option 2 - Calculated with Dim B along " & Format(qB.StockDimA, "###") & """ width."

        End If

        EnableSheetMetalRadios()

        With smPkg
            .smJobA = smJobA
            .smJobB = smJobB
            .smQuoteA = qA
            .smQuoteB = qB
        End With

        Return smPkg

    End Function
    Private Sub Process_Bar()
        Dim bq As New BarQuote
        Dim bj As New BarJob

        'Set Bar object properties
        With bq
            .Qty = CInt(txtBarQty.Text)
            .PartOAL = CDec(txtNetPartOAL.Text)

            .PartOffLength = CDec(txtPartOffLength.Text)
            .ExtraOpLength = CDec(txtExtraOpLength.Text)
            .CleanUpLength = CDec(txtCleanUpLength.Text)
            .RandomLength = CDec(txtRandomLength.Text)
            .BarFeedLength = CDec(txtBarFeedLength.Text)
            .BarEndLength = CDec(txtBarEndLength.Text)

            bj = BarCalc(bq)

            txtRandomYield.Text = CStr(bj.PartYieldPerRandom)
            txtRequiredRndQty.Text = CStr(bj.QtyRandomsRequired)

            txtMachineBarYield.Text = CStr(bj.PartYieldPerMachineBar)
            txtMachineBarQty.Text = CStr(bj.QtyMachineBarsReq)

            txtPartOAL.Text = Format(bj.PartOAL, "0.##0")
            txtGrossPartYield.Text = Format(bj.GrossPartYield, "#,###,##0")

            rtxtBarDetails.Text = "NET OAL: " & Format(.PartOAL, "0.000") & "  YIELD: " & Format(.Qty, "#,##0") & _
                    " - RND_LN: " & CStr(.RandomLength) & """" & "  RND_QTY: " & Format(bj.QtyRandomsRequired, "#,##0") & vbCrLf & _
                    "PART OFF: " & Format(.PartOAL, "0.000") & " EXTRA OP: " & Format(.ExtraOpLength, "0.000") & vbCrLf & _
                    "CLEAN UP: " & Format(.CleanUpLength, "0.000") & " - PART OAL= " & Format(.PartOAL, "0.000")

            'Copy details to clipboard
            My.Computer.Clipboard.SetText(rtxtBarDetails.Text, TextDataFormat.Text)

        End With
    End Sub
    Private Sub Process_Plate()
        Dim pq As New PlateQuote
        Dim pj As New PlateJob

        'Set object properties
        With pq
            .PartDimST = CDec(txtPlatePartDimST.Text)
            .PartDimLT = CDec(txtPlatePartDimLT.Text)
            .PartDimL = CDec(txtPlatePartDimL.Text)
            .ExtraDimST = CDec(txtPlateExtraDimST.Text)
            .ExtraDimLT = CDec(txtPlateExtraDimLT.Text)
            .ExtraDimL = CDec(txtPlateExtraDimL.Text)
            .StdMatAdd = CDec(txtStdAdd.Text)


            pj = PlateCalc(pq) 'Get results
            'Display results
            txtOrderDimST.Text = Format(pj.DimST, "0.000")
            txtOrderDimLT.Text = Format(pj.DimLT, "0.000")
            txtOrderDimL.Text = Format(pj.DimL, "0.000")

            txtPlatePartDimST.Text = Format(.PartDimST, "0.000")
            txtPlatePartDimLT.Text = Format(.PartDimLT, "0.000")
            txtPlatePartDimL.Text = Format(.PartDimL, "0.000")
            txtPlateExtraDimST.Text = Format(.ExtraDimST, "0.000")
            txtPlateExtraDimLT.Text = Format(.ExtraDimLT, "0.000")
            txtPlateExtraDimL.Text = Format(.ExtraDimL, "0.000")
            txtStdAdd.Text = Format(.StdMatAdd, "0.000")

            rtxtPlateDetails.Text = "PL DIMs- ST: " & Format(.PartDimST, "0.000") & "  LT: " & Format(.PartDimLT, "0.000") & _
                                                                                "  L: " & Format(.PartDimL, "0.000") & vbLf & _
                                    "ORDER DIMs: " & Format(pj.DimST, "0.000") & " X " & Format(pj.DimLT, "0.000") & _
                                                    " X " & Format(pj.DimL, "0.000")
            My.Computer.Clipboard.SetText(rtxtPlateDetails.Text, TextDataFormat.Text)
        End With

    End Sub

    Private Sub ResetSheetMetal()
        'Clear fields
        txtSheetMetalQty.Text = ""
        txtSheetMetalDimA.Text = ""
        txtSheetMetalDimB.Text = ""

        txtShearSizeA.Text = ""
        txtShearSizeB.Text = ""

        txtShearSheetCountA.Text = ""
        txtShearSheetCountB.Text = ""

        txtFullSheetCountA.Text = ""
        txtFullSheetCountB.Text = ""

        txtShearSheetPctA.Text = ""
        txtShearSheetPctB.Text = ""

        txtYieldA.Text = ""
        txtYieldB.Text = ""
        chkRemnant.Checked = False
        rtxtSheetMetalDetails.Text = ""

        rdoOption1.Visible = False
        rdoOption2.Visible = False

        rdoOption1.Checked = False
        rdoOption2.Checked = False

        'Set focus on Qty
        txtSheetMetalQty.Focus()
    End Sub
    Private Sub ResetBar()
        'Clear fields
        txtBarQty.Text = ""
        txtNetPartOAL.Text = ""

        txtRequiredRndQty.Text = ""
        txtRandomYield.Text = ""
        txtMachineBarYield.Text = ""
        txtRequiredRndQty.Text = ""
        txtMachineBarQty.Text = ""

        rtxtbardetails.text = ""
        txtGrossPartYield.Text = ""

        'Set focus
        txtBarQty.Focus()
    End Sub
    Private Sub ResetPlate()
        'Clear Fields
        txtPlatePartDimST.Text = "0.000"
        txtPlatePartDimLT.Text = "0.000"
        txtPlatePartDimL.Text = "0.000"

        txtPlateExtraDimST.Text = "0.000"
        txtPlateExtraDimLT.Text = "0.000"
        txtPlateExtraDimL.Text = "0.000"

        txtOrderDimST.Text = "0.000"
        txtOrderDimLT.Text = "0.000"
        txtOrderDimL.Text = "0.000"

        rtxtPlateDetails.Text = ""
        txtStdAdd.Text = "0.250"

    End Sub

    Private Function isSheetMetalValidData_ForCalc() As Boolean
        If chkRemnant.Checked Then 'Dont evaluate qty

            If isNumberGreaterZero(txtSheetMetalDimA) AndAlso _
               isNumberGreaterZero(txtSheetMetalDimB) AndAlso _
               isNumberGreaterZero(txtWorkHolding) AndAlso _
               isNumberGreaterZero(txtStockDimA) AndAlso _
               isNumberGreaterZero(txtStockDimB) AndAlso _
               isNumberGreaterZero(txtFrame) AndAlso _
               isNumberGreaterZero(txtShearDimA) Then
                Return True
            Else : Return False
            End If

        Else

            If isNumberGreaterZero(txtSheetMetalQty) AndAlso _
                isNumberGreaterZero(txtSheetMetalDimA) AndAlso _
                isNumberGreaterZero(txtSheetMetalDimB) AndAlso _
                isNumberGreaterZero(txtWorkHolding) AndAlso _
                isNumberGreaterZero(txtStockDimA) AndAlso _
                isNumberGreaterZero(txtStockDimB) AndAlso _
                isNumberGreaterZero(txtFrame) AndAlso _
                isNumberGreaterZero(txtShearDimA) Then
                Return True
            Else : Return False
            End If
        End If


    End Function
    Private Function isSheetMetalValidData_ForConvert() As Boolean
        If isNumberGreaterZero(txtWorkHolding) AndAlso _
         isNumberGreaterZero(txtStockDimA) AndAlso _
         isNumberGreaterZero(txtStockDimB) AndAlso _
         isNumberGreaterZero(txtFrame) AndAlso _
         isNumberGreaterZero(txtShearDimA) Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Function isBarValidData_ForCalc() As Boolean

        If isNumberGreaterZero(txtBarQty) AndAlso _
            isNumberGreaterZero(txtNetPartOAL) AndAlso _
            isNumberGreaterZero(txtPartOffLength) AndAlso _
            isNumberGreaterEqualZero(txtExtraOpLength) AndAlso _
            isNumberGreaterEqualZero(txtCleanUpLength) AndAlso _
            isNumberGreaterZero(txtRandomLength) AndAlso _
            isNumberGreaterZero(txtBarFeedLength) AndAlso _
            isNumberGreaterZero(txtBarEndLength) Then
            Return True

        Else
            Return False
        End If
    End Function
    Private Function isBarValidData_ForConvert() As Boolean

        If isNumberGreaterZero(txtPartOffLength) AndAlso _
            isNumberGreaterEqualZero(txtExtraOpLength) AndAlso _
            isNumberGreaterEqualZero(txtCleanUpLength) AndAlso _
            isNumberGreaterZero(txtRandomLength) AndAlso _
            isNumberGreaterZero(txtBarFeedLength) AndAlso _
            isNumberGreaterZero(txtBarEndLength) Then
            Return True

        Else
            Return False
        End If
    End Function
    Private Function isPlateValidData_ForCalc() As Boolean

        'If isNumberGreaterZero(txtPlateQty) AndAlso _
        If isNumberGreaterZero(txtPlatePartDimST) AndAlso _
        isNumberGreaterZero(txtPlatePartDimLT) AndAlso _
        isNumberGreaterZero(txtPlatePartDimL) AndAlso _
        isNumberGreaterZero(txtStdAdd) AndAlso _
        isNumberGreaterEqualZero(txtPlateExtraDimST) AndAlso _
        isNumberGreaterEqualZero(txtPlateExtraDimLT) AndAlso _
        isNumberGreaterEqualZero(txtPlateExtraDimL) Then
            Return True
        Else : Return False
        End If

    End Function
    Private Function isNumberGreaterZero(ByVal txtBox As TextBox) As Boolean

        If IsNumeric(txtBox.Text) AndAlso CDec(txtBox.Text) > 0 Then
            Return True
        Else
            MessageBox.Show("Please enter a number greater than zero!", txtBox.Name, _
                MessageBoxButtons.OK, MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1)
            txtBox.SelectAll()
            txtBox.Focus()
            Return False
        End If
    End Function
    Private Function isNumberGreaterEqualZero(ByVal txtBox As TextBox) As Boolean

        If IsNumeric(txtBox.Text) AndAlso CDec(txtBox.Text) >= 0 Then
            Return True
        Else
            MessageBox.Show("Please enter a number greater than zero!", txtBox.Name, _
                MessageBoxButtons.OK, MessageBoxIcon.Hand, MessageBoxDefaultButton.Button1)
            txtBox.SelectAll()
            txtBox.Focus()
            Return False
        End If
    End Function

    Private Sub txtDimA_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSheetMetalDimA.GotFocus
        txtSheetMetalDimA.SelectAll()
    End Sub
    Private Sub txtDimB_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSheetMetalDimB.GotFocus
        txtSheetMetalDimB.SelectAll()
    End Sub
    Private Sub txtQty_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSheetMetalQty.GotFocus
        txtSheetMetalQty.SelectAll()
    End Sub
    Private Sub txtWorkHolding_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtWorkHolding.GotFocus
        txtWorkHolding.SelectAll()
    End Sub
    Private Sub txtStockSize_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtStockDimA.GotFocus
        txtStockDimA.SelectAll()
    End Sub
    Private Sub txtMaxShearSheet_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtShearDimA.GotFocus
        txtShearDimA.SelectAll()
    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Dim page As TabPage = tabCtrl.SelectedTab

        Select Case page.Name
            Case "pgSheetMetal"
                ResetSheetMetal()
            Case "pgBar"
                ResetBar()
            Case "pgPlate"
                ResetPlate()
        End Select
    End Sub
    Private Sub btnDefault_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDefault.Click
        Dim page As TabPage = tabCtrl.SelectedTab

        Select Case page.Name
            Case "pgSheetMetal"
                SheetMetalDefaults()
                txtSheetMetalQty.Focus() 'Set focus on Qty

            Case "pgBar"
                BarDefaults()
                txtBarQty.Focus() 'Set focus

            Case "PgPlate"
                txtPlatePartDimST.Focus()
                ResetPlate()

        End Select
    End Sub
    Private Sub btnCalc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalc.Click
        Dim page As TabPage = tabCtrl.SelectedTab

        Select Case page.Name
            Case "pgSheetMetal"
                If isSheetMetalValidData_ForCalc() Then
                    rtxtSheetMetalDetails.Text = ""
                    Process_SheetMetal()
                End If

            Case "pgBar"
                If isBarValidData_ForCalc() Then Process_Bar()
            Case "pgPlate"
                If isPlateValidData_ForCalc() Then Process_Plate()
        End Select

    End Sub

    Private Sub tabCtrl_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabCtrl.SelectedIndexChanged
        Dim page As TabPage = tabCtrl.SelectedTab

        Select Case page.Name
            Case "pgSheetMetal"
                'Set focus on Qty
                txtSheetMetalQty.Focus()

            Case "pgBar"
                txtBarQty.Focus() 'Set focus
                'txtBarQty.Focus()
            Case "pgPlate"
                txtPlatePartDimST.SelectAll()
                txtPlatePartDimST.Focus()
        End Select
    End Sub

    Private Sub frmMain_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim myBuildinfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(Application.ExecutablePath)
        'Set version in Title Bar
        Me.Text = "Material Calculator Ver - " & myBuildinfo.FileVersion

        'Set focus on Qty
        txtSheetMetalQty.Focus()
    End Sub

    Private Sub rdoEnglish_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoEnglish.CheckedChanged
        Dim page As TabPage = tabCtrl.SelectedTab

        Select Case page.Name
            Case "pgSheetMetal"
                'Set focus on Qty
                txtSheetMetalQty.Focus()

            Case "pgBar"
                'Set focus
                txtBarQty.Focus()
                UnitConversion_Bar()
            Case "PgPlate"


        End Select
    End Sub
    Private Sub UnitConversion_Bar()
        If isBarValidData_ForConvert() Then
            If rdoMetric.Checked Then
                txtPartOffLength.Text = Format(EnglishToMetric(CDec(txtPartOffLength.Text)), "0.00000")
                txtExtraOpLength.Text = Format(EnglishToMetric(CDec(txtExtraOpLength.Text)), "0.00000")
                txtBarFeedLength.Text = Format(EnglishToMetric(CDec(txtBarFeedLength.Text)), "0.00000")
                txtCleanUpLength.Text = Format(EnglishToMetric(CDec(txtCleanUpLength.Text)), "0.00000")
                txtBarEndLength.Text = Format(EnglishToMetric(CDec(txtBarEndLength.Text)), "0.00000")
                txtRandomLength.Text = Format(EnglishToMetric(CDec(txtRandomLength.Text)), "0.00000")

            Else
                txtPartOffLength.Text = Format(MetricToEnglish(CDec(txtPartOffLength.Text)), "0.000")
                txtExtraOpLength.Text = Format(MetricToEnglish(CDec(txtExtraOpLength.Text)), "0.000")
                txtBarFeedLength.Text = Format(MetricToEnglish(CDec(txtBarFeedLength.Text)), "0.000")
                txtCleanUpLength.Text = Format(MetricToEnglish(CDec(txtCleanUpLength.Text)), "0.000")
                txtBarEndLength.Text = Format(MetricToEnglish(CDec(txtBarEndLength.Text)), "0.000")
                txtRandomLength.Text = Format(MetricToEnglish(CDec(txtRandomLength.Text)), "0.000")
            End If
        End If
    End Sub
    Private Function MetricToEnglish(ByVal x As Decimal) As Decimal
        '1 millimeter = 0.0393700787 inches
        'Return x * 0.0393700787
    End Function
    Private Function EnglishToMetric(ByVal x As Decimal) As Decimal
        'Return x * 0.254
    End Function

    Private Sub SheetMetalDefaults()
        ResetSheetMetal()

        txtWorkHolding.Text = "4.000"

        txtStockDimA.Text = "48.000"
        txtStockDimB.Text = "144.000"

        txtShearDimA.Text = "36.000"
        txtShearDimB.Text = "48.000"
        txtFrame.Text = "0.500"
    End Sub

    Private Sub BarDefaults()

        rdoEnglish.Checked = True

        txtPartOffLength.Text = "0.125"
        txtExtraOpLength.Text = "0.020"
        txtCleanUpLength.Text = "1.030"
        txtBarEndLength.Text = "1.000"
        txtBarFeedLength.Text = "48"
        txtRandomLength.Text = "144"

    End Sub
    Private Sub EnableSheetMetalRadios()
        rdoOption1.Visible = True
        rdoOption2.Visible = True
    End Sub

    Private Sub DisbleSheetMetalRadios()
        rdoOption1.Visible = False
        rdoOption2.Visible = False
    End Sub

    Private Sub txtPlatePartDimST_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPlatePartDimST.GotFocus
        txtPlatePartDimST.SelectAll()
    End Sub

    Private Sub txtPlatePartDimLT_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPlatePartDimLT.GotFocus
        txtPlatePartDimLT.SelectAll()
    End Sub

    Private Sub txtPlatePartDimL_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPlatePartDimL.GotFocus
        txtPlatePartDimL.SelectAll()
    End Sub

    Private Sub txtPlateExtraDimST_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPlateExtraDimST.GotFocus
        txtPlateExtraDimST.SelectAll()
    End Sub

    Private Sub txtPlateExtraDimLT_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPlateExtraDimLT.GotFocus
        txtPlateExtraDimLT.SelectAll()
    End Sub

    Private Sub txtPlateExtraDimL_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPlateExtraDimL.GotFocus
        txtPlateExtraDimL.SelectAll()
    End Sub

    Private Sub chkRemnant_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkRemnant.CheckedChanged
        'ResetSheetMetal()
        If chkRemnant.Checked Then
            lblSheetMetalQty.Visible = False
            txtSheetMetalQty.Visible = False
            txtShearDimA.ReadOnly = True
        Else
            lblSheetMetalQty.Visible = True
            txtSheetMetalQty.Visible = True
            txtShearDimA.ReadOnly = False
        End If

    End Sub

    'Private Sub DisplayDetails(ByVal j As SheetMetalJob, ByVal q As SheetMetalQuote)

    '    rtxtSheetMetalDetails.Text = "PL DIM: " & Format(CDec(txtSheetMetalDimA.Text), "0.000") & " X " & Format(CDec(txtSheetMetalDimB.Text), "0.000") & "  YIELD: " & Format(j.Yield, "#,000") & vbCrLf & _
    '                                "SHEAR DIM: " & Format(CDec(j.ShearLength), "0.000") & " X " & Format(CDec(q.StockDimA), "0.000") & _
    '                                "  - QTY: " & CStr(j.ShearSheetQty) & vbLf & _
    '                                "ORDER DIM: " & Format(q.StockDimA) & " X " & Format(q.StockDimB) & "  - QTY: " & CStr(j.FullSheetQty)

    '    'Paste to clipboard
    '    My.Computer.Clipboard.SetText(rtxtSheetMetalDetails.Text, TextDataFormat.Text)
    'End Sub

    Private Sub rdoOption1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoOption1.CheckedChanged
        Dim smP As New SheetMetalPackage
        Dim smJobA As New SheetMetalJob, smQuoteA As New SheetMetalQuote

        If rdoOption1.Checked Then

            If isSheetMetalValidData_ForCalc() Then
                smP = Process_SheetMetal()
                With smP
                    rtxtSheetMetalDetails.Text = "PL DIM: " & Format(CDec(txtSheetMetalDimA.Text), "0.00") & " X " & Format(CDec(txtSheetMetalDimB.Text), "0.000") & "  YIELD: " & _
                                                Format(.smJobA.Yield, "#,000") & vbCrLf & "SHEAR DIM: " & Format(CDec(.smJobA.ShearLength), "0.000") & _
                                                " X " & Format(CDec(.smQuoteA.StockDimA), "0.000") & "  - QTY: " & CStr(.smJobA.ShearSheetQty) & vbLf & _
                                                "STOCK ORDER DIM: " & Format(.smQuoteA.StockDimA) & " X " & Format(.smQuoteA.StockDimB) & "  - QTY: " & CStr(.smJobA.FullSheetQty)

                    My.Computer.Clipboard.SetText(rtxtSheetMetalDetails.Text, TextDataFormat.Text)
                End With
            End If


        End If
    End Sub
    Private Sub rdoOption2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoOption2.CheckedChanged
        Dim smP As New SheetMetalPackage
        Dim smJobB As New SheetMetalJob, smQuoteB As New SheetMetalQuote

        If rdoOption2.Checked Then

            If isSheetMetalValidData_ForCalc() Then
                smP = Process_SheetMetal()
                With smP
                    rtxtSheetMetalDetails.Text = "PL DIM: " & Format(CDec(txtSheetMetalDimA.Text), "0.00") & " X " & Format(CDec(txtSheetMetalDimB.Text), "0.000") & "  YIELD: " & _
                                                Format(.smJobB.Yield, "#,000") & vbCrLf & "SHEAR DIM: " & Format(CDec(.smJobB.ShearLength), "0.000") & _
                                                " X " & Format(CDec(.smQuoteB.StockDimA), "0.000") & "  - QTY: " & CStr(.smJobB.ShearSheetQty) & vbLf & _
                                                "STOCK ORDER DIM: " & Format(.smQuoteB.StockDimA) & " X " & Format(.smQuoteB.StockDimB) & "  - QTY: " & CStr(.smJobB.FullSheetQty)

                    My.Computer.Clipboard.SetText(rtxtSheetMetalDetails.Text, TextDataFormat.Text)
                End With
            End If
        End If
    End Sub

  
End Class