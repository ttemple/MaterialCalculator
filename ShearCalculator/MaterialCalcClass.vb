
Public Class SheetMetalPackage
    Protected m_SMQuoteA As SheetMetalQuote
    Protected m_SMQuoteB As SheetMetalQuote
    Protected m_SMJobA As SheetMetalJob
    Protected m_smJobB As SheetMetalJob

    Public Property smQuoteA() As SheetMetalQuote
        Get
            Return m_SMQuoteA
        End Get
        Set(ByVal value As SheetMetalQuote)
            m_SMQuoteA = value
        End Set
    End Property

    Public Property smQuoteB() As SheetMetalQuote
        Get
            Return m_SMQuoteB
        End Get
        Set(ByVal value As SheetMetalQuote)
            m_SMQuoteB = value
        End Set
    End Property

    Public Property smJobA As SheetMetalJob
        Get
            Return m_SMJobA
        End Get
        Set(ByVal value As SheetMetalJob)
            m_SMJobA = value
        End Set
    End Property
    Public Property smJobB() As SheetMetalJob
        Get
            Return m_smJobB
        End Get
        Set(ByVal value As SheetMetalJob)
            m_smJobB = value
        End Set
    End Property
End Class
Public Class SheetMetalQuote
    Protected m_Qty As Integer
    Protected m_DimA As Decimal
    Protected m_DimB As Decimal
    Protected m_Workholding As Decimal
    Protected m_StockSizeDimA As Decimal
    Protected m_StockSizeDimB As Decimal
    Protected m_MaxShearSize As Decimal
    Protected m_Frame As Decimal

    Public Property FrameDimension() As Decimal
        Get
            Return m_Frame
        End Get
        Set(ByVal value As Decimal)
            m_Frame = value
        End Set
    End Property


    Public Property ShearSizeDimA()
        Get
            Return m_MaxShearSize
        End Get
        Set(ByVal value)
            m_MaxShearSize = value
        End Set
    End Property

    Public Property TotalQty()
        Get
            Return m_Qty
        End Get
        Set(ByVal value)
            m_Qty = value
        End Set
    End Property
    Public Property dimA()
        Get
            Return m_DimA
        End Get
        Set(ByVal value)
            m_DimA = value
        End Set
    End Property
    Public Property dimB()
        Get
            Return m_DimB
        End Get
        Set(ByVal value)
            m_DimB = value
        End Set
    End Property
 
    Public Property Workholding()
        Get
            Return m_Workholding
        End Get
        Set(ByVal value)
            m_Workholding = value
        End Set
    End Property
    Public Property StockDimA()
        Get
            Return m_StockSizeDimA
        End Get
        Set(ByVal value)
            m_StockSizeDimA = value
        End Set
    End Property
    Public Property StockDimB()
        Get
            Return m_StockSizeDimB
        End Get
        Set(ByVal value)
            m_StockSizeDimB = value
        End Set
    End Property
End Class
Public Class SheetMetalJob
    Protected m_ShearLength As Decimal
    Protected m_ShearSheetQty As Integer
    Protected m_FullSheetQty As Integer
    Protected m_Yield As Integer
    Protected m_UnitPercentageShearSheet As Decimal

    Public Property Yield()
        Get
            Return m_Yield
        End Get
        Set(ByVal value)
            m_Yield = value
        End Set
    End Property
    Public Property FullSheetQty() As Integer
        Get
            Return m_FullSheetQty
        End Get
        Set(ByVal value As Integer)
            m_FullSheetQty = value
        End Set
    End Property
    Public Property ShearSheetQty() As Integer
        Get
            Return m_ShearSheetQty
        End Get
        Set(ByVal value As Integer)
            m_ShearSheetQty = value
        End Set
    End Property
    Public Property ShearLength() As Decimal
        Get
            Return m_ShearLength
        End Get
        Set(ByVal value As Decimal)
            m_ShearLength = value
        End Set
    End Property
    Public Property UnitPercentageShearSheet As Decimal
        Get
            Return m_UnitPercentageShearSheet
        End Get

        Set(ByVal value As Decimal)
            m_UnitPercentageShearSheet = value
        End Set

    End Property
End Class

Public Class BarQuote
    Protected m_Qty As Integer
    Protected m_PartLength As Decimal
    Protected m_DimB As Decimal
    Protected m_CleanUpLength As Decimal
    Protected m_PartOffLength As Decimal
    Protected m_XtraOpLength As Decimal
    Protected m_RandomLength As Decimal
    Protected m_BarFeedLength As Decimal
    Protected m_BarEndLength As Decimal

    Public Property BarEndLength() As Decimal
        Get
            Return m_BarEndLength
        End Get
        Set(ByVal value As Decimal)
            m_BarEndLength = value
        End Set
    End Property

    Public Property Qty() As Integer
        Get
            Return m_Qty
        End Get
        Set(ByVal value As Integer)
            m_Qty = value
        End Set
    End Property

    Public Property PartOAL() As Decimal
        Get
            Return m_PartLength
        End Get
        Set(ByVal value As Decimal)
            m_PartLength = value
        End Set
    End Property

    Public Property CleanUpLength() As Decimal
        Get
            Return m_CleanUpLength
        End Get
        Set(ByVal value As Decimal)
            m_CleanUpLength = value
        End Set
    End Property
    Public Property PartOffLength() As Decimal
        Get
            Return m_PartOffLength
        End Get
        Set(ByVal value As Decimal)
            m_PartOffLength = value
        End Set
    End Property
    Public Property ExtraOpLength() As Decimal
        Get
            Return m_XtraOpLength
        End Get
        Set(ByVal value As Decimal)
            m_XtraOpLength = value
        End Set
    End Property
    Public Property RandomLength() As Decimal
        Get
            Return m_RandomLength
        End Get
        Set(ByVal value As Decimal)
            m_RandomLength = value
        End Set
    End Property
    Public Property BarFeedLength() As Decimal
        Get
            Return m_BarFeedLength
        End Get
        Set(ByVal value As Decimal)
            m_BarFeedLength = value
        End Set
    End Property

End Class
Public Class BarJob
    Protected m_RandomYield As Integer
    Protected m_PartYieldPerMachineBar As Integer
    Protected m_MachineBarsPerRandom As Integer
    Protected m_MachineBarQty
    Protected m_RandomQty
    Protected m_PartOAL
    Protected m_GrossPartYield As Integer

    Public Property GrossPartYield() As Integer
        Get
            Return m_GrossPartYield
        End Get
        Set(ByVal value As Integer)
            m_GrossPartYield = value
        End Set
    End Property
    Public Property PartOAL() As Decimal
        Get
            Return m_PartOAL
        End Get
        Set(ByVal value As Decimal)
            m_PartOAL = value
        End Set
    End Property

    Public Property PartYieldPerRandom() As Integer
        Get
            Return m_RandomYield
        End Get
        Set(ByVal value As Integer)
            m_RandomYield = value
        End Set
    End Property
    Public Property MachineBarsPerRandom() As Integer
        Get
            Return m_MachineBarsPerRandom
        End Get
        Set(ByVal value As Integer)
            m_MachineBarsPerRandom = value
        End Set
    End Property
    Public Property PartYieldPerMachineBar() As Integer
        Get
            Return m_PartYieldPerMachineBar
        End Get
        Set(ByVal value As Integer)
            m_PartYieldPerMachineBar = value
        End Set
    End Property
    Public Property QtyMachineBarsReq() As Integer
        Get
            Return m_MachineBarQty
        End Get
        Set(ByVal value As Integer)
            m_MachineBarQty = value
        End Set
    End Property
    Public Property QtyRandomsRequired() As Integer
        Get
            Return m_RandomQty
        End Get
        Set(ByVal value As Integer)
            m_RandomQty = value
        End Set
    End Property
End Class
Public Class PlateJob
    Protected m_DimST As Decimal
    Protected m_DimLT As Decimal
    Protected m_DimL As Decimal
    Public Property DimST() As Decimal
        Get
            Return m_DimST
        End Get
        Set(ByVal value As Decimal)
            m_DimST = value
        End Set
    End Property
    Public Property DimLT() As Decimal
        Get
            Return m_DimLT
        End Get
        Set(ByVal value As Decimal)
            m_DimLT = value
        End Set
    End Property
    Public Property DimL() As Decimal
        Get
            Return m_DimL
        End Get
        Set(ByVal value As Decimal)
            m_DimL = value
        End Set
    End Property
End Class
Public Class PlateQuote
    Protected m_Qty As Integer
    Protected m_StdAddition As Decimal
    Protected m_PartDimST As Decimal
    Protected m_PartDimLT As Decimal
    Protected m_PartDimL As Decimal
    Protected m_PartExtraDimST As Decimal
    Protected m_PartExtraDimLT As Decimal
    Protected m_PartExtraDimL As Decimal

    Public Property Qty() As Integer
        Get
            Return m_Qty
        End Get
        Set(ByVal value As Integer)
            m_Qty = value
        End Set
    End Property
    Public Property StdMatAdd() As Decimal
        Get
            Return m_StdAddition
        End Get
        Set(ByVal value As Decimal)
            m_StdAddition = value
        End Set
    End Property

    Public Property PartDimST() As Decimal
        Get
            Return m_PartDimST
        End Get
        Set(ByVal value As Decimal)
            m_PartDimST = value
        End Set
    End Property
    Public Property PartDimLT() As Decimal
        Get
            Return m_PartDimLT
        End Get
        Set(ByVal value As Decimal)
            m_PartDimLT = value
        End Set
    End Property
    Public Property PartDimL() As Decimal
        Get
            Return m_PartDimL
        End Get
        Set(ByVal value As Decimal)
            m_PartDimL = value
        End Set
    End Property
    Public Property ExtraDimST() As Decimal
        Get
            Return m_PartExtraDimST
        End Get
        Set(ByVal value As Decimal)
            m_PartExtraDimST = value
        End Set
    End Property
    Public Property ExtraDimLT() As Decimal
        Get
            Return m_PartExtraDimLT
        End Get
        Set(ByVal value As Decimal)
            m_PartExtraDimLT = value
        End Set
    End Property
    Public Property ExtraDimL() As Decimal
        Get
            Return m_PartExtraDimL
        End Get
        Set(ByVal value As Decimal)
            m_PartExtraDimL = value
        End Set
    End Property
End Class