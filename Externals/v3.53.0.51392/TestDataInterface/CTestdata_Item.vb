Option Explicit On
Option Strict On
Option Compare Text
Option Infer On


Namespace TestDataInterface
    Public Class CTestData_Item
        ' Itemlist and Result data
        Protected mItemNumber As Integer
        Protected mItemName As String
        Protected mDescriptor As String
        Protected mDescription As String
        Protected mReportLevel As Integer
        Protected mUnits As String
        Protected mCriticalSpec As Integer
        Protected mWarningMin As Double
        Protected mWarningMinSpec As Boolean
        Protected mWarningMax As Double
        Protected mWarningMaxSpec As Boolean
        Protected mFailMin As Double
        Protected mFailMinSpec As Boolean
        Protected mFailMax As Double
        Protected mFailMaxSpec As Boolean
        Protected mSanityMin As Double
        Protected mSanityMinSpec As Boolean
        Protected mSanityMax As Double
        Protected mSanityMaxSpec As Boolean
        Protected mItemBlobDataExists As Integer
        Protected mHasSpecs As Boolean
        Protected mIsGroup As Boolean


        ' Properties (Read Only)
        Public ReadOnly Property ItemNumber As Integer
            Get
                Return mItemNumber
            End Get
        End Property

        Public ReadOnly Property ItemName As String
            Get
                Return mItemName
            End Get
        End Property

        Public ReadOnly Property Descriptor As String
            Get
                Return mDescriptor
            End Get
        End Property

        Public ReadOnly Property Description As String
            Get
                Return mDescription
            End Get
        End Property

        Public ReadOnly Property ReportLevel As Integer
            Get
                Return mReportLevel
            End Get
        End Property

        Public ReadOnly Property Units As String
            Get
                Return mUnits
            End Get
        End Property

        Public ReadOnly Property CriticalSpec As Integer
            Get
                Return mCriticalSpec
            End Get
        End Property

        Public ReadOnly Property WarningMax As Double
            Get
                Return mWarningMax
            End Get
        End Property

        Public ReadOnly Property WarningMaxSpec As Boolean
            Get
                Return mWarningMaxSpec
            End Get
        End Property

        Public ReadOnly Property WarningMin As Double
            Get
                Return mWarningMin
            End Get
        End Property

        Public ReadOnly Property WarningMinSpec As Boolean
            Get
                Return mWarningMinSpec
            End Get
        End Property


        Public ReadOnly Property FailMax As Double
            Get
                Return mFailMax
            End Get
        End Property

        Public ReadOnly Property FailMaxSpec As Boolean
            Get
                Return mFailMaxSpec
            End Get
        End Property

        Public ReadOnly Property FailMin As Double
            Get
                Return mFailMin
            End Get
        End Property

        Public ReadOnly Property FailMinSpec As Boolean
            Get
                Return mFailMinSpec
            End Get
        End Property


        Public ReadOnly Property SanityMax As Double
            Get
                Return mSanityMax
            End Get
        End Property

        Public ReadOnly Property SanityMaxSpec As Boolean
            Get
                Return mSanityMaxSpec
            End Get
        End Property

        Public ReadOnly Property SanityMin As Double
            Get
                Return mSanityMin
            End Get
        End Property

        Public ReadOnly Property SanityMinSpec As Boolean
            Get
                Return mSanityMinSpec
            End Get
        End Property


        Public ReadOnly Property ItemBlobDataExists As Integer
            Get
                Return mItemBlobDataExists
            End Get
        End Property

        Public ReadOnly Property HasSpecs As Boolean
            Get
                Return mHasSpecs
            End Get
        End Property


        ' Group Identification
        Public Property IsGroup As Boolean
            Get
                Return mIsGroup
            End Get
            Set
                mIsGroup = value
            End Set
        End Property

        '**********************************************************************
        '* Methods
        '**********************************************************************

        Friend Sub New(ItemNumber As Integer,
                       ItemName As String,
                       Descriptor As String,
                       Description As String,
                       ReportLevel As Integer,
                       Units As String,
                       CriticalSpec As Integer,
                       WarningMin As Double,
                       WarningMax As Double,
                       FailMin As Double,
                       FailMax As Double,
                       SanityMin As Double,
                       SanityMax As Double) _
' Function populates the item object,
            ' Specs are passed in as strings to allow for null (no specs)

            mItemNumber = ItemNumber
            mItemName = ItemName
            mDescriptor = Descriptor
            mDescription = Description
            mReportLevel = ReportLevel
            mUnits = Units
            mCriticalSpec = CriticalSpec

            mWarningMin = WarningMin
            mWarningMax = WarningMax
            mFailMin = FailMin
            mFailMax = FailMax
            mSanityMin = SanityMin
            mSanityMax = SanityMax

            mHasSpecs = False

            If Double.IsNaN(mWarningMin) Then
                mWarningMinSpec = False
            Else
                mWarningMinSpec = True
                mHasSpecs = True
            End If

            If Double.IsNaN(mWarningMax) Then
                mWarningMaxSpec = False
            Else
                mWarningMaxSpec = True
                mHasSpecs = True
            End If

            If Double.IsNaN(mFailMin) Then
                mFailMinSpec = False
            Else
                mFailMinSpec = True
                mHasSpecs = True
            End If

            If Double.IsNaN(mFailMax) Then
                mFailMaxSpec = False
            Else
                mFailMaxSpec = True
                mHasSpecs = True
            End If

            If Double.IsNaN(mSanityMin) Then
                mSanityMinSpec = False
            Else
                mSanityMinSpec = True
                mHasSpecs = True
            End If

            If Double.IsNaN(mSanityMax) Then
                mSanityMaxSpec = False
            Else
                mSanityMaxSpec = True
                mHasSpecs = True
            End If
        End Sub
    End Class
End Namespace
