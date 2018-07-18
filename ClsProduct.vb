'Copyright (c) 2009-2017 Dan Turk

#Region "Class / File Comment Header block"
'Program:                proj03-FillingStation-Fang-Anmin
'File:                   ClsProduct.vb
'Author:                 Anmin Fang
'Description:            This is a class to create a product object.
'Date:                   2017 Oct 5
'                          - Added attributes, get/set methods, constructors.
'                          - Added private/public createFuleTank() methods
'                        2017 Nov 7
'                          - Added translate Enum numbers into product type method
'Tier:                   Business Logic
'Exceptions:             None generated.
'Exception-Handling:     None.
'Events:                 None generated.
'Event-Handling:         None. 
#End Region 'Class / File Comment Header block

#Region "Option / Imports"
Option Explicit On      'Must declare variables before using them
Option Strict On        'Must perform explicit data type conversions
#End Region 'Option / Imports

Public Class Product

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    '********** Module-level constants

    '********** Module-level variables

    Private mId As String
    Private mFuelTank As FuelTank
    Private mProductName As String
    Private mProductType As ProductType
    Private mUnitOfMeasure As String
    Private mPricePerUnit As Decimal
    Private mLoyaltyDicountPerUnit As Decimal
    Private mRewardDiscountPerUnit As Decimal
    Private mTaxRate As Decimal

#End Region 'Attributes

#Region "Constructors"
    '******************************************************************
    'Constructors
    '******************************************************************

    'These are all public.

    '********** Default constructor
    '             - no parameters

    '********** Special constructor(s)
    '             - typically constructors have parameters 
    '               that are used to initialize attributes

    Public Sub New()

    End Sub

    Public Sub New(
           ByVal pId As String,
           ByVal pProductName As String,
           ByVal pProductType As ProductType,
           ByVal pUnitOfMeasure As String,
           ByVal pPricePerUnit As Decimal,
           ByVal pLoyaltyDiscountPerUnit As Decimal,
           ByVal pRewardDiscountPerUnit As Decimal,
           ByVal pTaxRate As Decimal
           )

        MyBase.New

        _id = pId
        _productName = pProductName
        _productType = pProductType
        _unitOfMeasure = pUnitOfMeasure
        _pricePerUnit = pPricePerUnit
        _loyaltyDiscountPerUnit = pLoyaltyDiscountPerUnit
        _rewardDiscountPerUnit = pRewardDiscountPerUnit
        _taxRate = pTaxRate
        '_fuelTank.id = pFuelTankId
        '_fuelTank.currentFuelTank = pCurrentFuelTank
        '_fuelTank.maxFuelTank = pMaxFuelTank

    End Sub 'New()

    '********** Copy constructor(s)
    '             - one parameter, an object of the same class

#End Region 'Constructors

#Region "Get/Set Methods"
    '******************************************************************
    'Get/Set Methods
    '******************************************************************

    '********** Public Get/Set Methods
    '             - call private get/set methods to implement

    Public ReadOnly Property id As String
        Get
            Return _id
        End Get
    End Property

    Public Property fuelTank As FuelTank
        Get
            Return _fuelTank
        End Get
        Set(ByVal pValue As FuelTank)
            _fuelTank = pValue
        End Set
    End Property

    Public ReadOnly Property productName As String
        Get
            Return _productName
        End Get
        'Set(ByVal pValue As String)
        '    _productName = pValue
        'End Set
    End Property

    Public ReadOnly Property productType As ProductType
        Get
            Return _productType
        End Get
        'Set(ByVal pValue As ProductType)
        '    _productType = pValue
        'End Set
    End Property

    Public ReadOnly Property unitOfMeasure As String
        Get
            Return _unitOfMeasure
        End Get
        'Set(ByVal pValue As String)
        '    _unitOfMeasure = pValue
        'End Set
    End Property

    Public Property pricePerUnit As Decimal
        Get
            Return _pricePerUnit
        End Get
        Set(ByVal pValue As Decimal)
            _pricePerUnit = pValue
        End Set
    End Property

    Public Property loyaltyDiscountPerUnit As Decimal
        Get
            Return _loyaltyDiscountPerUnit
        End Get
        Set(value As Decimal)

        End Set
    End Property

    Public Property rewardDiscountPerUnit As Decimal
        Get
            Return _rewardDiscountPerUnit
        End Get
        Set(ByVal pValue As Decimal)
            _rewardDiscountPerUnit = pValue
        End Set
    End Property

    Public Property taxRate As Decimal
        Get
            Return _taxRate
        End Get
        Set(ByVal pValue As Decimal)
            _taxRate = pValue
        End Set
    End Property

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)
    Private Property _id As String
        Get
            Return mId
        End Get
        Set(ByVal pValue As String)
            mId = pValue
        End Set
    End Property

    Private Property _fuelTank As FuelTank
        Get
            Return mFuelTank
        End Get
        Set(ByVal pValue As FuelTank)
            mFuelTank = pValue
        End Set
    End Property

    Private Property _productName As String
        Get
            Return mProductName
        End Get
        Set(ByVal pValue As String)
            mProductName = pValue
        End Set
    End Property

    Private Property _productType As ProductType
        Get
            Return mProductType
        End Get
        Set(ByVal pValue As ProductType)
            mProductType = pValue
        End Set
    End Property

    Private Property _unitOfMeasure As String
        Get
            Return mUnitOfMeasure
        End Get
        Set(ByVal pValue As String)
            mUnitOfMeasure = pValue
        End Set
    End Property

    Private Property _pricePerUnit As Decimal
        Get
            Return mPricePerUnit
        End Get
        Set(ByVal pValue As Decimal)
            mPricePerUnit = pValue
        End Set
    End Property

    Private Property _loyaltyDiscountPerUnit As Decimal
        Get
            Return mLoyaltyDicountPerUnit
        End Get
        Set(ByVal pValue As Decimal)
            mLoyaltyDicountPerUnit = pValue
        End Set
    End Property

    Private Property _rewardDiscountPerUnit As Decimal
        Get
            Return mRewardDiscountPerUnit
        End Get
        Set(ByVal pValue As Decimal)
            mRewardDiscountPerUnit = pValue
        End Set
    End Property

    Private Property _taxRate As Decimal
        Get
            Return mTaxRate
        End Get
        Set(ByVal pValue As Decimal)
            mTaxRate = pValue
        End Set
    End Property

#End Region 'Get/Set Methods

#Region "Behavioral Methods"
    '******************************************************************
    'Behavioral Methods
    '******************************************************************

    '********** Public Shared Behavioral Methods

    '********** Private Shared Behavioral Methods

    '********** Public Non-Shared Behavioral Methods

    Public Function createFuelTank(
           ByVal pId As String,
           ByVal pCurrentFuelQuantity As Decimal,
           ByVal pMaxFuelQuantity As Decimal
           ) _
       As _
           FuelTank

        Return _createFuelTank(
            pId,
            pCurrentFuelQuantity,
            pMaxFuelQuantity
            )

    End Function 'createFuelTank(pId,pCurrentFuelQuantity,pMaxFuelQuantity)

    Public Function ToString() As String

        Return _toString()

    End Function 'ToString()

    '********** Private Non-Shared Behavioral Methods

    Private Function _createFuelTank(
            ByVal pId As String,
            ByVal pCurrentFuelQuantity As Decimal,
            ByVal pMaxFuelQuantity As Decimal
            ) _
        As _
            FuelTank

        Dim theFuelTank As FuelTank

        theFuelTank = New FuelTank(
            pId,
            pCurrentFuelQuantity,
            pMaxFuelQuantity,
            Me
            )
        _fuelTank = theFuelTank

        Return theFuelTank

    End Function '_createFuelTank(pId,pCurrentFuelQuantity,pMaxFuelQuantity)

    Private Function _transferFromEnum(
            ByVal pProductType As ProductType
            ) _
        As _
            String

        Dim theProductType As String

        theProductType = CType(pProductType, ProductType).ToString

        Return theProductType

    End Function '_transferFromEnum(pProductType)

    Private Function _toString() As String

        Dim tempStr As String

        If _productType = ProductType.FUEL Then
            tempStr = "( PRODUCT:" _
                                   & " id=" & _id _
                                   & ", product name=" & _productName _
                                   & ", product type=" & _transferFromEnum(_productType) _
                                   & ", unit of measure=" & _unitOfMeasure _
                                   & ", price per unit=" & _pricePerUnit _
                                   & ", loyalty discount per unit=" & _loyaltyDiscountPerUnit _
                                   & ", reward discount per unit=" & _rewardDiscountPerUnit _
                                   & ", tax rate=" & _taxRate _
                                   & " ( FULETANK: " _
                                   & " id=" & _fuelTank.id _
                                   & ", current quantity=" & _fuelTank.currentFuelTank _
                                   & ", max quantity=" & _fuelTank.maxFuelTank _
                                   & " ))"
        Else
            tempStr = "( PRODUCT:" _
                                  & " id=" & _id _
                                  & ", product name=" & _productName _
                                  & ", product type=" & _transferFromEnum(_productType) _
                                  & ", unit of measure=" & _unitOfMeasure _
                                  & ", price per unit=" & _pricePerUnit _
                                  & ", loyalty discount per unit=" & _loyaltyDiscountPerUnit _
                                  & ", reward discount per unit=" & _rewardDiscountPerUnit _
                                  & ", tax rate=" & _taxRate _
                                  & " )"
        End If

        Return tempStr

    End Function '_toString()

#End Region 'Behavioral Methods

#Region "Event Procedures"
    '******************************************************************
    'Event Procedures
    '******************************************************************

    'No Event Procedures are currently defined.
    'These are all private.

    '********** User-Interface Event Procedures
    '             - Initiated explicitly by user

    '********** User-Interface Event Procedures
    '             - Initiated automatically by system

    '********** Business Logic Event Procedures
    '             - Initiated as a result of business logic
    '               method(s) running


#End Region 'Event Procedures

#Region "Events"
    '******************************************************************
    'Events
    '******************************************************************

    'These are all public.

#End Region 'Events

End Class 'Product
