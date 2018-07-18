'Copyright (c) 2009-2017 Dan Turk

#Region "Class / File Comment Header block"
'Program:                proj03-FillingStation-Fang-Anmin
'File:                   ClsSale.vb
'Author:                 Anmin Fang
'Description:            This is a class to create a sale object, which includes transaction information.
'Date:                   2017 Oct 5
'                          - Added attributes, get/set methods, constructors.
'                        2017 Nov 26
'                          - Revised toString method to adapt customer is nothing situation
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

Public Class Sale

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    '********** Module-level constants

    '********** Module-level variables

    Private mId As String
    Private mCustomer As LoyaltyCustomer
    Private mProduct As Product
    Private mSaleDate As DateTime
    Private mQuantityPurchased As Decimal
    Private mPricePerUnit As Decimal
    Private mDiscountPerUnit As Decimal
    Private mTaxRate As Decimal
    Private mSubTotalAmount As Decimal
    Private mTaxAmount As Decimal
    Private mTotalAmount As Decimal

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
    Public Sub New(
            ByVal pId As String,
            ByVal pCustomer As LoyaltyCustomer,
            ByVal pProduct As Product,
            ByVal pSaleDate As DateTime,
            ByVal pQuantityPruchased As Decimal,
            ByVal pPricePerUnit As Decimal,
            ByVal pDiscountPerUnit As Decimal,
            ByVal pTaxRate As Decimal,
            ByVal pSubTotalAmount As Decimal,
            ByVal pTaxAmount As Decimal,
            ByVal pTotalAmount As Decimal
            )

        MyBase.New

        _id = pId
        _customer = pCustomer
        _product = pProduct
        _SaleDate = pSaleDate
        _quantityPurchased = pQuantityPruchased
        _pricePerUnit = pPricePerUnit
        _discountPerUnit = pDiscountPerUnit
        _taxRate = pTaxRate
        _subTotalAmount = pSubTotalAmount
        _taxAmount = pTaxAmount
        _totalAmount = pTotalAmount

    End Sub 'New() with loyalty customer

    Public Sub New(
           ByVal pId As String,
           ByVal pProduct As Product,
           ByVal pSaleDate As DateTime,
           ByVal pQuantityPruchased As Decimal,
           ByVal pPricePerUnit As Decimal,
           ByVal pTaxRate As Decimal,
           ByVal pSubTotalAmount As Decimal,
           ByVal pTaxAmount As Decimal,
           ByVal pTotalAmount As Decimal
           )

        MyBase.New

        _id = pId
        _product = pProduct
        _SaleDate = pSaleDate
        _quantityPurchased = pQuantityPruchased
        _pricePerUnit = pPricePerUnit
        _taxRate = pTaxRate
        _subTotalAmount = pSubTotalAmount
        _taxAmount = pTaxAmount
        _totalAmount = pTotalAmount

    End Sub 'New() without loyalty customer

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
            Return _id()
        End Get
    End Property

    Public ReadOnly Property customer As LoyaltyCustomer
        Get
            Return _customer()
        End Get
    End Property

    Public ReadOnly Property product As Product
        Get
            Return _product
        End Get
    End Property

    Public Property saleDate As DateTime
        Get
            Return _SaleDate
        End Get
        Set(ByVal pValue As DateTime)
            _SaleDate = pValue
        End Set
    End Property

    Public Property quantityPurchased As Decimal
        Get
            Return _quantityPurchased
        End Get
        Set(ByVal pValue As Decimal)
            _quantityPurchased = pValue
        End Set
    End Property

    Public Property pricePerUnit As Decimal
        Get
            Return _pricePerUnit
        End Get
        Set(ByVal pValue As Decimal)
            _pricePerUnit = pValue
        End Set
    End Property

    Public Property discountPerUnit As Decimal
        Get
            Return _discountPerUnit
        End Get
        Set(ByVal pValue As Decimal)
            _discountPerUnit = pValue
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

    Public Property subTotalAmount As Decimal
        Get
            Return _subTotalAmount
        End Get
        Set(ByVal pValue As Decimal)
            _subTotalAmount = pValue
        End Set
    End Property

    Public Property taxAmount As Decimal
        Get
            Return _taxAmount
        End Get
        Set(ByVal pValue As Decimal)
            _taxAmount = pValue
        End Set
    End Property

    Public Property totalAmount As Decimal
        Get
            Return _totalAmount
        End Get
        Set(ByVal pValue As Decimal)
            _totalAmount = pValue
        End Set
    End Property

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _id() As String
        Get
            Return mId
        End Get
        Set(ByVal pValue As String)
            mId = pValue
        End Set
    End Property

    Private Property _customer() As LoyaltyCustomer
        Get
            Return mCustomer
        End Get
        Set(ByVal pValue As LoyaltyCustomer)
            mCustomer = pValue
        End Set
    End Property

    Private Property _product() As Product
        Get
            Return mProduct
        End Get
        Set(ByVal pValue As Product)
            mProduct = pValue
        End Set
    End Property

    Private Property _SaleDate() As DateTime
        Get
            Return mSaleDate
        End Get
        Set(ByVal pValue As DateTime)
            mSaleDate = pValue
        End Set
    End Property

    Private Property _quantityPurchased() As Decimal
        Get
            Return mQuantityPurchased
        End Get
        Set(ByVal pValue As Decimal)
            mQuantityPurchased = pValue
        End Set
    End Property

    Private Property _pricePerUnit() As Decimal
        Get
            Return mPricePerUnit
        End Get
        Set(ByVal pValue As Decimal)
            mPricePerUnit = pValue
        End Set
    End Property

    Private Property _discountPerUnit() As Decimal
        Get
            Return mDiscountPerUnit
        End Get
        Set(ByVal pValue As Decimal)
            mDiscountPerUnit = pValue
        End Set
    End Property

    Private Property _taxRate() As Decimal
        Get
            Return mTaxRate
        End Get
        Set(ByVal pValue As Decimal)
            mTaxRate = pValue
        End Set
    End Property

    Private Property _subTotalAmount() As Decimal
        Get
            Return mSubTotalAmount
        End Get
        Set(ByVal pValue As Decimal)
            mSubTotalAmount = pValue
        End Set
    End Property

    Private Property _taxAmount() As Decimal
        Get
            Return mTaxAmount
        End Get
        Set(ByVal pValue As Decimal)
            mTaxAmount = pValue
        End Set
    End Property

    Private Property _totalAmount() As Decimal
        Get
            Return mTotalAmount
        End Get
        Set(ByVal pValue As Decimal)
            mTotalAmount = pValue
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

    Public Function ToString() As String

        Return _toString()

    End Function 'ToString()

    '********** Private Non-Shared Behavioral Methods

    Private Function _toString() As String

        Dim tempStr As String

        If _customer Is Nothing Then
            tempStr = "( Sale: " _
                        & "id=" & _id _
                        & ", product=" & _product.ToString _
                        & ", sale date=" & _SaleDate _
                        & ", quantity purchased=" & _quantityPurchased _
                        & ", price per unit=" & _pricePerUnit _
                        & ", discount per unit=" & _discountPerUnit _
                        & ", tax rate=" & _taxRate _
                        & ", sub total amount=" & _subTotalAmount _
                        & ", tax amount=" & _taxAmount _
                        & ", total amount=" & _totalAmount _
                        & " )"
        Else
            tempStr = "( Sale: " _
                       & "id=" & _id _
                       & ", customer=" & _customer.ToString _
                       & ", product=" & _product.ToString _
                       & ", sale date=" & _SaleDate _
                       & ", quantity purchased=" & _quantityPurchased _
                       & ", price per unit=" & _pricePerUnit _
                       & ", discount per unit=" & _discountPerUnit _
                       & ", tax rate=" & _taxRate _
                       & ", sub total amount=" & _subTotalAmount _
                       & ", tax amount=" & _taxAmount _
                       & ", total amount=" & _totalAmount _
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

    'No Events are currently defined.
    'These are all public.

#End Region 'Events

End Class 'Sale
