'Copyright (c) 2009-2017 Dan Turk

#Region "Class / File Comment Header block"
'Program:                proj04-FillingStation-Fang-Anmin
'File:                   ClsFillingStation.vb
'Author:                 Anmin Fang
'Description:            This is a class to acquire information from other classes 
'                        and provide methods to userinterface when it needs.
'Date:                   2017 Oct 5
'                          - Added attributes, get/set methods, constructors, toString methods.
'                          - Added private/public createCustomer, createProduct,
'                            createFuleTank, createSale and createProduct methods
'                        2017 Nov 2
'                           - Added eventarts class into create method
'                        2017 Nov 10
'                           - Added modifyProduct method
'                           - Added array
'                        2017 Nov 26
'                           - Revise _createSale method to adapt customer is nothing
'                             situation
'                           - Added _findListboxCustomer & _findComboboxCustomer, 
'                             _findListboxProduct & _findComboboxProduct methods(private&public)
'                        2017 Nov 27
'                           - Added calculation of total sale and total tax amount methods(private&public)
'                           - Added metrics calculation
'                           - Added optional metrics(smallest amount of sale)
'                        2017 Nov 28
'                           - Revised smallest amount of sale method
'                           - Added biggest amount of sale method
'                        2017 Nov 29
'                           - Added _orderFuel method
'                           - Revised find methods
'                           - Added fule amount calculation in createSale method
'                        2017 Dec 1
'                           - Added transaction method into createCustomer, createProduct method
'                        2017 Dec 2
'                           - Added iterator methods for customer, product, fueltank, sale and transaction
'                             array(private&Public)
'                        2017 Dec 3
'                           - Added _readFromFile method
'                           - Revised _createCustomer,product,sale methods to adapt _readFromFile method
'                        2017 Dec 4
'                           - Revised modify function in _readFromFile method
'                           - Added writeToFile method
'                           - Revised createSale method and added validations
'                        2017 Dec 5
'                           - Added transaction date paramenter into createCustomer method
'                           - Revised orderFuel method that added trx ID parameter
'                           - Revised transaction text to adate text fiel that read transaction date from the
'                             and store into theTrxText
'                           - Revised transaction date in _createProduct, _createSale, _modifyProduct, _orderFuel
'                             methods
'                           - Revised sale.customer is existed or not problem in _calTotalAmountPerCustomer method
'                           - Revised member since age calculation method
'                        2017 Dec 8
'                           - Added get smallest and largest sale product type methods
'Tier:                   Business Logic
'Exceptions:             IndexOutOfRangeException, NullReferenceException
'Exception-Handling:     Throw,try/catch
'Events:                 customer, product, fueltank, sale, transaction added event
'                        product modified
'Event-Handling:         FillingStation_CustomerAdded, FillingStation_ProductAdded, FillingStation_FuelTankAdded
'                        FillingStation_SaleAdded, FillingStation_TransactionAdded, FillingStation_ProductModified
#End Region 'Class / File Comment Header block

#Region "Option / Imports"
Option Explicit On      'Must declare variables before using them
Option Strict On        'Must perform explicit data type conversions
Option Compare Text
Imports System.IO
Imports System.Globalization
#End Region 'Option / Imports

Public Class FillingStation

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    '********** Module-level constants

    Private Const mARRAY_SIZE__INCREATMENT_DEFAULT As Integer = 5

    '********** Module-level variables

    Private mFillingStationName As String
    Private mNumCustomers As Integer
    Private mNumProducts As Integer
    Private mNumSales As Integer
    Private mNumTransactions As Integer
    Private mNumFuelTanks As Integer
    Private mCustomers() As LoyaltyCustomer
    Private mProducts() As Product
    Private mSales() As Sale
    Private mTransactions() As Transaction
    Private mFuelTanks() As FuelTank

    Private mMaxCustomerSize As Integer
    Private mMaxProductSize As Integer
    Private mMaxFuelTankSize As Integer
    Private mMaxSaleSize As Integer
    Private mMaxTrxSize As Integer

#End Region 'Attributes

#Region "Constructors"
    '******************************************************************
    'Constructors
    '******************************************************************

    'These are all public.

    '********** Default constructor
    '             - no parameters

    Public Sub New()

        MyBase.New

        _numCustomers = 0
        _numProducts = 0
        _numSales = 0
        _numTransactions = 0

    End Sub

    '********** Special constructor(s)
    '             - typically constructors have parameters 
    '               that are used to initialize attributes

    Public Sub New(
           ByVal pFillingStationName As String,
           ByVal pNumCustomers As Integer,
           ByVal pNumProducts As Integer,
           ByVal pNumSales As Integer,
           ByVal pNumTransactions As Integer
           )
        'ByVal pCustomers As Array,
        'ByVal pProducts As Array,
        'ByVal pSales As Array,
        'ByVal pTransactions As Array
        ')

        MyBase.New

        _fillingStationName = pFillingStationName
        _numCustomers = pNumCustomers
        _numProducts = pNumProducts
        _numSales = pNumSales
        _numTransactions = pNumTransactions
        '_customers = pCustomers
        '_products = pProducts
        '_sales = pSales
        '_transactions = pTransactions

    End Sub

    '********** Copy constructor(s)
    '             - one parameter, an object of the same class

#End Region 'Constructors

#Region "Get/Set Methods"
    '******************************************************************
    'Get/Set Methods
    '******************************************************************

    '********** Public Get/Set Methods
    '             - call private get/set methods to implement

    Public ReadOnly Property fillingStationName As String
        Get
            Return _fillingStationName
        End Get
    End Property

    Public Property numCustomers As Integer
        Get
            Return _numCustomers
        End Get
        Set(ByVal pValue As Integer)
            _numCustomers = pValue
        End Set
    End Property

    Public Property numProducts As Integer
        Get
            Return _numProducts
        End Get
        Set(ByVal pValue As Integer)
            _numProducts = pValue
        End Set
    End Property
    Public Property numSales As Integer
        Get
            Return _numSales
        End Get
        Set(ByVal pValue As Integer)
            _numSales = pValue
        End Set
    End Property

    Public Property numTransactions As Integer
        Get
            Return _numTransactions
        End Get
        Set(ByVal pValue As Integer)
            _numTransactions = pValue
        End Set
    End Property

    Public Property numFuelTanks As Integer
        Get
            Return _numFuelTanks
        End Get
        Set(ByVal pValue As Integer)
            _numFuelTanks = pValue
        End Set
    End Property

    Public ReadOnly Property ithCustomers(ByVal pI As Integer) As LoyaltyCustomer
        Get
            Return _ithCustomers(pI)
        End Get
        'Set(ByVal pValue As LoyaltyCustomer)
        '    _ithCustomers(pI) = pValue
        'End Set
    End Property

    Public ReadOnly Property ithProducts(ByVal pI As Integer) As Product
        Get
            Return _ithProducts(pI)
        End Get
        'Set(ByVal pValue As Product)
        '    _ithProducts(pI) = pValue
        'End Set
    End Property

    Public ReadOnly Property ithFuelTanks(ByVal pI As Integer) As FuelTank
        Get
            Return _ithFuelTanks(pI)
        End Get
        'Set(ByVal pValue As FuelTank)
        '    _ithFuelTanks(pI) = pValue
        'End Set
    End Property

    Public ReadOnly Property ithSales(ByVal pI As Integer) As Sale
        Get
            Return _ithSales(pI)
        End Get
        'Set(ByVal pValue As Sale)
        '    _ithSales(pI) = pValue
        'End Set
    End Property

    Public ReadOnly Property ithTrxs(ByVal pI As Integer) As Transaction
        Get
            Return _ithTrxs(pI)
        End Get
        'Set(ByVal pValue As Transaction)
        '    _ithTrxs(pI) = pValue
        'End Set
    End Property

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _fillingStationName As String
        Get
            Return mFillingStationName
        End Get
        Set(ByVal pValue As String)
            mFillingStationName = pValue
        End Set
    End Property

    Private Property _numCustomers As Integer
        Get
            Return mNumCustomers
        End Get
        Set(ByVal pValue As Integer)
            mNumCustomers = pValue
        End Set
    End Property

    Private Property _numProducts As Integer
        Get
            Return mNumProducts
        End Get
        Set(ByVal pValue As Integer)
            mNumProducts = pValue
        End Set
    End Property

    Private Property _numSales As Integer
        Get
            Return mNumSales
        End Get
        Set(ByVal pValue As Integer)
            mNumSales = pValue
        End Set
    End Property

    Private Property _numTransactions As Integer
        Get
            Return mNumTransactions
        End Get
        Set(ByVal pValue As Integer)
            mNumTransactions = pValue
        End Set
    End Property

    Private Property _numFuelTanks As Integer
        Get
            Return mNumFuelTanks
        End Get
        Set(ByVal pValue As Integer)
            mNumFuelTanks = pValue
        End Set
    End Property

    Private Property _ithCustomers(ByVal pI As Integer) As LoyaltyCustomer
        Get
            Try
                If pI >= 0 And pI < _maxCustomerSize Then
                    Return mCustomers(pI)
                Else
                    Throw New IndexOutOfRangeException
                End If
            Catch ex As Exception
                MessageBox.Show("Error")
            End Try
        End Get
        Set(ByVal pvalue As LoyaltyCustomer)
            If pI >= 0 And pI < _maxCustomerSize Then
                mCustomers(pI) = pvalue
            Else
                Throw New IndexOutOfRangeException
            End If
        End Set
    End Property

    Private Property _ithProducts(ByVal pI As Integer) As Product
        Get
            Try
                If pI >= 0 And pI < _maxProductSize Then
                    Return mProducts(pI)
                Else
                    Throw New IndexOutOfRangeException
                End If
            Catch ex As Exception
                MessageBox.Show("Error")
            End Try
        End Get
        Set(ByVal pvalue As Product)
            If pI >= 0 And pI < _maxProductSize Then
                mProducts(pI) = pvalue
            Else
                Throw New IndexOutOfRangeException
            End If
        End Set
    End Property

    Private Property _ithFuelTanks(ByVal pI As Integer) As FuelTank
        Get
            Try
                If pI >= 0 And pI < _maxFuelTankSize Then
                    Return mFuelTanks(pI)
                Else
                    Throw New IndexOutOfRangeException
                End If
            Catch ex As Exception
                MessageBox.Show("Error")
            End Try
        End Get
        Set(ByVal pvalue As FuelTank)
            If pI >= 0 And pI < _maxFuelTankSize Then
                mFuelTanks(pI) = pvalue
            Else
                Throw New IndexOutOfRangeException
            End If
        End Set
    End Property

    Private Property _ithSales(ByVal pI As Integer) As Sale
        Get
            Try
                If pI >= 0 And pI < _maxSaleSize Then
                    Return mSales(pI)
                Else
                    Throw New IndexOutOfRangeException
                End If
            Catch ex As Exception
                MessageBox.Show("Error")
            End Try
        End Get
        Set(ByVal pvalue As Sale)
            If pI >= 0 And pI < _maxSaleSize Then
                mSales(pI) = pvalue
            Else
                Throw New IndexOutOfRangeException
            End If
        End Set
    End Property

    Private Property _ithTrxs(ByVal pI As Integer) As Transaction
        Get
            If pI >= 0 And pI < _maxTrxSize Then
                Return mTransactions(pI)
            Else
                Throw New IndexOutOfRangeException
            End If
        End Get
        Set(ByVal pvalue As Transaction)
            If pI >= 0 And pI < _maxTrxSize Then
                mTransactions(pI) = pvalue
            Else
                Throw New IndexOutOfRangeException
            End If
        End Set
    End Property

    Private Property _maxProductSize As Integer
        Get
            Return mMaxProductSize
        End Get
        Set(ByVal pValue As Integer)
            mMaxProductSize = pValue
        End Set
    End Property

    Private Property _maxFuelTankSize As Integer
        Get
            Return mMaxFuelTankSize
        End Get
        Set(ByVal pValue As Integer)
            mMaxFuelTankSize = pValue
        End Set
    End Property

    Private Property _maxCustomerSize As Integer
        Get
            Return mMaxCustomerSize
        End Get
        Set(ByVal pValue As Integer)
            mMaxCustomerSize = pValue
        End Set
    End Property

    Private Property _maxSaleSize As Integer
        Get
            Return mMaxSaleSize
        End Get
        Set(ByVal pValue As Integer)
            mMaxSaleSize = pValue
        End Set
    End Property

    Private Property _maxTrxSize As Integer
        Get
            Return mMaxTrxSize
        End Get
        Set(ByVal pValue As Integer)
            mMaxTrxSize = pValue
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

    Public Sub readFromFile(
            ByVal pFile As String
            )

        'declares file variables
        Dim inputFile As StreamReader
        Dim theLine As String
        Dim theField() As String
        Dim theTrxDate As Date

        'declares customer variables
        'Dim theCustomer As LoyaltyCustomer
        Dim theCustomerId As String
        Dim theCustomerName As String
        Dim thePhoneNum As String
        Dim theSecurityPin As String
        Dim theSecurityPinConfirm As String
        Dim theMemberSince As Date

        'declares product variables
        'Dim theProduct As Product
        Dim theProductId As String
        Dim theProductName As String
        Dim theProductType As ProductType
        Dim theUnitOfMeasure As String
        Dim thepricePerUnit As Decimal
        Dim theLoyaltyDiscount As Decimal
        Dim theRewardDiscount As Decimal
        Dim theTaxRate As Decimal
        Dim theFuelTankId As String
        Dim theMaxCapacity As Decimal

        'declares sale variables
        Dim theSaleId As String
        Dim theSaleDate As Date
        Dim theQuantityPurchased As Decimal
        Dim theSubTotal As Decimal
        Dim theTaxAmount As Decimal
        Dim theTotal As Decimal

        'declares transaction variables
        Dim theTrxId1 As String
        Dim theTrxId2 As String
        Dim theTrxText As String

        'assign a file
        Try
            inputFile = New StreamReader(pFile)
        Catch ex As Exception
            Throw New FileNotFoundException
            Exit Sub
        End Try

        Do While Not inputFile.EndOfStream

            theLine = inputFile.ReadLine
            'MessageBox.Show(theLine.First)

            If theLine = "" Then
                theTrxText = ""
                _createTransaction(
                    "Trx-COMMENT" & _numTransactions.ToString,
                    theTrxText,
                    True
                    )
            ElseIf theLine.First = "#" Then
                theTrxText = theLine
                _createTransaction(
                    "Trx-COMMENT" & _numTransactions.ToString,
                    theTrxText,
                    True
                    )
            ElseIf IsNumeric(theLine.First) Then

                If theLine.ToUpper.Contains("CREATE") Then

                    If theLine.ToUpper.Contains("LOYALTYCUSTOMER") Then

                        'split line into substring
                        theField = Split(theLine, "; ")
                        theTrxDate = Date.ParseExact(
                        Trim(theField(0)), "yyyyMMdd", CultureInfo.InvariantCulture)
                        theTrxId1 = Trim(theField(4))
                            theCustomerId = Trim(theField(5))
                            theCustomerName = Trim(theField(6))
                            thePhoneNum = Trim(theField(7))
                            theSecurityPin = Trim(theField(8))
                            theSecurityPinConfirm = Trim(theField(9))
                        theMemberSince = Date.ParseExact(
                            Trim(theField(10)), "yyyyMMdd", CultureInfo.InvariantCulture)

                        'create an new customer
                        _createLoyaltyCustomer(
                            theCustomerId,
                            theTrxDate,
                            theTrxId1,
                            theCustomerName,
                            thePhoneNum,
                            theSecurityPin,
                            theSecurityPinConfirm,
                            theMemberSince,
                            _calMemberSinceAge(theMemberSince),
                            0
                            )

                    ElseIf theLine.ToUpper.Contains("PRODUCT") Then

                        'split line into substring
                        theField = Split(theLine, "; ")
                        If theLine.ToUpper.Contains("FUEL") Then
                            theProductType = ProductType.FUEL
                            theTrxDate = Date.ParseExact(
                                Trim(theField(0)), "yyyyMMdd", CultureInfo.InvariantCulture)
                            theTrxId1 = Trim(theField(5))
                            theProductId = Trim(theField((6)))
                            theProductName = Trim(theField(7))
                            theTaxRate = CDec(Trim(theField(8)))
                            theUnitOfMeasure = Trim(theField(9))
                            thepricePerUnit = CDec(Trim(theField(10)))
                            theLoyaltyDiscount = CDec(Trim(theField(11)))
                            theRewardDiscount = CDec(Trim(theField(12)))
                            theTrxId2 = Trim(theField(13))
                            theFuelTankId = Trim(theField(14))
                            theMaxCapacity = CDec(Trim(theField(15)))
                        ElseIf theLine.ToUpper.Contains("CARWASH") Then
                            theProductType = ProductType.CARWASH
                            theTrxDate = Date.ParseExact(
                                Trim(theField(0)), "yyyyMMdd", CultureInfo.InvariantCulture)
                            theTrxId1 = Trim(theField(5))
                            theProductId = Trim(theField((6)))
                            theProductName = Trim(theField(7))
                            theTaxRate = CDec(Trim(theField(8)))
                            theUnitOfMeasure = Trim(theField(9))
                            thepricePerUnit = CDec(Trim(theField(10)))
                            theLoyaltyDiscount = CDec(Trim(theField(11)))
                            theRewardDiscount = 0
                            theTrxId2 = ""
                            theFuelTankId = ""
                            theMaxCapacity = 0
                        ElseIf theLine.ToUpper.Contains("MISC") Then
                            theProductType = ProductType.CARWASH
                            theTrxDate = Date.ParseExact(
                                Trim(theField(0)), "yyyyMMdd", CultureInfo.InvariantCulture)
                            theTrxId1 = Trim(theField(5))
                            theProductId = Trim(theField((6)))
                            theProductName = Trim(theField(7))
                            theTaxRate = CDec(Trim(theField(8)))
                            theUnitOfMeasure = Trim(theField(9))
                            thepricePerUnit = CDec(Trim(theField(10)))
                            theLoyaltyDiscount = 0
                            theRewardDiscount = 0
                            theTrxId2 = ""
                            theFuelTankId = ""
                            theMaxCapacity = 0
                        End If

                        'create an new product
                        _createProduct(
                            theProductId,
                            theTrxDate,
                            theTrxId1,
                            theTrxId2,
                            theProductName,
                            theProductType,
                            theUnitOfMeasure,
                            thepricePerUnit,
                            theLoyaltyDiscount,
                            theRewardDiscount,
                            theTaxRate,
                            theFuelTankId,
                            theMaxCapacity,
                            theMaxCapacity
                            )

                    ElseIf theLine.ToUpper.Contains("SALE") Then

                        'split line into substring
                        theField = Split(theLine, "; ")
                        theSaleDate = Date.ParseExact(
                                Trim(theField(0)), "yyyyMMdd", CultureInfo.InvariantCulture)
                        theTrxId1 = Trim(theField(4))
                        theSaleId = Trim(theField(5))
                        theCustomerId = Trim(theField(6))
                        theProductId = Trim(theField(7))
                        theQuantityPurchased = CDec(Trim(theField(8)))
                        thepricePerUnit = CDec(Trim(theField(9)))
                        theLoyaltyDiscount = CDec(Trim(theField(10)))
                        theTaxRate = CDec(Trim(theField(11)))

                        'create an new sale
                        _createSale(
                            theSaleId,
                             theSaleDate,
                            theTrxId1,
                            theCustomerId,
                            theProductId,
                            theQuantityPurchased
                            )
                    End If


                ElseIf theLine.ToUpper.Contains("MODIFY") Then

                    'split line into substring
                    theField = Split(theLine, ";")
                    theTrxDate = Date.ParseExact(
                                Trim(theField(0)), "yyyyMMdd", CultureInfo.InvariantCulture)
                    theTrxId1 = Trim(theField(4))
                    theProductId = Trim(theField(5))
                    theProductName = Trim(theField(6))
                    theTaxRate = CDec(Trim(theField(7)))
                    theUnitOfMeasure = Trim(theField(8))
                    thepricePerUnit = CDec(Trim(theField(9)))

                    If theField.Length = 12 Then
                        theLoyaltyDiscount = CDec(Trim(theField(10)))
                        theRewardDiscount = CDec(Trim(theField(11)))
                    ElseIf theField.Length = 11 Then
                        theLoyaltyDiscount = CDec(Trim(theField(10)))
                        theRewardDiscount = 0
                    ElseIf theField.Length = 10 Then
                        theLoyaltyDiscount = 0
                        theRewardDiscount = 0
                    End If

                    'modify product
                    _modifyProduct(
                        theProductId,
                        theTrxDate,
                        theTrxId1,
                        theProductName,
                        theUnitOfMeasure,
                        thepricePerUnit,
                        theLoyaltyDiscount,
                        theRewardDiscount, theTaxRate
                        )

                End If
            End If
        Loop

        inputFile.Close()

    End Sub 'readFromFile(pFile)

    Public Sub writeToFile()

        'declares variable
        Dim outputFile As StreamWriter
        Dim outputFileError As StreamWriter
        Dim i As Integer

        'write to a file
        outputFile = New StreamWriter("Transactions-out.txt")
        outputFileError = New StreamWriter("Transactions-errors.txt")

        For i = 0 To _numTransactions - 1
            If _ithTrxs(i).isError = False Then
                outputFile.WriteLine(
                    _ithTrxs(i).transactionText
                    )
            Else
                outputFileError.WriteLine(
                    _ithTrxs(i).transactionText
                    )
            End If
        Next

        outputFile.Close()
        outputFileError.Close()

    End Sub 'writeToFile()

    Public Function findCustomer(
            ByVal pCustomerId As String
            ) _
        As _
            LoyaltyCustomer

        Dim locationFound As Integer

        Return _findCustomer(
            pCustomerId,
            locationFound
            )

    End Function 'findCustomer(pCustomerId)

    Public Function findProduct(
           ByVal pProductId As String
           ) _
        As _
            Product

        Dim locationFound As Integer

        Return _findProduct(
            pProductId,
            locationFound
            )

    End Function 'findProduct(pProductId)

    Public Function findFuelTank(
           ByVal pFuelTankId As String
           ) _
       As _
           FuelTank

        Dim locationFound As Integer

        Return _findFuelTank(
            pFuelTankId,
            locationFound
            )

    End Function 'findFuelTank(pFuelTankId,pListBox)

    Public Function findSale(
           ByVal pSaleId As String
           ) _
       As _
           Sale

        Dim locationFound As Integer

        Return _findSale(
            pSaleId,
            locationFound
            )

    End Function 'findSale(pSaleId)

    Public Function findTrx(
           ByVal pTrxId As String
           ) _
       As _
           Transaction

        Dim locationFound As Integer

        Return _findTrx(
            pTrxId,
            locationFound
            )

    End Function 'findTrx(pTrxId)

    Public Function createLoyaltyCustomer(
           ByVal pCustomerId As String,
           ByVal pTrxDate As Date,
           ByVal pTrxId As String,
           ByVal pName As String,
           ByVal pPhone As String,
           ByVal pSecurityPin1 As String,
           ByVal pSecurityPin2 As String,
           ByVal pMemberSince As Date,
           ByVal pMembershipAge As Integer,
           ByVal pAccruedRewardGallon As Decimal
           ) _
       As _
           LoyaltyCustomer

        Return _createLoyaltyCustomer(
            pCustomerId,
            pTrxDate,
            pTrxId,
            pName,
            pPhone,
            pSecurityPin1,
            pSecurityPin2,
            pMemberSince,
            pMembershipAge,
            pAccruedRewardGallon
            )

    End Function 'createLoyaltyCustomer()

    Public Function createProduct(
           ByVal pProductId As String,
           ByVal pTrxDate As Date,
           ByVal pTrxId1 As String,
           ByVal pTrxId2 As String,
           ByVal pProductName As String,
           ByVal pProductType As ProductType,
           ByVal pUnitOfMeasure As String,
           ByVal pPricePerUnit As Decimal,
           ByVal pLoyaltyDiscountPerUnit As Decimal,
           ByVal pRewardDiscountPerUnit As Decimal,
           ByVal pTaxRate As Decimal,
           ByVal pFuelTankId As String,
           ByVal pCurrentFuelQuantity As Decimal,
           ByVal pMaxFuelQuantiy As Decimal) _
       As _
           Product

        Return _createProduct(
            pProductId,
            pTrxDate,
            pTrxId1,
            pTrxId2,
            pProductName,
            pProductType,
            pUnitOfMeasure,
            pPricePerUnit,
            pLoyaltyDiscountPerUnit,
            pRewardDiscountPerUnit,
            pTaxRate,
            pFuelTankId,
            pCurrentFuelQuantity,
            pMaxFuelQuantiy
            )

    End Function 'createProduct()

    Public Function orderFuel(
           ByVal pProductId As String,
           ByVal pTrxDate As Date,
           ByVal pTrxId As String,
           ByVal pFuelAmount As Decimal
           ) _
        As _
            Decimal

        Return _orderFuel(
            pProductId,
            pTrxDate,
            pTrxId,
            pFuelAmount
            )

    End Function 'orderFuel(pProductId,pTrxId,pFuelAmount)

    Public Function createSale(
           ByVal pSaleId As String,
           ByVal pTrxDate As Date,
           ByVal pTrxId As String,
           ByRef pCustomerId As String,
           ByRef pProductId As String,
           ByVal pQuantityPurchased As Decimal
           ) _
       As _
           Sale

        Return _createSale(
            pSaleId,
            pTrxDate,
            pTrxId,
            pCustomerId,
            pProductId,
            pQuantityPurchased
            )

    End Function 'createSale()

    Public Function createTransaction(
           ByVal pId As String,
           ByVal pTransactionText As String,
           ByVal pIsError As Boolean
           ) _
       As _
           Transaction

        Return _createTransaction(
            pId,
            pTransactionText,
            pIsError
            )

    End Function 'createTransaction()

    Public Sub modityProduct(
           ByVal pProductId As String,
           ByVal pTrxDate As Date,
           ByVal pTrxId As String,
           ByVal pProductName As String,
           ByVal pUnitOfMeasure As String,
           ByVal pPricePerUnit As Decimal,
           ByVal pLoyaltyDiscountPerUnit As Decimal,
           ByVal pRewardDiscountPerUnit As Decimal,
           ByVal pTaxRate As Decimal
           )

        _modifyProduct(
            pProductId,
            pTrxDate,
            pTrxId,
            pProductName,
            pUnitOfMeasure,
            pPricePerUnit,
            pLoyaltyDiscountPerUnit,
            pRewardDiscountPerUnit,
            pTaxRate
            )

    End Sub 'modifyProduct(pProductId,pTrxId,pPricePerUnit,pLoyaltyDiscountPerUnit,pRewardDiscountPerUnit,pTaxRate)

    Public Function calProductTypeTotalSale( 'summary tab
            ByVal pProductType As ProductType
            ) _
        As _
            Decimal

        Dim theTotalSale As Decimal
        Dim i As Integer
        For i = 0 To _numSales - 1
            If _ithSales(i).product.productType = pProductType Then
                theTotalSale += _ithSales(i).totalAmount
                Return theTotalSale
            End If
        Next

    End Function '_calProductTypeTotalSale(pProductType)

    Public Function calProductTypeTotalTaxAmount(
            ByVal pProductType As ProductType
            ) _
        As _
            Decimal

        Dim theTotalTaxAmount As Decimal
        Dim i As Integer
        For i = 0 To _numSales - 1
            If _ithSales(i).product.productType = pProductType Then
                theTotalTaxAmount += _ithSales(i).taxAmount
                Return theTotalTaxAmount
            End If
        Next

    End Function '_calProductTypeTotalTaxAmount(pProductType)

    Public Function calAverageAmountPerSale(
            ByVal pFillingStation As FillingStation
            ) _
        As _
            Decimal

        Dim theTotalSale As Decimal
        Dim theAverageAmount As Decimal
        Dim i As Integer
        For i = 0 To _numSales - 1
            theTotalSale += _ithSales(i).totalAmount
            theAverageAmount = theTotalSale / _numSales
        Next
        Return theAverageAmount

    End Function '_calAverageAmountPerSale(pFillingStation)

    Public Function calTotalAmountPerCustomer(
            ByVal pCustomer As LoyaltyCustomer
            ) _
        As _
            Decimal

        Dim theTotalAmount As Decimal
        Dim i As Integer
        For i = 0 To _numSales - 1
            If _ithSales(i).customer IsNot Nothing Then
                If _ithSales(i).customer.Equals(pCustomer) Then
                    theTotalAmount += _ithSales(i).totalAmount
                End If
            End If

        Next
        Return theTotalAmount

    End Function '_calTotalAmountPerCustomer(pCustomer)

    Public Function calTotalNumSalePerCustomer(
            ByVal pCustomer As LoyaltyCustomer
            ) _
        As _
            Integer

        Dim theTotalNum As Integer
        Dim i As Integer
        For i = 0 To _numSales - 1
            If _ithSales(i).customer.Equals(pCustomer) Then
                theTotalNum += 1
            End If
        Next
        Return theTotalNum

    End Function '_calTotalNumSalePerCustomer(pCustomer)

    Public Function calPercentageSalePerProductType(
            ByVal pProductType As ProductType
            ) _
        As _
            Decimal

        Dim theTotalAmountProductType As Decimal
        Dim theTotalAmount As Decimal
        Dim i As Integer
        For i = 0 To _numSales - 1
            If _ithSales(i).product.productType = pProductType Then
                theTotalAmountProductType += _ithSales(i).totalAmount
            End If
            theTotalAmount += _ithSales(i).totalAmount
        Next
        Return theTotalAmountProductType / theTotalAmount

    End Function '_calPercentageSalePerProductType(pProductType)

    Public Function calPercentageNumSalePerProductType(
            ByVal pProductType As ProductType
            ) _
        As _
            Decimal

        Dim theTotalNumProductType As Integer
        Dim i As Integer
        For i = 0 To _numSales - 1
            If _ithSales(i).product.productType = pProductType Then
                theTotalNumProductType += 1
            End If
        Next
        Return CDec(theTotalNumProductType / _numSales)

    End Function '_calPercentageNumSalePerProductType(pProductType)

    Public Function getSmallestAmountInSale(
            ByVal pFillingStation As FillingStation
            ) _
        As _
            Decimal

        Dim theSmallestAmount As Decimal = _ithSales(0).totalAmount
        Dim i As Integer
        For i = 0 To _numSales - 1
            If _ithSales(i).totalAmount < theSmallestAmount Then
                theSmallestAmount = _ithSales(i).totalAmount
            End If
        Next

        Return theSmallestAmount

    End Function '_getSmallestAmountInSale(pFillingStation)

    Public Function getBiggestAmountInSale(
            ByVal pFillingStation As FillingStation
            ) _
        As _
            Decimal

        Dim theBiggestAmount As Decimal = _ithSales(0).totalAmount
        'Dim theSale As Sale
        Dim i As Integer
        For i = 0 To _numSales - 1
            If _ithSales(i).totalAmount > theBiggestAmount Then
                theBiggestAmount = _ithSales(i).totalAmount
                'theSale = _ithSales(i)
            End If
        Next

        Return theBiggestAmount

    End Function '_getBiggestAmountInSale(pFillingStation)

    Public Function getSmallestSaleProductType(
           ByVal pFillingStation As FillingStation
           ) _
        As _
            ProductType

        Dim i As Integer
        Dim theSmallestSale As Decimal = _ithSales(0).totalAmount
        Dim theIndex As Integer

        For i = 0 To _numSales - 1
            If _ithSales(i).totalAmount < theSmallestSale Then
                theSmallestSale = _ithSales(i).totalAmount
                theIndex = i
            End If
        Next

        Return _ithSales(theIndex).product.productType

    End Function 'getSmallestSaleProductType(pFillingStation)

    Public Function getLargestSaleProductType(
           ByVal pFillingStation As FillingStation
           ) _
        As _
            ProductType

        Dim i As Integer
        Dim theLargestSale As Decimal = _ithSales(0).totalAmount
        Dim theIndex As Integer

        For i = 0 To _numSales - 1
            If _ithSales(i).totalAmount > theLargestSale Then
                theLargestSale = _ithSales(i).totalAmount
                theIndex = i
            End If
        Next

        Return _ithSales(theIndex).product.productType

    End Function 'getlargestSaleProductType(pFillingStation)

    Public Function transferFromEnum(
            ByVal pProductType As ProductType
            ) _
        As _
            String

        Dim theProductType As String

        theProductType = CType(pProductType, ProductType).ToString

        Return theProductType

    End Function '_transferFromEnum(pProductType)

    Public Iterator Function iterateCustomer(
           ) _
        As _
           IEnumerable(Of Object)

        Dim theObject As Object

        For Each theObject In _iterateCustomer()
            Yield theObject
        Next theObject

    End Function 'iterateCustomer()

    Public Iterator Function iterateProduct(
           ) _
        As _
           IEnumerable(Of Object)

        Dim theObject As Object

        For Each theObject In _iterateProduct()
            Yield theObject
        Next theObject

    End Function 'iterateProduct()

    Public Iterator Function iterateFuelTank(
           ) _
        As _
           IEnumerable(Of Object)

        Dim theObject As Object

        For Each theObject In _iterateFuelTank()
            Yield theObject
        Next theObject

    End Function 'iterateFuelTank()

    Public Iterator Function iterateSale(
           ) _
        As _
           IEnumerable(Of Object)

        Dim theObject As Object

        For Each theObject In _iterateSale()
            Yield theObject
        Next theObject

    End Function 'iterateSale()

    Public Iterator Function iterateTrx(
           ) _
        As _
           IEnumerable(Of Object)

        Dim theObject As Object

        For Each theObject In _iterateTrx()
            Yield theObject
        Next theObject

    End Function 'iterateTrx()

    Public Overrides Function ToString() As String

        Return _toString()

    End Function 'ToString()

    '********** Private Non-Shared Behavioral Methods

    Private Function _findCustomer(
            ByVal pCustomerId As String,
            ByRef pLocationFound As Integer
            ) _
        As _
            LoyaltyCustomer

        'Dim i As Integer
        For pLocationFound = 0 To _numCustomers - 1
            If _ithCustomers(pLocationFound).id = pCustomerId Then
                Return _ithCustomers(pLocationFound)
            End If
        Next plocationfound

        Return Nothing

    End Function '_findCustomer(pCustomerId,pLocationFound)

    Private Function _findProduct(
            ByVal pProductId As String,
            ByRef pLocationFound As Integer
            ) _
        As _
            Product


        For pLocationFound = 0 To _numProducts - 1
            If _ithProducts(pLocationFound).id = pProductId Then
                Return _ithProducts(pLocationFound)
            End If
        Next pLocationFound

        Return Nothing

    End Function '_findComboboxProduct(pProductId,pLocationFound)

    Private Function _findFuelTank(
            ByVal pFuelTankId As String,
            ByRef pLocationFound As Integer
            ) _
        As _
            FuelTank

        For pLocationFound = 0 To _numFuelTanks - 1
            If _ithFuelTanks(pLocationFound).id = pFuelTankId Then
                Return _ithFuelTanks(pLocationFound)
            End If
        Next pLocationFound

        Return Nothing

    End Function '_findFuelTank(pFuelTankId,pLocationFound)

    Private Function _findSale(
            ByVal pSaleId As String,
            ByRef pLocationFound As Integer
            ) _
        As _
            Sale


        For pLocationFound = 0 To _numSales - 1
            If _ithSales(pLocationFound).id = pSaleId Then
                Return _ithSales(pLocationFound)
            End If
        Next pLocationFound

        Return Nothing

    End Function '_findSale(pSaleId,pLocationFound)

    Private Function _findTrx(
            ByVal pTrxId As String,
            ByRef pLocationFound As Integer
            ) _
        As _
            Transaction

        For pLocationFound = 0 To _numTransactions - 1
            If _ithTrxs(pLocationFound).id = pTrxId Then
                Return _ithTrxs(pLocationFound)
            End If
        Next pLocationFound

        Return Nothing

    End Function '_findTrx(pTrxId,pLocationFound)

    Private Iterator Function _iterateCustomer(
            ) _
        As _
            IEnumerable(Of Object)

        Dim i As Integer

        For i = 0 To _numCustomers - 1
            Yield _ithCustomers(i)
        Next

    End Function '_iterateCustomer()

    Private Iterator Function _iterateProduct(
            ) _
        As _
        IEnumerable(Of Object)

        Dim i As Integer

        For i = 0 To _numProducts - 1
            Yield _ithProducts(i)
        Next

    End Function '_iterateProduct()

    Private Iterator Function _iterateFuelTank(
            ) _
        As _
            IEnumerable(Of Object)

        Dim i As Integer

        For i = 0 To _numFuelTanks - 1
            Yield _ithFuelTanks(i)
        Next

    End Function '_iterateFuelTank()

    Private Iterator Function _iterateSale(
            ) _
        As _
            IEnumerable(Of Object)

        Dim i As Integer

        For i = 0 To _numSales - 1
            Yield _ithSales(i)
        Next

    End Function '_iterateSale()

    Private Iterator Function _iterateTrx(
            ) _
        As _
            IEnumerable(Of Object)

        Dim i As Integer

        For i = 0 To _numTransactions - 1
            Yield _ithTrxs(i)
        Next i

    End Function '_iterateTrx()

    Private Function _createLoyaltyCustomer(
            ByVal pCustomerId As String,
            ByVal pTrxDate As Date,
            ByVal pTrxId As String,
            ByVal pName As String,
            ByVal pPhone As String,
            ByVal pSecurityPin1 As String,
            ByVal pSecurityPin2 As String,
            ByVal pMemberSince As Date,
            ByVal pMembershipAge As Integer,
            ByVal pAccruedRewardGallon As Decimal
            ) _
        As _
            LoyaltyCustomer

        'declares local variables
        Dim theLocationFound As Integer
        Dim theCustomer As LoyaltyCustomer
        Dim theTrxText As String = pTrxDate.ToString("yyyyMMdd") & "; " & pTrxDate.ToString("hhmm") & "; " & "LoyaltyCustomer; " _
            & "Create; " & pTrxId & "; " & pCustomerId & "; " _
            & pName & "; " & pPhone & "; " & pSecurityPin1 _
            & "; " & pSecurityPin2 & "; " & pMemberSince.ToString("yyyyMMdd")

        'validates
        If _findCustomer(pCustomerId, theLocationFound) IsNot Nothing Then 'check duplicate customer ID
            MessageBox.Show("The customer ID is existed, please enter another new one.")
            _createTransaction(pTrxId, theTrxText, True)
            Exit Function
        End If
        If _findTrx(pTrxId, theLocationFound) IsNot Nothing Then
            MessageBox.Show("The transaction ID is existed, please enter another new one.")
            Exit Function
        End If
        If pSecurityPin1.Equals(pSecurityPin2) = False Then
            MessageBox.Show("The two security pins are not the same, please enter again.")
            _createTransaction(pTrxId, theTrxText, True)
            Exit Function
        End If
        If pMemberSince > Now Then
            MessageBox.Show("Please select an earlier day than today.")
            _createTransaction(pTrxId, theTrxText, True)
            Exit Function
        End If

        'creates a new customer
        theCustomer = New LoyaltyCustomer(
                            pCustomerId,
                            pName,
                            pPhone,
                            pSecurityPin1,
                            pMemberSince,
                            pMembershipAge,
                            pAccruedRewardGallon
                            )

        'checking array size
        If _numCustomers >= _maxCustomerSize Then
            _maxCustomerSize += mARRAY_SIZE__INCREATMENT_DEFAULT
            ReDim Preserve mCustomers(_maxCustomerSize - 1)
        End If

        'adding into array
        Try
            _ithCustomers(_numCustomers) = theCustomer
        Catch ex As Exception
            Throw New IndexOutOfRangeException
        End Try

        'counting
        _numCustomers += 1

        'adds transaction
        _createTransaction(
            pTrxId,
            theTrxText,
            False)

        RaiseEvent FillingStation_CustomerAdded(
            Me,
            New FillingStation_EventArgs_loyaltyCustomerAdded(theCustomer)
            )

        Return theCustomer

    End Function '_createCustomer(pId, pName,pPhone,pSecurity,pMemberSince,pMembershipAge,pAccrueRewardGallon)

    Private Function _createProduct(
            ByVal pProductId As String,
            ByVal pTrxDate As Date,
            ByVal pTrxId1 As String,
            ByVal pTrxId2 As String,
            ByVal pProductName As String,
            ByVal pProductType As ProductType,
            ByVal pUnitOfMeasure As String,
            ByVal pPricePerUnit As Decimal,
            ByVal pLoyaltyDiscountPerUnit As Decimal,
            ByVal pRewardDiscountPerUnit As Decimal,
            ByVal pTaxRate As Decimal,
            ByVal pFuelTankId As String,
            ByVal pCurrentFuelQuantity As Decimal,
            ByVal pMaxFuelQuantiy As Decimal
            ) _
        As _
            Product

        'declares local variables
        Dim theLocationFound As Integer
        Dim theProduct As Product
        Dim theFuelTank As FuelTank
        Dim theTrxText As String

        'prepares transaction text
        If pProductType = ProductType.FUEL Then
            theTrxText = pTrxDate.ToString("yyyyMMdd") & "; " & pTrxDate.ToString("hhmm") & "; " & "Product; " _
                & "Create; " & transferFromEnum(pProductType) & "; " & pTrxId1 & "; " & pProductId & "; " _
                & pProductName & "; " & pTaxRate.ToString & "; " & pUnitOfMeasure _
                & "; " & pPricePerUnit & "; " & pLoyaltyDiscountPerUnit.ToString & "; " _
                & pRewardDiscountPerUnit.ToString & "; " & pTrxId2 & "; " & pFuelTankId _
                & "; " & pMaxFuelQuantiy.ToString
        ElseIf pProductType = ProductType.CARWASH Then
            theTrxText = pTrxDate.ToString("yyyyMMdd") & "; " & pTrxDate.ToString("hhmm") & "; " & "Product; " _
                & "Create; " & transferFromEnum(pProductType) & "; " & pTrxId1 & "; " & pProductId & "; " _
                & pProductName & "; " & pTaxRate.ToString & "; " & pUnitOfMeasure _
                & "; " & pPricePerUnit & "; " & pLoyaltyDiscountPerUnit.ToString & "; " _
                & pRewardDiscountPerUnit.ToString & "; "
        Else
            theTrxText = pTrxDate.ToString("yyyyMMdd") & "; " & pTrxDate.ToString("hhmm") & "; " & "Product; " _
                & "Create; " & transferFromEnum(pProductType) & "; " & pTrxId1 & "; " & pProductId & "; " _
                & pProductName & "; " & pTaxRate.ToString & "; " & pUnitOfMeasure _
                & "; " & pPricePerUnit & "; "
        End If

        'validates duplicate product ID
        If _findProduct(pProductId, theLocationFound) IsNot Nothing Then
            MessageBox.Show("The customer ID is existed, please enter another new one.")
            If pProductType = ProductType.FUEL Then
                _createTransaction(pTrxId1, theTrxText, True)
                _createTransaction(pTrxId2, theTrxText, True)
                Exit Function
            Else
                _createTransaction(pTrxId1, theTrxText, True)
                Exit Function
            End If
        End If

        If pProductType = ProductType.FUEL Then
            theProduct = New Product(
                pProductId,
                pProductName,
                pProductType,
                pUnitOfMeasure,
                pPricePerUnit,
                pLoyaltyDiscountPerUnit,
                pRewardDiscountPerUnit,
                pTaxRate
                )
            theFuelTank = theProduct.createFuelTank(pFuelTankId, pCurrentFuelQuantity, pMaxFuelQuantiy)

            'checking
            If _numProducts >= _maxProductSize Then
                _maxProductSize += mARRAY_SIZE__INCREATMENT_DEFAULT
                ReDim Preserve mProducts(_maxProductSize - 1)
            End If
            If _numFuelTanks >= _maxFuelTankSize Then
                _maxFuelTankSize += mARRAY_SIZE__INCREATMENT_DEFAULT
                ReDim Preserve mFuelTanks(_maxFuelTankSize - 1)
            End If

            'adding into array
            Try
                _ithProducts(_numProducts) = theProduct
                _ithFuelTanks(_numFuelTanks) = theFuelTank
            Catch ex As Exception
                Throw New IndexOutOfRangeException
            End Try

            _numProducts += 1
            _numFuelTanks += 1

            'add transaction
            _createTransaction(
                pTrxId1,
                theTrxText,
                False)
            _createTransaction(
                pTrxId2,
                theTrxText,
                False)

            RaiseEvent FillingStation_ProductAdded(
                Me,
                New FillingStation_EventArgs_ProductAdded(theProduct)
                )

            RaiseEvent FillingStation_FuelTankAdded(
                Me,
                New FillingStation_EventArgs_FuelTankAdded(theFuelTank)
                )
        Else
            theProduct = New Product(
                pProductId,
                pProductName,
                pProductType,
                pUnitOfMeasure,
                pPricePerUnit,
                pLoyaltyDiscountPerUnit,
                pRewardDiscountPerUnit,
                pTaxRate
                )

            'checking
            If _numProducts >= _maxProductSize Then
                _maxProductSize += mARRAY_SIZE__INCREATMENT_DEFAULT
                ReDim Preserve mProducts(_maxProductSize - 1)
            End If

            'adding into array
            Try
                _ithProducts(_numProducts) = theProduct
            Catch ex As Exception
                Throw New IndexOutOfRangeException
            End Try

            _numProducts += 1

            'add transaction
            _createTransaction(pTrxId1, theTrxText, False)

            RaiseEvent FillingStation_ProductAdded(
                Me,
                New FillingStation_EventArgs_ProductAdded(theProduct)
                )
        End If

        Return theProduct

    End Function '_createProduct(pId,pProductName,pProductType,pUnitOfMeasure,pPricePerUnit,
    '             pLoyaltyDiscountPerUnit,pRewardDiscountPerUnit,pTaxRate,pFuelTankId,pCurrentFuelQuantity,pMaxFuelQuantity)

    Private Sub _modifyProduct(
            ByVal pProductId As String,
            ByVal pTrxDate As Date,
            ByVal pTrxId As String,
            ByVal pProductName As String,
            ByVal pUnitOfMeasure As String,
            ByVal pPricePerUnit As Decimal,
            ByVal pLoyaltyDiscountPerUnit As Decimal,
            ByVal pRewardDiscountPerUnit As Decimal,
            ByVal pTaxRate As Decimal
            )

        'declares local variable
        Dim theLocationFound As Integer
        Dim theTrxText As String

        'assign value to Text
        If _findProduct(pProductId, theLocationFound).productType = ProductType.FUEL Then
            theTrxText = pTrxDate.ToString("yyyyMMdd") & "; " & pTrxDate.ToString("hhmm") & "; " _
                    & "Product; " & "Update; " & pTrxId & "; " & pProductId & "; " _
                    & pProductName & "; " & pTaxRate.ToString & "; " & pUnitOfMeasure _
                    & "; " & pPricePerUnit & "; " & pLoyaltyDiscountPerUnit.ToString & "; " _
                    & pRewardDiscountPerUnit.ToString
        ElseIf _findProduct(pProductId, theLocationFound).productType = ProductType.CARWASH Then
            theTrxText = pTrxDate.ToString("yyyyMMdd") & "; " & pTrxDate.ToString("hhmm") & "; " _
                    & "Product; " & "Update; " & pTrxId & "; " & pProductId & "; " _
                    & pProductName & "; " & pTaxRate.ToString & "; " & pUnitOfMeasure _
                    & "; " & pPricePerUnit & "; " & pLoyaltyDiscountPerUnit.ToString
        Else
            theTrxText = pTrxDate.ToString("yyyyMMdd") & "; " & pTrxDate.ToString("hhmm") & "; " _
                & "Product; " & "Update; " & pTrxId & "; " & pProductId & "; " _
                & pProductName & "; " & pTaxRate.ToString & "; " & pUnitOfMeasure _
                & "; " & pPricePerUnit
        End If

        'checks productId and transactionId
        If _findTrx(pTrxId, theLocationFound) IsNot Nothing Then
            MessageBox.Show("The transaction ID is existed, please enter another new one.")
            Exit Sub
        End If
        If _findProduct(pProductId, theLocationFound) Is Nothing Then
            _createTransaction(pTrxId, theTrxText, True) 'adds transaction
            MessageBox.Show("The product is not existed.")
            Exit Sub
        End If

        'modifing
        _findProduct(pProductId, theLocationFound).pricePerUnit = pPricePerUnit
        _findProduct(pProductId, theLocationFound).loyaltyDiscountPerUnit = pLoyaltyDiscountPerUnit
        _findProduct(pProductId, theLocationFound).rewardDiscountPerUnit = pRewardDiscountPerUnit
        _findProduct(pProductId, theLocationFound).taxRate = pTaxRate

        'adds transaction
        _createTransaction(pTrxId, theTrxText, False)

        RaiseEvent FillingStation_ProductModified(
            Me,
            New FillingStation_EventArgs_ProductModified(_findProduct(pProductId, theLocationFound)))

    End Sub '_modifyProduct(pPricePerUnit, pLoyaltyDiscountPerUnit, pRewardDiscountPerUnit, pTaxRate)

    Private Function _orderFuel(
            ByVal pProductId As String,
            ByVal pTrxDate As Date,
            ByVal pTrxId As String,
            ByVal pFuelAmount As Decimal) _
        As _
            Decimal

        'declares
        Dim theLocationFound As Integer
        Dim theProduct As Product = _findProduct(pProductId, theLocationFound)
        Dim theFuelTank As FuelTank = theProduct.fuelTank
        Dim theTrxText As String

        'assigns
        theTrxText = pTrxDate.ToString("yyyyMMdd") & "; " & pTrxDate.ToString("hhmm") & "; " _
                    & "Product; " & "Update; " & pTrxId & "; " & pProductId & "; " _
                    & theProduct.productName & "; " & theProduct.taxRate.ToString & "; " & theProduct.unitOfMeasure _
                    & "; " & theProduct.pricePerUnit.ToString & "; " & theProduct.loyaltyDiscountPerUnit.ToString & "; " _
                    & theProduct.rewardDiscountPerUnit.ToString & "; " & theProduct.fuelTank.currentFuelTank.ToString & "; " _
                    & theProduct.fuelTank.maxFuelTank.ToString & "; " & pFuelAmount.ToString

        'validates 
        If theProduct.productType <> ProductType.FUEL Then
            MessageBox.Show("The product is not fuel, you can not do replenishing fuel operation.")
            _createTransaction(
                pTrxId,
                theTrxText,
                True)
            Exit Function
        End If
        If theFuelTank.currentFuelTank + pFuelAmount <= theFuelTank.maxFuelTank Then

            theFuelTank.currentFuelTank = theFuelTank.addFuel(pFuelAmount)

            _createTransaction(pTrxId, theTrxText, False)

            RaiseEvent FillingStation_FuelTankRefilled(
            Me,
            New FillingStation_EventArgs_FuelTankRefilled(theProduct))

            Return CDec(theFuelTank.currentFuelTank)
        Else
            MessageBox.Show("The amount you want to replenish exceeds capacity of fule tank.")
            _createTransaction(
                pTrxId,
                theTrxText,
                True)
            Exit Function
        End If

    End Function '_orderFuel(pProductId,pTrxId,pFuelAmount)

    Private Function _createSale(
            ByVal pSaleId As String,
            ByVal pTrxDate As Date,
            ByVal pTrxId As String,
            ByRef pCustomerID As String,
            ByRef pProductID As String,
            ByVal pQuantityPurchased As Decimal
            ) _
        As _
            Sale

        'declares local variables
        Dim theLocationFound As Integer
        Dim theCustomer As LoyaltyCustomer
        Dim theProduct As Product
        Dim theSale As Sale
        Dim theNetPricePerUnit As Decimal
        Dim theSubTotal As Decimal
        Dim theTaxAmount As Decimal
        Dim theTotal As Decimal
        Dim theTrxText As String

        'checks product ID
        If pProductID = "" Then
            MessageBox.Show("You must choose a product.")
            theProduct = Nothing
            theTrxText = pTrxDate.ToString("yyyyMMdd") & "; " & pTrxDate.ToString("hhmm") _
            & "; " & "Sale; " & "Create; " & pTrxId & "; " & pSaleId & "; " _
            & pCustomerID & "; " & "" & "; " & pQuantityPurchased.ToString & "; " _
            & 0 & "; " & 0 & "; " & 0
            _createTransaction(
                pTrxId,
                theTrxText,
                True
                )
            Exit Function
        Else
            If _findProduct(pProductID, theLocationFound) Is Nothing Then
                MessageBox.Show("The product is not existed.")
                theTrxText = pTrxDate.ToString("yyyyMMdd") & "; " & pTrxDate.ToString("hhmm") _
            & "; " & "Sale; " & "Create; " & pTrxId & "; " & pSaleId & "; " _
            & pCustomerID & "; " & pProductID & "; " & pQuantityPurchased.ToString & "; " _
            & 0 & "; " & 0 & "; " & 0
                _createTransaction(
                    pTrxId,
                    theTrxText,
                    True
                    )
                Exit Function
            Else
                theProduct = _findProduct(pProductID, theLocationFound)
                theTrxText = pTrxDate.ToString("yyyyMMdd") & "; " & pTrxDate.ToString("hhmm") _
                    & "; " & "Sale; " & "Create; " & pTrxId & "; " & pSaleId & "; " _
                    & pCustomerID & "; " & pProductID & "; " & pQuantityPurchased.ToString & "; " _
                    & theProduct.pricePerUnit.ToString & "; " & theProduct.loyaltyDiscountPerUnit.ToString & "; " & theProduct.taxRate.ToString
            End If
        End If

        'checks customer ID
        If pCustomerID = "" Then
            theCustomer = Nothing
            theTrxText = pTrxDate.ToString("yyyyMMdd") & "; " & pTrxDate.ToString("hhmm") _
                & "; " & "Sale; " & "Create; " & pTrxId & "; " & pSaleId & "; " _
                & "" & "; " & pProductID & "; " & pQuantityPurchased.ToString & "; " _
                & theProduct.pricePerUnit.ToString & "; " & 0 & "; " & theProduct.taxRate.ToString
        Else
            If _findCustomer(pCustomerID, theLocationFound) Is Nothing Then
                MessageBox.Show("The customer is not existed.")
                theTrxText = pTrxDate.ToString("yyyyMMdd") & "; " & pTrxDate.ToString("hhmm") _
                    & "; " & "Sale; " & "Create; " & pTrxId & "; " & pSaleId & "; " _
                    & pCustomerID & "; " & pProductID & "; " & pQuantityPurchased.ToString & "; " _
                    & theProduct.pricePerUnit.ToString & "; " & theProduct.loyaltyDiscountPerUnit.ToString & "; " & theProduct.taxRate.ToString
                _createTransaction(
                pTrxId,
                theTrxText,
                True
                )
                Exit Function
            Else
                theCustomer = _findCustomer(pCustomerID, theLocationFound)
                theTrxText = pTrxDate.ToString("yyyyMMdd") & "; " & pTrxDate.ToString("hhmm") _
                    & "; " & "Sale; " & "Create; " & pTrxId & "; " & pSaleId & "; " _
                    & pCustomerID & "; " & pProductID & "; " & pQuantityPurchased.ToString & "; " _
                    & theProduct.pricePerUnit.ToString & "; " & theProduct.loyaltyDiscountPerUnit.ToString & "; " & theProduct.taxRate.ToString
            End If
        End If

        'check duplicate sale ID
        If _findSale(pSaleId, theLocationFound) IsNot Nothing Then
            MessageBox.Show("The sale ID is existed, please enter another new one.")
            _createTransaction(pTrxId, theTrxText, True)
        End If

        'create an new sale
        If theCustomer Is Nothing Then

            'calculates 
            theSubTotal = pQuantityPurchased * theProduct.pricePerUnit
            If theProduct.productType = ProductType.FUEL Then
                theTaxAmount = 0
            Else
                theTaxAmount = theSubTotal * theProduct.taxRate
            End If
            theTotal = theSubTotal + theTaxAmount

            'creates an new sale
            theSale = New Sale(
                pSaleId,
                theProduct,
                pTrxDate,
                pQuantityPurchased,
                theProduct.pricePerUnit,
                theProduct.taxRate,
                theSubTotal,
                theTaxAmount,
                theTotal
                )

            'checking
            If _numSales >= _maxSaleSize Then
                _maxSaleSize += mARRAY_SIZE__INCREATMENT_DEFAULT
                ReDim Preserve mSales(_maxSaleSize - 1)
            End If

            'adding into array
            Try
                _ithSales(_numSales) = theSale
            Catch ex As Exception
                Throw New IndexOutOfRangeException
            End Try
            _numSales += 1

            'adds transaction
            _createTransaction(pTrxId, theTrxText, False)
        Else

            'calculates depends on product type
            If theProduct.productType = ProductType.FUEL Then

                If theCustomer.accruedRewardGallon >= 100 Then

                    'apply reward discount
                    theNetPricePerUnit =
                        theProduct.pricePerUnit - theProduct.rewardDiscountPerUnit

                    'calculates
                    theSubTotal = pQuantityPurchased * theNetPricePerUnit
                    theTaxAmount = 0
                    theTotal = theSubTotal + theTaxAmount

                    'create an new sale
                    theSale = New Sale(
                        pSaleId,
                        theCustomer,
                        theProduct,
                        pTrxDate,
                        pQuantityPurchased,
                        theProduct.pricePerUnit,
                        theProduct.loyaltyDiscountPerUnit,
                        theProduct.taxRate,
                        theSubTotal,
                        theTaxAmount,
                        theTotal
                        )

                    'remove customer accrued gallon to zero
                    theCustomer.accruedRewardGallon = 0

                Else

                    'apply regular loyalty discount
                    theNetPricePerUnit =
                        theProduct.pricePerUnit - theProduct.loyaltyDiscountPerUnit

                    'calculates
                    theSubTotal = pQuantityPurchased * theProduct.pricePerUnit
                    theTaxAmount = 0
                    theTotal = theSubTotal + theTaxAmount

                    'create an new sale
                    theSale = New Sale(
                        pSaleId,
                        theCustomer,
                        theProduct,
                        pTrxDate,
                        pQuantityPurchased,
                        theProduct.pricePerUnit,
                        theProduct.loyaltyDiscountPerUnit,
                        theProduct.taxRate,
                        theSubTotal,
                        theTaxAmount,
                        theTotal
                        )

                    'add accrued gallon into loyalty customer account
                    If theProduct.productType = ProductType.FUEL Then
                        _findCustomer(theCustomer.id, theLocationFound).accruedRewardGallon =
                            theCustomer.accruedRewardGallon + pQuantityPurchased
                    End If

                End If 'The accrued gallons is larger than 100 gallon or not

            ElseIf theProduct.productType = ProductType.CARWASH Then

                'apply loyalty dicount
                theNetPricePerUnit =
                    theProduct.pricePerUnit - theProduct.loyaltyDiscountPerUnit

                'regular calculation
                theSubTotal = pQuantityPurchased * theProduct.pricePerUnit
                theTaxAmount = theSubTotal * theProduct.taxRate
                theTotal = theSubTotal + theTaxAmount

                'create an new sale
                theSale = New Sale(
                    pSaleId,
                    theCustomer,
                    theProduct,
                    pTrxDate,
                    pQuantityPurchased,
                    theProduct.pricePerUnit,
                    theProduct.loyaltyDiscountPerUnit,
                    theProduct.taxRate,
                    theSubTotal,
                    theTaxAmount,
                    theTotal
                    )
            Else

                'regular calculation
                theSubTotal = pQuantityPurchased * theProduct.pricePerUnit
                theTaxAmount = theSubTotal * theProduct.taxRate
                theTotal = theSubTotal + theTaxAmount

                'create an new sale
                theSale = New Sale(
                    pSaleId,
                    theCustomer,
                    theProduct,
                    pTrxDate,
                    pQuantityPurchased,
                    theProduct.pricePerUnit,
                    theProduct.loyaltyDiscountPerUnit,
                    theProduct.taxRate,
                    theSubTotal,
                    theTaxAmount,
                    theTotal
                    )

            End If 'different product type

            'checking
            If _numSales >= _maxSaleSize Then
                _maxSaleSize += mARRAY_SIZE__INCREATMENT_DEFAULT
                ReDim Preserve mSales(_maxSaleSize - 1)
            End If

            'adding an new sale into array
            Try
                _ithSales(_numSales) = theSale
            Catch ex As Exception
                Throw New IndexOutOfRangeException
            End Try
            _numSales += 1

            'adds transaction
            _createTransaction(pTrxId, theTrxText, False)

        End If 'non-loyalty customer and loyalty customer

        'remove fuel amount
        If theProduct.productType = ProductType.FUEL Then
            _findFuelTank(theProduct.fuelTank.id, theLocationFound).currentFuelTank =
                        _findFuelTank(theProduct.fuelTank.id, theLocationFound).currentFuelTank - pQuantityPurchased
            If _findFuelTank(theProduct.fuelTank.id, theLocationFound).maxFuelTank * 0.02 < _findFuelTank(theProduct.fuelTank.id, theLocationFound).currentFuelTank AndAlso
            _findFuelTank(theProduct.fuelTank.id, theLocationFound).currentFuelTank < _findFuelTank(theProduct.fuelTank.id, theLocationFound).maxFuelTank * 0.05 Then
                MessageBox.Show("The current amount of fuel tank is below 5% of maximum capacity, please replenish.")
            End If
            If _findFuelTank(theProduct.fuelTank.id, theLocationFound).currentFuelTank < _findFuelTank(theProduct.fuelTank.id, theLocationFound).maxFuelTank * 0.02 Then
                MessageBox.Show("The current amount of fuel tank is below 2%, and filling station will replenish automatically.")
                _findFuelTank(theProduct.fuelTank.id, theLocationFound).currentFuelTank = _findFuelTank(theProduct.fuelTank.id, theLocationFound).maxFuelTank
            End If
        End If

        RaiseEvent FillingStation_SaleAdded(
            Me,
            New FillingStation_EventArgs_SaleAdded(theSale)
            )

        Return theSale

    End Function '_createSale(pId,pCustomer,pProduct,pSaleDate,pQuantityPurchased,
    '               pPricePerUnit,pDiscountPerUnit,pTaxRate,pSubTotalAmount,pTaxAmount,pTotalAmount)

    Private Function _createTransaction(
            ByVal pId As String,
            ByVal pTransactionText As String,
            ByVal pIsError As Boolean
            ) _
        As _
            Transaction

        'declares local variables
        Dim theLocationFound As Integer
        Dim theTransaction As Transaction

        'validates duplicate trx ID
        If _findTrx(pId, theLocationFound) IsNot Nothing Then
            MessageBox.Show("The Transaction ID is existed, please enter another new one.")
            Exit Function
        End If

        theTransaction = New Transaction(
            pId,
            pTransactionText,
            pIsError
            )

        'checking array size
        If _numTransactions >= _maxTrxSize Then
            _maxTrxSize += mARRAY_SIZE__INCREATMENT_DEFAULT
            ReDim Preserve mTransactions(_maxTrxSize - 1)
        End If

        'adding into array
        Try
            _ithTrxs(_numTransactions) = theTransaction
        Catch ex As Exception
            Throw New IndexOutOfRangeException
        End Try

        _numTransactions += 1

        RaiseEvent FillingStation_TransactionAdded(
            Me,
            New FillingStation_EventArgs_TransactionAdded(theTransaction)
            )

        Return theTransaction

    End Function '_createTransaction(pId, pTransactionText, pIsError)

    Private Function _calMemberSinceAge(
            ByVal pDate As Date
            ) _
        As _
            Integer

        Dim theAge As Integer
        If Now.Month > pDate.Month Then
            theAge = Now.Year - pDate.Year
        ElseIf Now.Month < pDate.Month Then
            theAge = Now.Year - pDate.Year - 1
        Else
            If Now.Day >= pDate.Day Then
                theAge = Now.Year - pDate.Year
            Else
                theAge = Now.Year - pDate.Year - 1
            End If
        End If

        Return theAge

    End Function '_calMemberSinceAge(pMemberSince)

    Private Function _toString() As String

        Dim tempString As String

        tempString = "( FILLING STATION: " _
            & "number of loyalty customers=" & _numCustomers _
            & ", number of products=" & _numProducts _
            & ", number of sales=" & _numSales _
            & ", number of transactions" & _numTransactions _
            & " )"

        Return tempString

    End Function '_toString()

#End Region 'Behavioral Methods

#Region "Event Procedures"
    '******************************************************************
    'Event Procedures
    '******************************************************************

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

    Public Event FillingStation_CustomerAdded(
            ByVal sender As Object,
            ByVal e As EventArgs)

    Public Event FillingStation_ProductAdded(
           ByVal sender As Object,
           ByVal e As EventArgs)

    Public Event FillingStation_SaleAdded(
           ByVal sender As Object,
           ByVal e As EventArgs)

    Public Event FillingStation_TransactionAdded(
           ByVal sender As Object,
           ByVal e As EventArgs)

    Public Event FillingStation_FuelTankAdded(
           ByVal sender As Object,
           ByVal e As EventArgs)

    Public Event FillingStation_ProductModified(
           ByVal sender As Object,
           ByVal e As EventArgs)

    Public Event FillingStation_FuelTankRefilled(
           ByVal sender As Object,
           ByVal e As EventArgs)

#End Region 'Events

End Class 'FillingStation
