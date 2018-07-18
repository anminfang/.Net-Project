'Copyright (c) 2009-2017 Dan Turk

#Region "Class / File Comment Header block"
'Program:                proj03-FillingStation-Fang-Anmin
'File:                   ClsFuelTank.vb
'Author:                 Anmin Fang
'Description:            This is a class to create a fuel tank object.
'Date:                   2017 Oct 5
'                          - Added attributes, get/set methods, constructors.
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

Public Class FuelTank

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    '********** Module-level constants

    '********** Module-level variables

    Private mId As String
    Private mProduct As Product
    Private mCurrentFuelQuantity As Decimal
    Private mMaxFuelQuantity As Decimal

#End Region 'Attributes

#Region "Constructors"
    '******************************************************************
    'Constructors
    '******************************************************************

    'These are all public.

    '********** Default constructor
    '             - no parameters

    Public Sub New()

    End Sub

    '********** Special constructor(s)
    '             - typically constructors have parameters 
    '               that are used to initialize attributes

    Public Sub New(
           ByVal pId As String,
           ByVal pCurrentFuelTank As Decimal,
           ByVal pMaxFuelTank As Decimal,
           ByVal pProduct As Product
           )

        MyBase.New

        _id = pId
        _currentFuelQuantity = pCurrentFuelTank
        _maxFuelQuantity = pMaxFuelTank
        _product = pProduct

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

    Public Property id As String
        Get
            Return _id()
        End Get
        Set(ByVal pValue As String)
            _id = pValue
        End Set
    End Property

    Public ReadOnly Property product As Product
        Get
            Return _product
        End Get
    End Property

    Public Property currentFuelTank As Decimal
        Get
            Return _currentFuelQuantity
        End Get
        Set(ByVal pValue As Decimal)
            _currentFuelQuantity = pValue
        End Set
    End Property

    Public Property maxFuelTank As Decimal
        Get
            Return _maxFuelQuantity
        End Get
        Set(ByVal pValue As Decimal)
            _maxFuelQuantity = pValue
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

    Private Property _product() As Product
        Get
            Return mProduct
        End Get
        Set(ByVal pValue As Product)
            mProduct = pValue
        End Set
    End Property

    Private Property _currentFuelQuantity() As Decimal
        Get
            Return mCurrentFuelQuantity
        End Get
        Set(ByVal pValue As Decimal)
            mCurrentFuelQuantity = pValue
        End Set
    End Property

    Private Property _maxFuelQuantity() As Decimal
        Get
            Return mMaxFuelQuantity
        End Get
        Set(ByVal pValue As Decimal)
            mMaxFuelQuantity = pValue
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

    Public Function addFuel(
           ByVal pValue As Decimal
           ) _
       As _
           Decimal

        Return _addFuel(pValue)

    End Function

    Public Function removeFuel(
           ByVal pValue As Decimal
           ) _
       As _
           Decimal

        Return _removeFuel(pValue)

    End Function

    Public Overrides Function ToString() As String

        Return _toString()

    End Function 'ToString()

    '********** Private Non-Shared Behavioral Methods

    Private Function _addFuel(
            ByVal pValue As Decimal
            ) _
        As _
            Decimal

        currentFuelTank += pValue

        Return currentFuelTank

    End Function

    Private Function _removeFuel(
            ByVal pValue As Decimal
            ) _
        As _
            Decimal

        currentFuelTank -= pValue

        Return currentFuelTank

    End Function

    Private Function _toString() As String

        Dim tempStr As String

        tempStr = "( FUEL TANK: " _
            & "id=" & _id _
            & ", current fuel quantity=" & _currentFuelQuantity.ToString _
            & ", max fuel quantity=" & _maxFuelQuantity.ToString _
            & " )"

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

End Class 'FuelTank
