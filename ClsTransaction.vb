'Copyright (c) 2009-2017 Dan Turk

#Region "Class / File Comment Header block"
'Program:                proj03-FillingStation-Fang-Anmin
'File:                   ClsTransaction.vb
'Author:                 Anmin Fang
'Description:            This is a class to create a transaction object, 
'                        which includes transaction information.
'Date:                   2017 Oct 5
'                          - Added attributes, get/set methods, constructors, toString methods.
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

Public Class Transaction

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    'No Attributes are currently defined.

    '********** Module-level constants

    '********** Module-level variables
    Private mId As String
    Private mTransactionText As String
    Private mIsError As Boolean

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
           ByVal pTransactionText As String,
           ByVal pIsError As Boolean
           )

        MyBase.New

        _id = pId
        _transactionText = pTransactionText
        _isError = pIsError

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

    Public ReadOnly Property id As String
        Get
            Return _id
        End Get
    End Property

    Public Property transactionText As String
        Get
            Return _transactionText
        End Get
        Set(ByVal pValue As String)
            _transactionText = pValue
        End Set
    End Property

    Public Property isError As Boolean
        Get
            Return _isError
        End Get
        Set(ByVal pValue As Boolean)
            _isError = pValue
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

    Private Property _transactionText As String
        Get
            Return mTransactionText
        End Get
        Set(ByVal pValue As String)
            mTransactionText = pValue
        End Set
    End Property

    Private Property _isError As Boolean
        Get
            Return mIsError
        End Get
        Set(ByVal pValue As Boolean)
            mIsError = pValue
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

        Dim tempString As String

        tempString = "( Transaction: " _
            & "id=" & _id _
            & ", transaction text=" & _transactionText _
            & ", is error=" & _isError _
            & " )"

        Return tempString

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

End Class 'Transaction
