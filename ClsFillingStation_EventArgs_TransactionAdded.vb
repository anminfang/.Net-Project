'Copyright (c) 2009-2017 Dan Turk

#Region "Class / File Comment Header block"
'Program:                proj03-FillingStation-Fang-Anmin
'File:                   ClsFillingStation_EventArgs_TransactionAdded.vb
'Author:                 Anmin Fang
'Description:            This is a class to inherit eventargs.
'Date:                   2017 Nov 2
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

Public Class FillingStation_EventArgs_TransactionAdded
    Inherits EventArgs

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    '********** Module-level constants

    '********** Module-level variables

    Private mTheTransaction As Transaction

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
           ByVal pTheTransaction As Transaction
           )
        MyBase.New

        _theTransaction = pTheTransaction

    End Sub 'New(pTheTransaction)

    '********** Copy constructor(s)
    '             - one parameter, an object of the same class

#End Region 'Constructors

#Region "Get/Set Methods"
    '******************************************************************
    'Get/Set Methods
    '******************************************************************

    '********** Public Get/Set Methods
    '             - call private get/set methods to implement

    Public ReadOnly Property theTransaction As Transaction
        Get
            Return _theTransaction
        End Get
    End Property

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _theTransaction As Transaction
        Get
            Return mTheTransaction
        End Get
        Set(ByVal pValue As Transaction)
            mTheTransaction = pValue
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

    Public Overrides Function ToString() As String

        Return _toString()

    End Function 'ToString()

    '********** Private Non-Shared Behavioral Methods

    Private Function _toString() As String

        Dim tempStr As String

        tempStr = "( TRANSACTION  EVENT_ARGS TRANSACTION CREATED: " _
            & "Transaction=" & _theTransaction.ToString _
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

End Class 'FillingStation_EventArgs_transactionAdded
