'Copyright (c) 2009-2017 Dan Turk

#Region "Class / File Comment Header block"
'Program:                proj03-FillingStation-Fang-Anmin
'File:                   ClsLoyaltyCustomer.vb
'Author:                 Anmin Fang
'Description:            This is a class to create a loyalty customer object.
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

Public Class LoyaltyCustomer

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    '********** Module-level constants

    '********** Module-level variables

    Private mId As String
    Private mName As String
    Private mPhone As String
    Private mSecurityPin As String
    Private mMemberSince As Date
    Private mMembershipAge As Integer
    Private mAccruedRewardGallon As Decimal

#End Region 'Attributes

#Region "Constructors"
    '******************************************************************
    'Constructors
    '******************************************************************

    'No Constructors are currently defined.
    'These are all public.

    '********** Default constructor
    '             - no parameters
    Public Sub New()

        MyBase.New

    End Sub 'New()

    '********** Special constructor(s)
    '             - typically constructors have parameters 
    '               that are used to initialize attributes

    Public Sub New(
            ByVal pId As String,
            ByVal pName As String,
            ByVal pPhone As String,
            ByVal pSecurityPin As String,
            ByVal pMemberSince As Date,
            ByVal pMembershipAge As Integer,
            ByVal pAccruedRewardGallon As Decimal
            )

        MyBase.New

        _id = pId
        _name = pName
        _phone = pPhone
        _securityPin = pSecurityPin
        _memberSince = pMemberSince
        _membershipAge = pMembershipAge
        _accruedRewardGallon = pAccruedRewardGallon

    End Sub 'New()

    '********** Copy constructor(s)
    '             - one parameter, an object of the same class

#End Region 'Constructors

#Region "Get/Set Methods"
    '******************************************************************
    'Get/Set Methods
    '******************************************************************

    'No Get/Set Methods are currently defined.

    '********** Public Get/Set Methods
    '             - call private get/set methods to implement

    Public Property id As String
        Get
            Return _id
        End Get
        Set(ByVal pValue As String)
            _id = pValue
        End Set
    End Property

    Public Property name As String
        Get
            Return _name
        End Get
        Set(ByVal pValue As String)
            _name = pValue
        End Set
    End Property

    Public ReadOnly Property phone As String
        Get
            Return _phone
        End Get
    End Property

    Public Property securityPin As String
        Get
            Return _securityPin
        End Get
        Set(ByVal pValue As String)
            _securityPin = pValue
        End Set
    End Property

    Public Property membershipSince As Date
        Get
            Return _memberSince
        End Get
        Set(ByVal pValue As Date)
            _memberSince = pValue
        End Set
    End Property

    Public Property membershipAge As Integer
        Get
            Return _membershipAge
        End Get
        Set(ByVal pValue As Integer)
            _membershipAge = pValue
        End Set
    End Property

    Public Property accruedRewardGallon As Decimal
        Get
            Return _accruedRewardGallon
        End Get
        Set(ByVal pValue As Decimal)
            _accruedRewardGallon = pValue
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

    Private Property _name() As String
        Get
            Return mName
        End Get
        Set(ByVal pValue As String)
            mName = pValue
        End Set
    End Property

    Private Property _phone() As String
        Get
            Return mPhone
        End Get
        Set(ByVal pValue As String)
            mPhone = pValue
        End Set
    End Property

    Private Property _securityPin() As String
        Get
            Return mSecurityPin
        End Get
        Set(ByVal pValue As String)
            mSecurityPin = pValue
        End Set
    End Property

    Private Property _memberSince() As Date
        Get
            Return mMemberSince
        End Get
        Set(ByVal pValue As Date)
            mMemberSince = pValue
        End Set
    End Property

    Private Property _membershipAge() As Integer
        Get
            Return mMembershipAge
        End Get
        Set(ByVal pValue As Integer)
            mMembershipAge = pValue
        End Set
    End Property

    Private Property _accruedRewardGallon() As Decimal
        Get
            Return mAccruedRewardGallon
        End Get
        Set(ByVal pValue As Decimal)
            mAccruedRewardGallon = pValue
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

        Dim tempString As String

        tempString = "( LOYALTY CUSTOMER: " _
            & "id=" & _id _
            & ", name=" & _name _
            & ", phone=" & _phone _
            & ", security pin=" & _securityPin _
            & ", member since=" & _memberSince _
            & ", membership age=" & _membershipAge _
            & ", accrued reward gallon=" & _accruedRewardGallon _
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

End Class 'LoyaltyCustomer
