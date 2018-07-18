'Copyright (c) 2009-2017 Dan Turk

#Region "Class / File Comment Header block"
'Program:            Filling Station, v 02
'File:               FrmMain.vb
'Author:             Dan Turk
'Description:        This is the GUI for the Filling Station program.
'                      It allows for display of summary / metric and
'                      transaction information, as well as creating, 
'                      viewing, and editing Loyalty Customers, Products,
'                      Storage Tanks, and Sales, just as might be done
'                      in a system supporting a real-world Filling
'                      Station.
'Date:               2017 Sep 11,13,18,19
'                      - Created by Dan Turk
'                    2017 Sep 20
'                      - Modified _initializeUserInterface() code
'                        to name Proj02.
'                    2017 Sep 27
'                      - Modified GUI fields to match updates in 
'                        Class Diagram and correspondingly revised
'                        the tab order of GUI fields.
'                      - Added ToolTips everywhere.
'                    2017 Sep 28
'                      - Finished adding ToolTips in many places.
'                      - Pruned code to make this the starting
'                        point for Proj02.
'                    2017 Oct 5
'                      - Created methods to act when button clicked
'                      - Finished adding buttons click events method
'                    2017 Nov 2
'                      - Created methods for customer, product, sale,
'                        transaction, fueltank added event And product
'                        modified Event
'                    2017 Nov 5
'                      - Overrided gui control, such as selectedIndexChanged,
'                        textChanged, keyPress
'                    2017 Nov 8
'                      - Crash test, add try/catch and exceptions
'                    2017 Nov 26
'                      - Revised _btnCreateSale method to adapt cobCustomer is blank
'                    2017 Nov 27
'                      - Revised sale tab that compare customer with pin 
'                      - Added toString of calculation total sale and total tax amount
'                        for each product type in _saleAdded method
'                      - Added metrics methods toString in _saleAdded method
'                      - Revised _modifyProduct method(key in textbox and reaction)
'                    2017 Nov 28
'                      - Added lowest and highest amount sale toString in _saleAdded method
'                      - Added redundancy ID validation in createCustomer, createProduct,
'                        createSale methods
'                      - Added transaction ID labels and textboxes in customer, product, sale tabs
'                    2017 Nov 29
'                      - Added replenish fuel validation in btnProcessSale method
'                    2017 Dec 1
'                      - Added _fuelTankRefilled event procedure
'                      - Revised createCustomer, createProduct validation
'                    2017 Dec 2
'                      - Added customer tab listbox additems function and display content to textbox
'                      - Added product tab listbox additems functon and display content to textbox
'                    2017 Dec 3
'                      - Added button open file function
'                    2017 Dec 4
'                      - Revised button process sale, deleted pricePerUnit, taxRate, subTotal, taxAmount,
'                        totalAmount parameters
'                    2017 Dec 7
'                      - Replaced ithCustomer,Product,Sale,FuelTank,Trx with iterator(1441)
'                      - Revised find methods in selectedIndexChanged controls
'                    2017 Dec 8
'                      - Revised combobox CustomerID in tab Sale, to clear textbox AccruedGallon every time
'                        when re-select customer ID
'Tier:               User Interface
'Exceptions:         Generic exceptions
'Exception-Handling: try/catch
'Events:             None defined.
'Event-Handling:     Only a few normal GUI events are handled.
#End Region 'Class / File Comment Header block

#Region "Option / Imports"
Option Explicit On      'Must declare variables before using them
Option Strict On        'Must perform explicit data type conversions
Imports System.Globalization
Imports System.IO
#End Region 'Option / Imports

Public Class FrmMain

#Region "Attributes"
    '******************************************************************
    'Attributes + Module-level Constants+Variables
    '******************************************************************

    '********** Module-level constants

    '********** Module-level variables

    Private WithEvents mFillingStation As FillingStation

#End Region 'Attributes

#Region "Constructors"
    '******************************************************************
    'Constructors
    '******************************************************************

    'No Constructors are currently defined.
    'These are all public.

    '********** Default constructor
    '             - no parameters

    '********** Special constructor(s)
    '             - typically constructors have parameters 
    '               that are used to initialize attributes

    '********** Copy constructor(s)
    '             - one parameter, an object of the same class

#End Region 'Constructors

#Region "Get/Set Methods"
    '******************************************************************
    'Get/Set Methods
    '******************************************************************


    '********** Public Get/Set Methods
    '             - call private get/set methods to implement

    '********** Private Get/Set Methods
    '             - access attributes, begin name with underscore (_)

    Private Property _fillingStation As FillingStation
        Get
            Return mFillingStation
        End Get
        Set(ByVal pValue As FillingStation)
            mFillingStation = pValue
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

    '********** Private Non-Shared Behavioral Methods

    Private Sub _initializeBusinessLogic()

        _fillingStation = New FillingStation

    End Sub '_initializeBusinessLogic()

    Private Sub _initializeUserInterface()

        'Form

        '########## Me.Text = _windowLabel & " - " & _theFillingStation.name
        Me.Text = "Filling Station (v.Proj04) - Dan Turk"
        'Me.AcceptButton = ####
        Me.CancelButton = btnExit

        'set numeric control
        Me.nudTaxRateGrpCreateModifyTabProductTbcMain.Minimum = 0
        Me.nudTaxRateGrpCreateModifyTabProductTbcMain.Increment = 0.0001D

        Me.nudPricePerUnitGrpCreateModifyTabProductTbcMain.Minimum = 0
        Me.nudPricePerUnitGrpCreateModifyTabProductTbcMain.Increment = 1

        Me.nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Minimum = 0
        Me.nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Increment = 0.001D

        Me.nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Minimum = 0
        Me.nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Increment = 0.001D

        Me.nudMaxCapacityGrpFuelTankGrpCreateModifyTabProductTbcMain.Minimum = 0
        Me.nudMaxCapacityGrpFuelTankGrpCreateModifyTabProductTbcMain.Maximum = 5000D
        Me.nudMaxCapacityGrpFuelTankGrpCreateModifyTabProductTbcMain.Increment = 1D

        Me.nudCurrentVolumeGrpFuelTankGrpCreateModifyTabProductTbcMain.Minimum = 0
        Me.nudCurrentVolumeGrpFuelTankGrpCreateModifyTabProductTbcMain.Maximum = 5000D
        Me.nudCurrentVolumeGrpFuelTankGrpCreateModifyTabProductTbcMain.Increment = 1D

        'set sale tab date time picker to be today or future 
        dtpDateTabSalesTbcMain.MinDate = Date.Today
        dtpDateTabSalesTbcMain.MaxDate = Date.Today

        Dim btnCreateGrpCreateTabLoyaltyCustomerTbcMain As New Button
        Dim btnCreateGrpCreateModifyTabProductTbcMain As New Button

        tbcMain.SelectedTab = tabSummaryTbcMain
        'tbcMain.SelectedTab = tabLoyaltyCustomerTbcMain
        'tbcMain.SelectedTab = tabProductTbcMain
        'tbcMain.SelectedTab = tabSalesTbcMain
        'tbcMain.SelectedTab = tabFilesTbcMain


        'txtTrxLog.Text &=
        '    vbCrLf
        '"#### Show Transaction Log Here ####" _
        '& vbCrLf & vbCrLf
        Me.tipMain.SetToolTip(
            txtTrxLog,
            "Displays a chronological listing " _
            & "of all the transactions " _
            & "that have been carried out."
            )

        Me.tipMain.SetToolTip(
            btnExit,
            "Exit from and stop the running of this program."
            )

        'Main Tab Control

        'Main Tab Control - Summary Tab

        tabSummaryTbcMain.ToolTipText =
            "The Summary tab shows a summary of the information in the " _
            & "FillingStation system.  You can view more details by " _
            & "clicking on an ID in the list boxes, And can scroll " _
            & "through the TrxLog to see a listing of all the " _
            & "transactions that have been recorded."

        'lstLoyaltyCustomerTabSummaryTbcMain.Items.Add("Customer")
        'lstLoyaltyCustomerTabSummaryTbcMain.Items.Add("ID")
        'lstLoyaltyCustomerTabSummaryTbcMain.Items.Add("Here")
        lstLoyaltyCustomerTabSummaryTbcMain.SelectedIndex =
            Math.Min(0, lstLoyaltyCustomerTabSummaryTbcMain.Items.Count - 1)
        tipMain.SetToolTip(
            lstLoyaltyCustomerTabSummaryTbcMain,
            "Select an Id to see more information in the " _
            & "ToString information box to the right."
            )
        lblLoyaltyCustomerCountTabSummaryTbcMain.Text = "0"
        tipMain.SetToolTip(
            lblLoyaltyCustomerCountTabSummaryTbcMain,
            "This label shows the number of items recorded in the " _
            & "system and displayed in the above list."
            )

        'lstProductTabSummaryTbcMain.Items.Add("Product")
        'lstProductTabSummaryTbcMain.Items.Add("ID")
        'lstProductTabSummaryTbcMain.Items.Add("Here")
        lstProductTabSummaryTbcMain.SelectedIndex =
            Math.Min(0, lstProductTabSummaryTbcMain.Items.Count - 1)
        tipMain.SetToolTip(
            lstProductTabSummaryTbcMain,
            "Select an Id to see more information in the " _
            & "ToString information box to the right."
            )
        lblProductCountTabSummaryTbcMain.Text = "0"
        tipMain.SetToolTip(
            lblProductCountTabSummaryTbcMain,
            "This label shows the number of items recorded in the " _
            & "system and displayed in the above list."
            )

        'lstFuelTankTabSummaryTbcMain.Items.Add("FuelTank")
        'lstFuelTankTabSummaryTbcMain.Items.Add("ID")
        'lstFuelTankTabSummaryTbcMain.Items.Add("Here")
        lstFuelTankTabSummaryTbcMain.SelectedIndex =
            Math.Min(0, lstFuelTankTabSummaryTbcMain.Items.Count - 1)
        tipMain.SetToolTip(
            lstFuelTankTabSummaryTbcMain,
            "Select an Id to see more information in the " _
            & "ToString information box to the right."
            )
        lblFuelTankCountTabSummaryTbcMain.Text = "0"
        tipMain.SetToolTip(
            lblFuelTankCountTabSummaryTbcMain,
            "This label shows the number of items recorded in the " _
            & "system and displayed in the above list."
            )

        'lstSaleTabSummaryTbcMain.Items.Add("Sale")
        'lstSaleTabSummaryTbcMain.Items.Add("ID")
        'lstSaleTabSummaryTbcMain.Items.Add("Here")
        lstSaleTabSummaryTbcMain.SelectedIndex =
            Math.Min(0, lstSaleTabSummaryTbcMain.Items.Count - 1)
        tipMain.SetToolTip(
            lstSaleTabSummaryTbcMain,
            "Select an Id to see more information in the " _
            & "ToString information box to the right."
            )
        lblSaleCountTabSummaryTbcMain.Text = "0"
        tipMain.SetToolTip(
            lblSaleCountTabSummaryTbcMain,
            "This label shows the number of items recorded in the " _
            & "system and displayed in the above list."
            )

        'lstTrxTabSummaryTbcMain.Items.Add("Trx")
        'lstTrxTabSummaryTbcMain.Items.Add("ID")
        'lstTrxTabSummaryTbcMain.Items.Add("Here")
        lstTrxTabSummaryTbcMain.SelectedIndex =
            Math.Min(0, lstTrxTabSummaryTbcMain.Items.Count - 1)
        tipMain.SetToolTip(
            lstTrxTabSummaryTbcMain,
            "Select an Id to see more information in the " _
            & "ToString information box to the right."
            )
        lblTrxCountTabSummaryTbcMain.Text = "0"
        tipMain.SetToolTip(
            lblTrxCountTabSummaryTbcMain,
            "This label shows the number of items recorded in the " _
            & "system and displayed in the above list."
            )

        'txtInfoTabSummaryTbcMain.Text =
        '    "#### Show ToString Information Here ####"
        txtInfoTabSummaryTbcMain.ScrollToCaret()
        txtInfoTabSummaryTbcMain.AppendText(Text)
        tipMain.SetToolTip(
            txtInfoTabSummaryTbcMain,
            "The ToString information box displays information " _
            & "about the Id selected in one of the lists to the left."
            )

        txtMetricsGrpMetricsTabSummaryTbcMain.Text =
            "#### Show Metrics (KPIs, Key Performance Indicators) here ####"
        txtMetricsGrpMetricsTabSummaryTbcMain.ScrollToCaret()
        tipMain.SetToolTip(
            txtMetricsGrpMetricsTabSummaryTbcMain,
            "The Metrics information box displays information " _
            & "that has been calculated from data in the system."
            )

        'Main Tab Control - Loyalty Customer Tab

        'cboLoyaltyCustomerIDTabLoyaltyCustomerTbcMain.Items.Add("Customer")
        'cboLoyaltyCustomerIDTabLoyaltyCustomerTbcMain.Items.Add("ID")
        'cboLoyaltyCustomerIDTabLoyaltyCustomerTbcMain.Items.Add("Here")
        cboLoyaltyCustomerIDTabLoyaltyCustomerTbcMain.SelectedIndex =
            Math.Min(0, cboLoyaltyCustomerIDTabLoyaltyCustomerTbcMain.Items.Count - 1)
        tipMain.SetToolTip(
            cboLoyaltyCustomerIDTabLoyaltyCustomerTbcMain,
            "Enter a unique new Id, or " _
            & "select an Id to see more information in the " _
            & "ToString information box to the right."
            )

        txtNameGrpCreateTabLoyaltyCustomerTbcMain.Text = ""
        txtPhoneGrpCreateTabLoyaltyCustomerTbcMain.Text = ""
        txtPINGrpCreateTabLoyaltyCustomerTbcMain.Text = ""
        txtPINGrpCreateTabLoyaltyCustomerTbcMain.UseSystemPasswordChar = False
        txtPINGrpCreateTabLoyaltyCustomerTbcMain.PasswordChar = " "c     'Cast to a "Character" type
        tipMain.SetToolTip(
            txtPINGrpCreateTabLoyaltyCustomerTbcMain,
            "Enter the Personal Identification Number (PIN) " _
            & "that the customer will use to authenticate themself " _
            & "when purchasing products."
            )
        txtConfirmPINGrpCreateTabLoyaltyCustomerTbcMain.Text = ""
        txtConfirmPINGrpCreateTabLoyaltyCustomerTbcMain.UseSystemPasswordChar = False
        txtConfirmPINGrpCreateTabLoyaltyCustomerTbcMain.PasswordChar = " "c     'Cast to a "Character" type
        tipMain.SetToolTip(
            txtConfirmPINGrpCreateTabLoyaltyCustomerTbcMain,
            "Re-enter (confirm) the Personal Identification Number (PIN) " _
            & "that the customer will use to authenticate themself " _
            & "when purchasing products."
            )
        dtpMemberSinceGrpCreateTabLoyaltyCustomerTbcMain.Value = Now
        tipMain.SetToolTip(
            dtpMemberSinceGrpCreateTabLoyaltyCustomerTbcMain,
            "Choose/enter the date the customer became a Loyalty Customer."
            )
        lblMemberAgeGrpCreateTabLoyaltyCustomerTbcMain.Text = "0"
        tipMain.SetToolTip(
            lblMemberAgeGrpCreateTabLoyaltyCustomerTbcMain,
            "This displays the number of years the person has been a " _
            & "Loyalty Customer."
            )
        lblAccruedRewardGallonsGrpCreateTabLoyaltyCustomerTbcMain.Text = "0"
        tipMain.SetToolTip(
            lblAccruedRewardGallonsGrpCreateTabLoyaltyCustomerTbcMain,
            "This displays the number of Reward Gallons currently accrued " _
            & " by the Loyalty Customer."
            )

        'lstCustTrxTabLoyaltyCustomerTbcMain.Items.Add("Customer's")
        'lstCustTrxTabLoyaltyCustomerTbcMain.Items.Add("Trx ID")
        'lstCustTrxTabLoyaltyCustomerTbcMain.Items.Add("Here")
        lstCustTrxTabLoyaltyCustomerTbcMain.SelectedIndex =
            Math.Min(0, lstCustTrxTabLoyaltyCustomerTbcMain.Items.Count - 1)
        tipMain.SetToolTip(
            lstCustTrxTabLoyaltyCustomerTbcMain,
            "Select an Id to see more information in the " _
            & "ToString information box to the right."
            )
        lblCustTrxCountTabLoyaltyCustomerTbcMain.Text = "0"
        tipMain.SetToolTip(
            lblCustTrxCountTabLoyaltyCustomerTbcMain,
            "This label shows the number of items recorded in the " _
            & "system and displayed in the above list."
            )

        'txtInfoTabLoyaltyCustomerTbcMain.Text =
        '    "#### Show ToString Information Here ####"
        tipMain.SetToolTip(
            txtInfoTabLoyaltyCustomerTbcMain,
            "The ToString information box displays information " _
            & "about the Id selected in one of the lists to the left."
            )

        'Main Tab Control - Product Tab

        'cboProductIDTabProductTbcMain.Items.Add("Product")
        'cboProductIDTabProductTbcMain.Items.Add("ID")
        'cboProductIDTabProductTbcMain.Items.Add("Here")
        cboProductIDTabProductTbcMain.SelectedIndex =
            Math.Min(0, cboProductIDTabProductTbcMain.Items.Count - 1)
        tipMain.SetToolTip(
            cboProductIDTabProductTbcMain,
            "Enter a unique new Id, or " _
            & "select an Id to see more information in the " _
            & "ToString information box to the right."
            )

        txtNameGrpCreateModifyTabProductTbcMain.Text = ""
        nudTaxRateGrpCreateModifyTabProductTbcMain.Value = 0
        tipMain.SetToolTip(
            nudTaxRateGrpCreateModifyTabProductTbcMain,
            "Enter the tax rate in decimal, not percent.  " _
            & "  Ex: Enter 0.0125 for 1.25%"
            )
        radFuelGrpTypeGrpCreateModifyTabProductTbcMain.Checked = True
        txtUnitGrpCreateModifyTabProductTbcMain.Text = ""
        tipMain.SetToolTip(
            txtUnitGrpCreateModifyTabProductTbcMain,
            "The unit of measure should be 'Gallons' or 'Each'."
            )
        'nudPricePerUnitGrpCreateModifyTabProductTbcMain.Value = 0
        'nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value = 0
        'nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value = 0

        txtFuelTankIDGrpFuelTankGrpCreateModifyTabProductTbcMain.Text = ""
        tipMain.SetToolTip(
            txtFuelTankIDGrpFuelTankGrpCreateModifyTabProductTbcMain,
            "Enter a unique new Id."
            )

        'nudMaxCapacityGrpFuelTankGrpCreateModifyTabProductTbcMain.Minimum = 0
        'nudMaxCapacityGrpFuelTankGrpCreateModifyTabProductTbcMain.Maximum = 0
        'nudMaxCapacityGrpFuelTankGrpCreateModifyTabProductTbcMain.Maximum =
        '    FuelTank.MAX_FUEL_TANK_CAPACITY_DEFAULT
        nudMaxCapacityGrpFuelTankGrpCreateModifyTabProductTbcMain.Value =
            nudMaxCapacityGrpFuelTankGrpCreateModifyTabProductTbcMain.Maximum
        tipMain.SetToolTip(
            nudMaxCapacityGrpFuelTankGrpCreateModifyTabProductTbcMain,
            "Enter or select the fuel tank's maximum capacity."
            )
        'nudCurrentVolumeGrpFuelTankGrpCreateModifyTabProductTbcMain.Value = 0

        txtQtyGrpOrderFuelTabProductTbcMain.Text = ""
        tipMain.SetToolTip(
            txtQtyGrpOrderFuelTabProductTbcMain,
            "Enter the quantity of fuel to order for " _
            & "replenishing the tank. This may not be more than " _
            & "the amount of space remaining in the tank."
            )

        'lstProdTrxTabProductTbcMain.Items.Add("Product's")
        'lstProdTrxTabProductTbcMain.Items.Add("Trx ID")
        'lstProdTrxTabProductTbcMain.Items.Add("Here")
        lstProdTrxTabProductTbcMain.SelectedIndex =
            Math.Min(0, lstProdTrxTabProductTbcMain.Items.Count - 1)
        tipMain.SetToolTip(
            lstProdTrxTabProductTbcMain,
            "Select an Id to see more information in the " _
            & "ToString information box to the right."
            )
        lblProdTrxCountTabProductTbcMain.Text = "0"
        tipMain.SetToolTip(
            lblProdTrxCountTabProductTbcMain,
            "This label shows the number of items recorded in the " _
            & "system and displayed in the above list."
            )

        'txtInfoTabProductTbcMain.Text =
        '    "#### Show ToString Information Here ####"
        tipMain.SetToolTip(
            txtInfoTabProductTbcMain,
            "The ToString information box displays information " _
            & "about the Id selected in one of the lists to the left."
            )

        'Main Tab Control - Sales Tab

        txtSaleIdTabSalesTbcMain.Text = ""
        tipMain.SetToolTip(
            txtSaleIdTabSalesTbcMain,
            "Enter a unique new Id."
            )
        'dtpDateTabSalesTbcMain.Value = Now
        tipMain.SetToolTip(
            dtpDateTabSalesTbcMain,
            "Enter or select the date of the sale."
            )
        'radFuelGrpTypeTabSalesTbcMain.Checked = True
        tipMain.SetToolTip(
            grpTypeTabSalesTbcMain,
            "Select the type of sale."
            )
        'tipMain.SetToolTip(
        '    radFuelGrpTypeTabSalesTbcMain,
        '    "Fuel sale."
        '    )
        'tipMain.SetToolTip(
        '    radCarWashGrpTypeTabSalesTbcMain,
        '    "Car Wash sale."
        '    )
        'tipMain.SetToolTip(
        '    radMiscGrpTypeTabSalesTbcMain,
        '    "Misc sale."
        '    )

        'cboCustomerIdTabSalesTbcMain.Items.Add("Customer")
        'cboCustomerIdTabSalesTbcMain.Items.Add("ID")
        'cboCustomerIdTabSalesTbcMain.Items.Add("Here")
        cboCustomerIdTabSalesTbcMain.SelectedIndex =
            Math.Min(0, cboCustomerIdTabSalesTbcMain.Items.Count - 1)
        tipMain.SetToolTip(
            cboCustomerIdTabSalesTbcMain,
            "Select an Id, more information will be shown in the " _
            & "ToString information box to the right."
            )

        txtPINTabSalesTbcMain.Text = ""
        txtPINTabSalesTbcMain.UseSystemPasswordChar = False
        txtPINTabSalesTbcMain.PasswordChar = " "c     'Cast to a "Character" type
        tipMain.SetToolTip(
            txtPINTabSalesTbcMain,
            "Enter a Personal Identification Number.  " _
            & "For security reasons, " _
            & "nothing will be displayed as you type this."
            )

        txtAccruedRewardGallonsTabSalesTbcMain.Text = "0"
        lblRewardMsgTabSalesTbcMain.Font = New Font(FontFamily.GenericSansSerif, 8, FontStyle.Italic)
        lblRewardMsgTabSalesTbcMain.ForeColor = Color.Red
        lblRewardMsgTabSalesTbcMain.TextAlign = ContentAlignment.BottomLeft
        lblRewardMsgTabSalesTbcMain.Text = "#### Show Rewards Msg Here ####"

        'cboProductIdTabSalesTbcMain.Items.Add("Product")
        'cboProductIdTabSalesTbcMain.Items.Add("ID")
        'cboProductIdTabSalesTbcMain.Items.Add("Here")
        cboProductIdTabSalesTbcMain.SelectedIndex =
            Math.Min(0, cboProductIdTabSalesTbcMain.Items.Count - 1)
        tipMain.SetToolTip(
            cboCustomerIdTabSalesTbcMain,
            "Select an Id, more information will be shown in the " _
            & "ToString information box to the right."
            )

        txtQtyTabSalesTbcMain.Text = ""
        tipMain.SetToolTip(
            txtQtyTabSalesTbcMain,
            "Enter the quantity of the product you wish to purchase."
            )
        txtPricePerUnitTabSalesTbcMain.Text = ""
        txtDiscountPerUnitTabSalesTbcMain.Text = ""
        txtNetPricePerUnitTabSalesTbcMain.Text = ""
        txtSubtotalPriceTabSalesTbcMain.Text = ""
        txtTaxTabSalesTbcMain.Text = ""
        txtTotalPriceTabSalesTbcMain.Text = ""

        'txtInfoTabSalesTbcMain.Text =
        '    "#### Show ToString Information Here ####"

        'Main Tab Control - Files Tab

        txtFilenameTabFilesTbcMain.Text = ""

    End Sub '_initializeUserInterface()

    Private Sub _processTestData()

        Dim cust01 As LoyaltyCustomer
        Dim cust02 As LoyaltyCustomer
        Dim cust03 As LoyaltyCustomer
        Dim prod01 As Product
        Dim prod02 As Product
        Dim prod03 As Product
        Dim prod04 As Product
        Dim sale01 As Sale
        Dim sale02 As Sale
        Dim sale03 As Sale
        Dim sale04 As Sale
        Dim sale05 As Sale
        Dim sale06 As Sale
        Dim sale07 As Sale
        Dim sale08 As Sale

        cust01 = _fillingStation.createLoyaltyCustomer(
            "CUST-01",
            Date.ParseExact(
            "20151101", "yyyyMMdd", CultureInfo.InvariantCulture),
            "TRX-01TEST",
            "Dan Brown",
            "9708251234",
            "123",
            "123",
            Date.Today,
            CInt(0),
            0
            )
        cust02 = _fillingStation.createLoyaltyCustomer(
            "CUST-02",
            Date.ParseExact(
            "20170508", "yyyyMMdd", CultureInfo.InvariantCulture),
            "TRX-02TEST",
            "Sam Smith",
            "9708255678",
            "456",
            "456",
            New Date(2015, 11, 26),
            CInt(2),
            0
            )
        cust03 = _fillingStation.createLoyaltyCustomer(
            "CUST-03",
            Date.ParseExact(
            "20001101", "yyyyMMdd", CultureInfo.InvariantCulture),
            "TRX-03TEST",
            "Tiger Woodz",
            "9708259876",
            "789",
             "789",
            New Date(2013, 11, 26),
            CInt(4),
            0
            )
        prod01 = _fillingStation.createProduct(
            "PROD-01",
            Date.ParseExact(
            Trim("20171101"), "yyyyMMdd", CultureInfo.InvariantCulture),
            "TRX-04TEST",
            "TRX-05TEST",
            "#82",
            ProductType.FUEL,
            "Gallon",
            CDec(2.35),
            CDec(0.03),
            CDec(0.1),
            CDec(0.09),
            "#1TANK",
            CDec(3600),
            CDec(5000)
            )
        prod02 = _fillingStation.createProduct(
            "PROD-02",
            Date.ParseExact(
            Trim("20171001"), "yyyyMMdd", CultureInfo.InvariantCulture),
            "TRX-06TEST",
            "",
            "Carwash",
            ProductType.CARWASH,
            "Each",
            CDec(5.0),
            1,
            0,
            CDec(0.09),
            "",
            0,
            0
            )
        prod03 = _fillingStation.createProduct(
            "PROD-03",
            Date.ParseExact(
            Trim("20170901"), "yyyyMMdd", CultureInfo.InvariantCulture),
            "TRX-07TEST",
            "",
            "Clean & Clear Misc",
            ProductType.MISC,
            "Each",
            CDec(2),
            0,
            0,
            CDec(0.09),
            "",
            0,
            0
            )
        prod04 = _fillingStation.createProduct(
            "PROD-04",
            Date.ParseExact(
            Trim("20170801"), "yyyyMMdd", CultureInfo.InvariantCulture),
            "TRX-08TEST",
            "TRX-09TEST",
            "#85",
            ProductType.FUEL,
            "Gallon",
            CDec(2.65),
            CDec(0.03),
            CDec(0.1),
            CDec(0.09),
            "#2TANK",
            CDec(2500),
            CDec(5000)
            )
        sale01 = _fillingStation.createSale(
            “SALE-01”,
            Date.Today,
            "TRX-10TEST",
            "CUST-01",
            "PROD-01",
            CInt(100)
            )
        sale02 = _fillingStation.createSale(
            “SALE-02”,
            Date.Today,
             "TRX-11TEST",
            "CUST-02",
            "PROD-02",
            CInt(1)
            )
        sale03 = _fillingStation.createSale(
            "SALE-03",
            Date.Today,
            "TRX-12TEST",
            "CUST-01",
            "PROD-01",
            25
            )
        sale04 = _fillingStation.createSale(
            "SALE-04",
            Date.Today,
            "TRX-13TEST",
            "CUST-01",
            "PROD-04",
            80
            )
        sale05 = _fillingStation.createSale(
            "SALE-05",
            Date.Today,
            "TRX-14TEST",
            "CUST-02",
            "PROD-02",
            1
            )
        sale06 = _fillingStation.createSale(
            "SALE-06",
            Date.Today,
            "TRX-15TEST",
            "CUST-03",
            "PROD-03",
            4
            )
        sale07 = _fillingStation.createSale(
            "SALE-07",
            Date.Today,
            "TRX-16TEST",
            "CUST-04",
            "PROD-01",
            10
            )
        sale08 = _fillingStation.createSale(
            "SALE-08",
            Date.Today,
            "TRX-17TEST",
            "CUST-02",
            "PROD-05",
            10)

        txtTrxLog.Text &=
            vbCrLf & "Starting _processTestData()"
        txtTrxLog.Text &=
            vbCrLf & "- Filling Station CREATED: " & _fillingStation.ToString
        txtTrxLog.Text &=
            vbCrLf & "- CUSTOMER CREATED: " & cust01.ToString
        txtTrxLog.Text &=
            vbCrLf & "- CUSTOMER CREATED: " & cust02.ToString
        txtTrxLog.Text &=
            vbCrLf & "- CUSTOMER CREATED: " & cust03.ToString
        txtTrxLog.Text &=
            vbCrLf & "- PRODUCT CREATED: " & prod01.ToString
        txtTrxLog.Text &=
            vbCrLf & "- PRODUCT CREATED: " & prod02.ToString
        txtTrxLog.Text &=
            vbCrLf & "- PRODUCT CREATED: " & prod03.ToString
        txtTrxLog.Text &=
            vbCrLf & "- PRODUCT CREATED: " & prod04.ToString
        txtTrxLog.Text &=
           vbCrLf & "- SALE CREATED: " & sale01.ToString
        txtTrxLog.Text &=
           vbCrLf & "- SALE CREATED: " & sale02.ToString
        txtTrxLog.Text &=
          vbCrLf & "- SALE CREATED: " & sale03.ToString
        txtTrxLog.Text &=
           vbCrLf & "- SALE CREATED: " & sale04.ToString
        txtTrxLog.Text &=
           vbCrLf & "- SALE CREATED: " & sale05.ToString
        txtTrxLog.Text &=
           vbCrLf & "- SALE CREATED: " & sale06.ToString

    End Sub '_processTestData()

    Private Function _datatimepickerConvertToDateCreateCustomer() As Date 'convert string to date datatype

        Dim theDate As Date

        theDate = New Date(
            dtpMemberSinceGrpCreateTabLoyaltyCustomerTbcMain.Value.Year,
            dtpMemberSinceGrpCreateTabLoyaltyCustomerTbcMain.Value.Month,
            dtpMemberSinceGrpCreateTabLoyaltyCustomerTbcMain.Value.Day
            )

        Return theDate

    End Function '_datatimepickerConvertToDateCreateCustomer

    Private Sub _clearTrxLog()

        txtTrxLog.Clear()

    End Sub '_clearTrxLog()

    Private Sub _resetAll()

        Controls.Clear()
        InitializeComponent()

    End Sub '_resetAll()

    Private Sub _displayFillingStation()

        txtTrxLog.Text &=
           vbCrLf & "Proj04--FillingStation-Fang-Anmin" _
           & vbCrLf _
           & mFillingStation.ToString

    End Sub

#End Region 'Behavioral Methods

#Region "Event Procedures"
    '******************************************************************
    'Event Procedures
    '******************************************************************

    'These are all private.

    '********** User-Interface Event Procedures
    '             - Initiated explicitly by user

    '******************VALIDATION**********************
    Private Sub _radioButtonProduct_CheckedChanged(
            sender As Object,
            e As EventArgs) _
        Handles _
            radCarWashGrpTypeGrpCreateModifyTabProductTbcMain.CheckedChanged,
            radMiscGrpTypeGrpCreateModifyTabProductTbcMain.CheckedChanged,
            radFuelGrpTypeGrpCreateModifyTabProductTbcMain.CheckedChanged

        If radFuelGrpTypeGrpCreateModifyTabProductTbcMain.Checked Then
            grpFuelTankGrpCreateModifyTabProductTbcMain.Visible = True
            lblLoyaltyDiscountPerUnitTabProductTbcMain.Visible = True
            nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Visible = True
            lblRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Visible = True
            nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Visible = True
        End If
        If radCarWashGrpTypeGrpCreateModifyTabProductTbcMain.Checked Then
            grpFuelTankGrpCreateModifyTabProductTbcMain.Visible = False
            lblLoyaltyDiscountPerUnitTabProductTbcMain.Visible = True
            nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Visible = True
            lblRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Visible = False
            nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Visible = False
        End If
        If radMiscGrpTypeGrpCreateModifyTabProductTbcMain.Checked Then
            grpFuelTankGrpCreateModifyTabProductTbcMain.Visible = False
            lblLoyaltyDiscountPerUnitTabProductTbcMain.Visible = False
            nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Visible = False
            lblRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Visible = False
            nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Visible = False
        End If

    End Sub '_radioButtonProduct_CheckedChanged(sender,e)

    Private Sub _txtPhoneGrpCreateTabLoyaltyCustomerTbcMain_KeyPress(
            ByVal sender As Object,
            ByVal e As System.Windows.Forms.KeyPressEventArgs) _
    Handles _
            txtPhoneGrpCreateTabLoyaltyCustomerTbcMain.KeyPress

        If Asc(e.KeyChar) <> 13 AndAlso Asc(e.KeyChar) <> 8 AndAlso Not IsNumeric(e.KeyChar) Then
            MessageBox.Show("Please enter numbers only")
            txtPhoneGrpCreateTabLoyaltyCustomerTbcMain.SelectAll()
            txtPhoneGrpCreateTabLoyaltyCustomerTbcMain.Focus()
            e.Handled = True
        End If

        If txtPhoneGrpCreateTabLoyaltyCustomerTbcMain.Text.Length > 10 Then
            If e.KeyChar <> ControlChars.Back Then
                MessageBox.Show("Phone number can not be over 10 digits.")
                txtPhoneGrpCreateTabLoyaltyCustomerTbcMain.SelectAll()
                txtPhoneGrpCreateTabLoyaltyCustomerTbcMain.Focus()
                e.Handled = True
            End If

        End If

    End Sub '_txtPhoneGrpCreateTabLoyaltyCustomerTbcMain_KeyPress(sender,KeyPressEventArgs)

    Private Sub _dtpMemberSinceGrpCreateTabLoyaltyCustomerTbcMain_ValueChanged(
                sender As Object,
                e As EventArgs) _
         Handles _
                dtpMemberSinceGrpCreateTabLoyaltyCustomerTbcMain.ValueChanged

        With dtpMemberSinceGrpCreateTabLoyaltyCustomerTbcMain.Value
            Dim theMemberAge As Integer
            Dim theMemberSince As DateTime = New DateTime(Now.Year, .Month, .Day)
            If Now.Month > .Month Then
                theMemberAge = Now.Year - .Year
            ElseIf Now.Month < .Month Then
                theMemberAge = Now.Year - .Year - 1
            Else
                If Now.Day >= .Day Then
                    theMemberAge = Now.Year - .Year
                Else
                    theMemberAge = Now.Year - .Year - 1
                End If
            End If
            If theMemberAge < 0 Then theMemberAge = 0
            lblMemberAgeGrpCreateTabLoyaltyCustomerTbcMain.Text = CStr(theMemberAge)
        End With

    End Sub '_dtpMemberSinceGrpCreateTabLoyaltyCustomerTbcMain_ValueChanged(sender,e)

    Private Sub _txtUnitGrpCreateModifyTabProductTbcMain_KeyPress(   'limit only number in unit of measure textbox 
            ByVal sender As Object,
            ByVal e As System.Windows.Forms.KeyPressEventArgs) _
    Handles _
            txtUnitGrpCreateModifyTabProductTbcMain.KeyPress

        '97 - 122 = Ascii codes for simple letters
        '65 - 90  = Ascii codes for capital letters
        '32 = Ascii codes for space

        Select Case Asc(e.KeyChar)
            Case 8, 32, 65 To 90, 97 To 122
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("Please enter letters only. Ex: Gallon or Each")
                txtNameGrpCreateTabLoyaltyCustomerTbcMain.SelectAll()
                txtNameGrpCreateTabLoyaltyCustomerTbcMain.Focus()
                Exit Sub
        End Select

    End Sub '_txtUnitGrpCreateModifyTabProductTbcMain_KeyPress(sender,KeyPressEventArgs)

    Private Sub _txtQtyGrpOrderFuelTabProductTbcMain_KeyPress(   'limit number only in quantity order textbox
            ByVal sender As Object,
            ByVal e As KeyPressEventArgs
            ) _
        Handles _
            txtQtyGrpOrderFuelTabProductTbcMain.KeyPress

        Select Case Asc(e.KeyChar)
            Case 8, 46, 48 To 57
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("Please enter decimal numbers only. Ex: 123.456")
                txtQtyGrpOrderFuelTabProductTbcMain.SelectAll()
                txtQtyGrpOrderFuelTabProductTbcMain.Focus()
                Exit Sub
        End Select

    End Sub '_txtQtyGrpOrderFuelTabProductTbcMain_KeyPress(sender,KeyPressEventArgs)

    Private Sub _txtQtyTabSalesTbcMain_KeyPress(
            ByVal sender As Object,
            ByVal e As System.Windows.Forms.KeyPressEventArgs
            ) _
        Handles _
            txtQtyTabSalesTbcMain.KeyPress

        Select Case Asc(e.KeyChar)
            Case 8, 46, 48 To 57
                e.Handled = False
            Case Else
                e.Handled = True
                MessageBox.Show("Please enter decimal numbers only. Ex: 123.456")
                txtQtyTabSalesTbcMain.SelectAll()
                txtQtyTabSalesTbcMain.Focus()
                Exit Sub
        End Select

    End Sub '_txtQtyTabSalesTbcMain_KeyPress(sender,e)

    '-------------------DISPLAY---------------------
    Private Sub _lstLoyaltyCustomerTabSummaryTbcMain_SelectedIndexChanged(
            ByVal sender As Object,
            ByVal e As EventArgs
            ) _
        Handles _
            lstLoyaltyCustomerTabSummaryTbcMain.SelectedIndexChanged

        Dim theSelectedItem As String = CStr(
            lstLoyaltyCustomerTabSummaryTbcMain.SelectedItem)
        'Dim theArrayIndex As Integer = _fillingStation.findCustomer(
        '    theSelectedItem)
        'Dim theCustomer As LoyaltyCustomer = _fillingStation.ithCustomers(
        '    theArrayIndex)
        Dim theCustomer As LoyaltyCustomer

        For Each theCustomer In _fillingStation.iterateCustomer
            If theCustomer.id = theSelectedItem Then
                txtInfoTabSummaryTbcMain.Text &=
                    vbCrLf & "Customer ID selected: " & theCustomer.id _
                    & vbCrLf _
                    & "Customer=" & theCustomer.ToString
            End If
        Next

        'If theArrayIndex = -1 Then
        '    MessageBox.Show("Customer is not found.")
        '    Exit Sub
        'End If

    End Sub '_lstLoyaltyCustomerTabSummaryTbcMain_SelectedIndexChanged(sender,e)

    Private Sub _lstProductTabSummaryTbcMain_SelectedIndexChanged(
            ByVal sender As Object,
            ByVal e As EventArgs
            ) _
        Handles _
            lstProductTabSummaryTbcMain.SelectedIndexChanged

        Dim theSelectedItem As String = CStr(
            lstProductTabSummaryTbcMain.SelectedItem)
        'Dim theArrayIndex As Integer = _fillingStation.findProduct(
        '    theSelectedItem)
        'Dim theProduct As Product = _fillingStation.ithProducts(
        '    theArrayIndex)
        'If theArrayIndex = -1 Then
        '    MessageBox.Show("Product is not found.")
        '    Exit Sub
        'End If

        Dim theProduct As Product

        For Each theProduct In _fillingStation.iterateProduct
            If theProduct.id = theSelectedItem Then
                txtInfoTabSummaryTbcMain.Text &=
                    vbCrLf & "Product ID selected: " & theProduct.id _
                    & vbCrLf _
                    & "Product=" & theProduct.ToString
            End If
        Next

    End Sub '_lstProductTabSummaryTbcMain_SelectedIndexChanged(sender,e)

    Private Sub _lstFuelTankTabSummaryTbcMain_SelectedIndexChanged(
            ByVal sender As Object,
            ByVal e As EventArgs
            ) _
        Handles _
            lstFuelTankTabSummaryTbcMain.SelectedIndexChanged

        Dim theSelectedItem As String = CStr(
            lstFuelTankTabSummaryTbcMain.SelectedItem)
        'Dim theArrayIndex As Integer = _fillingStation.findFuelTank(
        '    theSelectedItem)
        'Dim theFuelTank As FuelTank = _fillingStation.ithFuelTanks(
        '    theArrayIndex)
        'If theArrayIndex = -1 Then
        '    MessageBox.Show("FuelTank is not found.")
        '    Exit Sub
        'End If

        Dim theFuelTank As FuelTank

        For Each theFuelTank In _fillingStation.iterateFuelTank
            If theFuelTank.id = theSelectedItem Then
                txtInfoTabSummaryTbcMain.Text &=
                    vbCrLf & "FuelTank ID selected: " & theFuelTank.id _
                    & vbCrLf _
                    & "FuelTank=" & theFuelTank.ToString
            End If
        Next

    End Sub '_lstFuelTankTabSummaryTbcMain_SelectedIndexChanged(sender,e)

    Private Sub _lstSaleTabSummaryTbcMain_SelectedIndexChanged(
            ByVal sender As Object,
            ByVal e As EventArgs
            ) _
        Handles _
            lstSaleTabSummaryTbcMain.SelectedIndexChanged

        Dim theSelectedItem As String = CStr(
            lstSaleTabSummaryTbcMain.SelectedItem)
        'Dim theArrayIndex As Integer = _fillingStation.findSale(
        '    theSelectedItem)
        'Dim theSale As Sale = _fillingStation.ithSales(
        '    theArrayIndex)
        'If theArrayIndex = -1 Then
        '    MessageBox.Show("Sale is not found.")
        '    Exit Sub
        'End If

        Dim theSale As Sale

        For Each theSale In _fillingStation.iterateSale
            If theSale.id = theSelectedItem Then
                txtInfoTabSummaryTbcMain.Text &=
                    vbCrLf & "Sale ID selected: " & theSale.id _
                    & vbCrLf _
                    & "Sale=" & theSale.ToString
            End If
        Next

    End Sub '_lstSaleTabSummaryTbcMain_SelectedIndexChanged(sender,e)

    Private Sub _lstTrxTabSummaryTbcMain_SelectedIndexChanged(
            ByVal sender As Object,
            ByVal e As EventArgs
            ) _
        Handles _
            lstTrxTabSummaryTbcMain.SelectedIndexChanged

        Dim theSelectedItem As String = CStr(
            lstTrxTabSummaryTbcMain.SelectedItem)
        'Dim theArrayIndex As Integer = _fillingStation.findTrx(
        '    theSelectedItem)
        'Dim theTrx As Transaction = _fillingStation.ithTrxs(
        '    theArrayIndex)
        'If theArrayIndex = -1 Then
        '    MessageBox.Show("Trx is not found.")
        '    Exit Sub
        'End If

        Dim theTrx As Transaction

        For Each theTrx In _fillingStation.iterateTrx
            If theTrx.id = theSelectedItem Then
                txtInfoTabSummaryTbcMain.Text &=
                    vbCrLf & "Transaction ID selected: " & theTrx.id _
                    & vbCrLf _
                    & "Transaction=" & theTrx.ToString
            End If
        Next

    End Sub '_lstTrxTabSummaryTbcMain_SelectedIndexChanged(sender,e)

    Private Sub _cboLoyaltyCustomerIDTabLoyaltyCustomerTbcMain_SelectedIndexChanged(
            sender As Object,
            e As EventArgs) _
        Handles _
            cboLoyaltyCustomerIDTabLoyaltyCustomerTbcMain.SelectedIndexChanged

        'clear listbox
        lstCustTrxTabLoyaltyCustomerTbcMain.Items.Clear()

        'refer to customer array
        Dim theSelectedItem As String = CStr(
            cboLoyaltyCustomerIDTabLoyaltyCustomerTbcMain.SelectedItem)
        'Dim theArrayIndex As Integer = _fillingStation.findCustomer(
        '    theSelectedItem)
        'Dim theCustomer As LoyaltyCustomer = _fillingStation.ithCustomers(
        '    theArrayIndex)

        'declares trx variables
        Dim theCustomer As LoyaltyCustomer
        Dim theTrx As Transaction
        'Dim theCoutner As Integer

        'iterate trx array
        For Each theTrx In _fillingStation.iterateTrx
            If theTrx.transactionText.Contains(theSelectedItem) Then
                lstCustTrxTabLoyaltyCustomerTbcMain.Items.Add(theTrx.id)
            End If
        Next

        'add items counter into textbox
        lblCustTrxCountTabLoyaltyCustomerTbcMain.Text = lstCustTrxTabLoyaltyCustomerTbcMain.Items.Count.ToString

        'display customer info
        For Each theCustomer In _fillingStation.iterateCustomer
            If theCustomer.id = theSelectedItem Then
                txtNameGrpCreateTabLoyaltyCustomerTbcMain.Text = theCustomer.name
                txtPhoneGrpCreateTabLoyaltyCustomerTbcMain.Text = theCustomer.phone
                dtpMemberSinceGrpCreateTabLoyaltyCustomerTbcMain.Value = theCustomer.membershipSince
                lblMemberAgeGrpCreateTabLoyaltyCustomerTbcMain.Text = theCustomer.membershipAge.ToString
                lblAccruedRewardGallonsGrpCreateTabLoyaltyCustomerTbcMain.Text = theCustomer.accruedRewardGallon.ToString
            End If
        Next

    End Sub '_cboLoyaltyCustomerIDTabLoyaltyCustomerTbcMain_SelectedIndexChanged(sender,e)

    Private Sub _lstCustTrxTabLoyaltyCustomerTbcMain_SelectedIndexChanged(
            sender As Object,
            e As EventArgs
            ) _
        Handles _
            lstCustTrxTabLoyaltyCustomerTbcMain.SelectedIndexChanged

        Dim theSelectedItem As String = CStr(
            lstCustTrxTabLoyaltyCustomerTbcMain.SelectedItem)
        Dim theTrx As Transaction

        'display in rightside textbox
        For Each theTrx In _fillingStation.iterateTrx
            If theTrx.transactionText.Contains(theSelectedItem) Then
                txtInfoTabLoyaltyCustomerTbcMain.Text &=
                            vbCrLf & theTrx.ToString
            End If
        Next

    End Sub '_lstCustTrxTabLoyaltyCustomerTbcMain_SelectedIndexChanged(sender,e)

    Private Sub _cboProductIDTabProductTbcMain_SelectedItem(
            ByVal sender As Object,
            ByVal e As EventArgs
            ) _
        Handles _
            cboProductIDTabProductTbcMain.SelectedIndexChanged

        'clear listbox
        lstProdTrxTabProductTbcMain.Items.Clear()

        'reset product tab
        radFuelGrpTypeGrpCreateModifyTabProductTbcMain.Checked = True

        'set items to disable
        txtNameGrpCreateModifyTabProductTbcMain.Enabled = False
        grpTypeGrpCreateModifyTabProductTbcMain.Enabled = False
        txtUnitGrpCreateModifyTabProductTbcMain.Enabled = False
        grpFuelTankGrpCreateModifyTabProductTbcMain.Enabled = False

        'assign product ID to get info from product array
        Dim theSelectedItem As String = CStr(
            cboProductIDTabProductTbcMain.SelectedItem)

        'declares trx variables
        Dim theTrx As Transaction
        Dim theProduct As Product
        'Dim theCoutner As Integer

        'iterate trx array
        For Each theTrx In _fillingStation.iterateTrx
            If theTrx.transactionText.Contains(theSelectedItem) Then
                lstProdTrxTabProductTbcMain.Items.Add(theTrx.id)
            End If
        Next

        'add items counter into textbox
        lblProdTrxCountTabProductTbcMain.Text = lstProdTrxTabProductTbcMain.Items.Count.ToString

        For Each theProduct In _fillingStation.iterateProduct
            If theProduct.id = theSelectedItem Then
                txtNameGrpCreateModifyTabProductTbcMain.Text = theProduct.productName
                nudTaxRateGrpCreateModifyTabProductTbcMain.Value = theProduct.taxRate
                If theProduct.productType = ProductType.FUEL Then
                    radFuelGrpTypeGrpCreateModifyTabProductTbcMain.Checked = True
                    txtFuelTankIDGrpFuelTankGrpCreateModifyTabProductTbcMain.Text = theProduct.fuelTank.id
                    nudMaxCapacityGrpFuelTankGrpCreateModifyTabProductTbcMain.Value = theProduct.fuelTank.maxFuelTank
                    nudCurrentVolumeGrpFuelTankGrpCreateModifyTabProductTbcMain.Value = theProduct.fuelTank.currentFuelTank
                ElseIf theProduct.productType = ProductType.CARWASH Then
                    radCarWashGrpTypeGrpCreateModifyTabProductTbcMain.Checked = True
                Else
                    radMiscGrpTypeGrpCreateModifyTabProductTbcMain.Checked = True
                End If
                txtUnitGrpCreateModifyTabProductTbcMain.Text = theProduct.unitOfMeasure
                nudPricePerUnitGrpCreateModifyTabProductTbcMain.Value = theProduct.pricePerUnit
                nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value = theProduct.loyaltyDiscountPerUnit
                nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value = theProduct.rewardDiscountPerUnit

                txtInfoTabProductTbcMain.Text = theProduct.ToString
            End If
        Next

    End Sub '_cboProductIDTabProductTbcMain_SelectedItem(sender,e)

    Private Sub _lstProdTrxTabProductTbcMain_SelectedIndexChanged(
            sender As Object,
            e As EventArgs
            ) _
        Handles _
            lstProdTrxTabProductTbcMain.SelectedIndexChanged

        Dim theSelectedItem As String = CStr(
            lstProdTrxTabProductTbcMain.SelectedItem)
        Dim theTrx As Transaction

        'display in rightside textbox
        For Each theTrx In _fillingStation.iterateTrx
            If theTrx.transactionText.Contains(theSelectedItem) Then
                txtInfoTabProductTbcMain.Text &=
                    vbCrLf & theTrx.ToString
            End If
        Next

    End Sub '_lstProdTrxTabProductTbcMain_SelectedIndexChanged(sender,e)

    Private Sub _cboProductIDTabProductTbcMain_TextChanged(
            ByVal sender As Object,
            ByVal e As EventArgs
            ) _
        Handles _
            cboProductIDTabProductTbcMain.TextChanged

        'reset product tab
        nudTaxRateGrpCreateModifyTabProductTbcMain.Value = 0
        nudPricePerUnitGrpCreateModifyTabProductTbcMain.Value = 0
        nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value = 0
        nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value = 0

        txtNameGrpCreateModifyTabProductTbcMain.Enabled = True
        grpTypeGrpCreateModifyTabProductTbcMain.Enabled = True
        txtUnitGrpCreateModifyTabProductTbcMain.Enabled = True
        grpFuelTankGrpCreateModifyTabProductTbcMain.Enabled = True
        grpFuelTankGrpCreateModifyTabProductTbcMain.Visible = True
        lblLoyaltyDiscountPerUnitTabProductTbcMain.Visible = True
        nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Visible = True
        lblRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Visible = True
        nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Visible = True

        txtNameGrpCreateModifyTabProductTbcMain.Text = ""
        txtUnitGrpCreateModifyTabProductTbcMain.Text = ""
        txtFuelTankIDGrpFuelTankGrpCreateModifyTabProductTbcMain.Text = ""
        nudMaxCapacityGrpFuelTankGrpCreateModifyTabProductTbcMain.Value = 0

    End Sub '_cboProductIDTabProductTbcMain_TextChanged(sender,e)

    Private Sub _txtPINTabSalesTbcMain_TextChanged(
            sender As Object,
            e As EventArgs) _
        Handles _
            txtPINTabSalesTbcMain.TextChanged

        Dim theCustomer As LoyaltyCustomer

        For Each theCustomer In _fillingStation.iterateCustomer
            If theCustomer.id = CStr(cboCustomerIdTabSalesTbcMain.SelectedItem) Then
                If theCustomer.securityPin = txtPINTabSalesTbcMain.Text.Trim Then
                    txtAccruedRewardGallonsTabSalesTbcMain.Text = theCustomer.accruedRewardGallon.ToString
                End If
            End If
        Next

    End Sub '_txtPINTabSalesTbcMain_TextChanged(sender,e)

    Private Sub _cboProductIdTabSalesTbcMain_SelectedIndexChanged(
            sender As Object,
            e As EventArgs) _
        Handles _
            cboProductIdTabSalesTbcMain.SelectedIndexChanged

        Dim theSelectedItem As String = CStr(
            cboProductIdTabSalesTbcMain.SelectedItem)
        Dim theProduct As Product

        'display price and discount in textbox
        For Each theProduct In _fillingStation.iterateProduct
            If theProduct.id.Equals(theSelectedItem) Then
                txtPricePerUnitTabSalesTbcMain.Text = theProduct.pricePerUnit.ToString
                If CDec(txtAccruedRewardGallonsTabSalesTbcMain.Text.Trim) >= 100 Then
                    txtDiscountPerUnitTabSalesTbcMain.Text = theProduct.rewardDiscountPerUnit.ToString
                Else
                    txtDiscountPerUnitTabSalesTbcMain.Text = theProduct.loyaltyDiscountPerUnit.ToString
                End If
                txtNetPricePerUnitTabSalesTbcMain.Text = (
                    theProduct.pricePerUnit - CDec(txtDiscountPerUnitTabSalesTbcMain.Text.Trim)).ToString
            End If
        Next

    End Sub '_cboProductIdTabSalesTbcMain_SelectedIndexChanged(sender,e)

    Private Sub _cboProductIdTabSalesTbcMain_TextChanged(
            sender As Object,
            e As EventArgs
            ) _
        Handles _
            cboProductIdTabSalesTbcMain.TextChanged

        txtAccruedRewardGallonsTabSalesTbcMain.Text = "0"
        txtPricePerUnitTabSalesTbcMain.Clear()
        txtDiscountPerUnitTabSalesTbcMain.Clear()
        txtNetPricePerUnitTabSalesTbcMain.Clear()

    End Sub '_cboProductIdTabSalesTbcMain_TextChanged(sender,e)

    Private Sub _txtQtyTabSalesTbcMain_TextChanged(
            sender As Object,
            e As EventArgs) _
        Handles _
            txtQtyTabSalesTbcMain.TextChanged


        Dim theLocationFound As Integer
        Dim theSubTotal As Decimal
        Dim theTotal As Decimal
        Dim theTaxAmount As Decimal
        Dim theProduct As Product

        'calculate sale bill
        For Each theProduct In _fillingStation.iterateProduct
            If theProduct.id.Equals(CStr(cboProductIdTabSalesTbcMain.SelectedItem)) Then
                If cboProductIdTabSalesTbcMain.Text = "" Then
                    txtQtyTabSalesTbcMain.Clear()
                    txtSubtotalPriceTabSalesTbcMain.Clear()
                    txtTaxTabSalesTbcMain.Clear()
                    txtTotalPriceTabSalesTbcMain.Clear()
                ElseIf _fillingStation.findProduct(
                    CStr(cboProductIdTabSalesTbcMain.SelectedItem)) Is Nothing Then
                    MessageBox.Show("The product is not existed, please select another item or enter another product ID.")
                    cboProductIdTabSalesTbcMain.SelectAll()
                    cboProductIdTabSalesTbcMain.Focus()
                    Exit Sub
                ElseIf txtQtyTabSalesTbcMain.Text.Trim = "" Then
                    theSubTotal = 0
                    theTaxAmount = 0
                    theTotal = 0
                ElseIf theProduct.productType = ProductType.FUEL Then
                    theSubTotal = CDec(txtQtyTabSalesTbcMain.Text.Trim) * CDec(txtNetPricePerUnitTabSalesTbcMain.Text.Trim)
                    theTaxAmount = 0
                    theTotal = theSubTotal + theTaxAmount
                Else
                    theSubTotal = CDec(txtQtyTabSalesTbcMain.Text.Trim) * CDec(txtNetPricePerUnitTabSalesTbcMain.Text.Trim)
                    theTaxAmount = CDec(txtQtyTabSalesTbcMain.Text.Trim) * theProduct.taxRate
                    theTotal = theSubTotal + theTaxAmount
                End If
            End If
        Next

        txtSubtotalPriceTabSalesTbcMain.Text = theSubTotal.ToString
        txtTaxTabSalesTbcMain.Text = theTaxAmount.ToString
        txtTotalPriceTabSalesTbcMain.Text = theTotal.ToString

    End Sub '_txtQtyTabSalesTbcMain_TextChanged(sender,e)

    Private Sub _cboCustomerIdTabSalesTbcMain_SelectedIndexChanged(
            sender As Object,
            e As EventArgs) _
        Handles _
            cboCustomerIdTabSalesTbcMain.SelectedIndexChanged

        txtAccruedRewardGallonsTabSalesTbcMain.Clear()

    End Sub '_cboCustomerIdTabSalesTbcMain_SelectedIndexChanged(sender,e)

    '**********************BUTTON CLICK*********************************
    Private Sub _btnExit_Click(
            sender As Object,
            e As EventArgs) _
        Handles _
            btnExit.Click

        Me.Close()

    End Sub '_btnExit_Click(sender, e)

    Private Sub _btnCreateCustomer_click(
            sender As Object,
            e As EventArgs) _
        Handles _
            btnCreateGrpCreateTabLoyaltyCustomerTbcMain.Click

        Dim theCustomer As LoyaltyCustomer

        If cboLoyaltyCustomerIDTabLoyaltyCustomerTbcMain.Text.Trim = "" Then
            MessageBox.Show("Please enter customer ID.")
            cboLoyaltyCustomerIDTabLoyaltyCustomerTbcMain.SelectAll()
            cboLoyaltyCustomerIDTabLoyaltyCustomerTbcMain.Focus()
            Exit Sub
        End If
        If txtTrxIdGrpCreateTabLoyaltyCustomerTbcMain.Text = "" Then
            MessageBox.Show("Please enter transaction ID.")
            txtTrxIdGrpCreateTabLoyaltyCustomerTbcMain.SelectAll()
            txtTrxIdGrpCreateTabLoyaltyCustomerTbcMain.Focus()
            Exit Sub
        End If
        If txtNameGrpCreateTabLoyaltyCustomerTbcMain.Text.Trim = "" Then
            MessageBox.Show("Space is not allowable, please enter customer name.")
            txtNameGrpCreateTabLoyaltyCustomerTbcMain.SelectAll()
            txtNameGrpCreateTabLoyaltyCustomerTbcMain.Focus()
            Exit Sub
        End If
        If txtPhoneGrpCreateTabLoyaltyCustomerTbcMain.Text.Trim = "" Then
            MessageBox.Show("Please enter phone number.")
            txtPhoneGrpCreateTabLoyaltyCustomerTbcMain.SelectAll()
            txtPhoneGrpCreateTabLoyaltyCustomerTbcMain.Focus()
            Exit Sub
        End If
        If txtPINGrpCreateTabLoyaltyCustomerTbcMain.Text.Trim = "" Then
            MessageBox.Show("Please enter pin number.")
            txtPINGrpCreateTabLoyaltyCustomerTbcMain.SelectAll()
            txtPINGrpCreateTabLoyaltyCustomerTbcMain.Focus()
            Exit Sub
        End If
        If txtConfirmPINGrpCreateTabLoyaltyCustomerTbcMain.Text.Trim = "" Then
            MessageBox.Show("Please enter the same pin number.")
            txtConfirmPINGrpCreateTabLoyaltyCustomerTbcMain.SelectAll()
            txtConfirmPINGrpCreateTabLoyaltyCustomerTbcMain.Focus()
            Exit Sub
        End If

        Try
            theCustomer = _fillingStation.createLoyaltyCustomer(
                cboLoyaltyCustomerIDTabLoyaltyCustomerTbcMain.Text.Trim,
                Date.Now,
                txtTrxIdGrpCreateTabLoyaltyCustomerTbcMain.Text.Trim,
                txtNameGrpCreateTabLoyaltyCustomerTbcMain.Text.Trim,
                txtPhoneGrpCreateTabLoyaltyCustomerTbcMain.Text.Trim,
                txtPINGrpCreateTabLoyaltyCustomerTbcMain.Text.Trim,
                txtConfirmPINGrpCreateTabLoyaltyCustomerTbcMain.Text.Trim,
                _datatimepickerConvertToDateCreateCustomer,
                Convert.ToInt16(lblMemberAgeGrpCreateTabLoyaltyCustomerTbcMain.Text),
                Convert.ToDecimal(lblAccruedRewardGallonsGrpCreateTabLoyaltyCustomerTbcMain.Text)
                )
        Catch ex As Exception
            MessageBox.Show("Error is happening. Please enter again.")
            cboLoyaltyCustomerIDTabLoyaltyCustomerTbcMain.SelectAll()
            cboLoyaltyCustomerIDTabLoyaltyCustomerTbcMain.Focus()
            Exit Sub
        End Try

        'reset
        cboLoyaltyCustomerIDTabLoyaltyCustomerTbcMain.Text = ""
        txtTrxIdGrpCreateTabLoyaltyCustomerTbcMain.Text = ""
        txtNameGrpCreateTabLoyaltyCustomerTbcMain.Text = ""
        txtPhoneGrpCreateTabLoyaltyCustomerTbcMain.Text = ""
        txtPINGrpCreateTabLoyaltyCustomerTbcMain.Text = ""
        txtConfirmPINGrpCreateTabLoyaltyCustomerTbcMain.Text = ""
        dtpMemberSinceGrpCreateTabLoyaltyCustomerTbcMain.Value = Today
        lblMemberAgeGrpCreateTabLoyaltyCustomerTbcMain.Text = "0"
        lblAccruedRewardGallonsGrpCreateTabLoyaltyCustomerTbcMain.Text = "0"

    End Sub '_btnCreateCusomter_Click(sender,e)

    Private Sub _btnCreateProduct_Click(
            sender As Object,
            e As EventArgs) _
        Handles _
            btnCreateGrpCreateModifyTabProductTbcMain.Click

        'declares variables
        Dim theProduct As Product

        'checks
        If cboProductIDTabProductTbcMain.Text.Trim = "" Then
            MessageBox.Show("Please select item or enter product ID.")
            cboProductIDTabProductTbcMain.SelectAll()
            cboProductIDTabProductTbcMain.Focus()
            Exit Sub
        End If
        If txtTrxId1GrpCreateModifyTabProductTbcMain.Text.Trim = "" Then
            MessageBox.Show("Please enter transaction ID.")
            txtTrxId1GrpCreateModifyTabProductTbcMain.SelectAll()
            txtTrxId1GrpCreateModifyTabProductTbcMain.Focus()
            Exit Sub
        End If
        If txtNameGrpCreateModifyTabProductTbcMain.Text.Trim = "" Then
            MessageBox.Show("Please enter product name.")
            txtNameGrpCreateModifyTabProductTbcMain.SelectAll()
            txtNameGrpCreateModifyTabProductTbcMain.Focus()
            Exit Sub
        End If
        If nudTaxRateGrpCreateModifyTabProductTbcMain.Value = 0 Then
            MessageBox.Show("Please set tax rate.")
            nudTaxRateGrpCreateModifyTabProductTbcMain.Select()
            nudTaxRateGrpCreateModifyTabProductTbcMain.Focus()
            Exit Sub
        End If
        If txtUnitGrpCreateModifyTabProductTbcMain.Text.Trim = "" Then
            MessageBox.Show("Please enter unit of measure.")
            txtUnitGrpCreateModifyTabProductTbcMain.SelectAll()
            txtUnitGrpCreateModifyTabProductTbcMain.Focus()
            Exit Sub
        End If
        If nudPricePerUnitGrpCreateModifyTabProductTbcMain.Value = 0 Then
            MessageBox.Show("Please set price per unit.")
            nudPricePerUnitGrpCreateModifyTabProductTbcMain.Select()
            nudPricePerUnitGrpCreateModifyTabProductTbcMain.Focus()
            Exit Sub
        End If
        If radFuelGrpTypeGrpCreateModifyTabProductTbcMain.Checked = True Then
            If nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value = 0 Then
                MessageBox.Show("Please set loyalty customer discount rate per unit.")
                nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Select()
                nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Focus()
                Exit Sub
            End If
            If nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value = 0 Then
                MessageBox.Show("Please set loyalty customer reward discount per unit.")
                nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Select()
                nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Focus()
                Exit Sub
            End If
            If txtTrxId2GrpFuelTankGrpFuelTankGrpCreateModifyTabProductTbcMain.Text.Trim = "" Then
                MessageBox.Show("Please type in transaction ID. Ex: TRX-00")
                txtTrxId2GrpFuelTankGrpFuelTankGrpCreateModifyTabProductTbcMain.SelectAll()
                txtTrxId2GrpFuelTankGrpFuelTankGrpCreateModifyTabProductTbcMain.Focus()
                Exit Sub
            End If
            If txtFuelTankIDGrpFuelTankGrpCreateModifyTabProductTbcMain.Text = "" Then
                MessageBox.Show("Please type in fuel tank ID. Ex: #1Tank")
                txtFuelTankIDGrpFuelTankGrpCreateModifyTabProductTbcMain.SelectAll()
                txtFuelTankIDGrpFuelTankGrpCreateModifyTabProductTbcMain.Focus()
                Exit Sub
            End If
            If nudMaxCapacityGrpFuelTankGrpCreateModifyTabProductTbcMain.Value = 0 Then
                MessageBox.Show("Please set maximum capacity for the fuel tank.")
                nudMaxCapacityGrpFuelTankGrpCreateModifyTabProductTbcMain.Focus()
                Exit Sub
            End If
        End If
        If radCarWashGrpTypeGrpCreateModifyTabProductTbcMain.Checked = True Then
            If nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value = 0 Then
                MessageBox.Show("Please set loyalty customer discount rate per unit.")
                nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Select()
                nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Focus()
                Exit Sub
            End If
        End If

        'creates an new product
        Try
            If radFuelGrpTypeGrpCreateModifyTabProductTbcMain.Checked Then

                theProduct = _fillingStation.createProduct(
                                    cboProductIDTabProductTbcMain.Text.Trim,
                                    Date.Now,
                                    txtTrxId1GrpCreateModifyTabProductTbcMain.Text.Trim,
                                    txtTrxId2GrpFuelTankGrpFuelTankGrpCreateModifyTabProductTbcMain.Text.Trim,
                                    txtNameGrpCreateModifyTabProductTbcMain.Text.Trim,
                                    ProductType.FUEL,
                                    txtUnitGrpCreateModifyTabProductTbcMain.Text.Trim,
                                    Convert.ToDecimal(nudPricePerUnitGrpCreateModifyTabProductTbcMain.Value),
                                    Convert.ToDecimal(nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value),
                                    Convert.ToDecimal(nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value),
                                    Convert.ToDecimal(nudTaxRateGrpCreateModifyTabProductTbcMain.Value),
                                    txtFuelTankIDGrpFuelTankGrpCreateModifyTabProductTbcMain.Text.Trim,
                                    Convert.ToDecimal(nudCurrentVolumeGrpFuelTankGrpCreateModifyTabProductTbcMain.Value),
                                    Convert.ToDecimal(nudMaxCapacityGrpFuelTankGrpCreateModifyTabProductTbcMain.Value)
                                    )
            ElseIf radCarWashGrpTypeGrpCreateModifyTabProductTbcMain.Checked Then

                theProduct = _fillingStation.createProduct(
                                    cboProductIDTabProductTbcMain.Text.Trim,
                                    Date.Now,
                                    txtTrxId1GrpCreateModifyTabProductTbcMain.Text.Trim,
                                    txtTrxId2GrpFuelTankGrpFuelTankGrpCreateModifyTabProductTbcMain.Text.Trim,
                                    txtNameGrpCreateModifyTabProductTbcMain.Text.Trim,
                                    ProductType.CARWASH,
                                    txtUnitGrpCreateModifyTabProductTbcMain.Text.Trim,
                                    Convert.ToDecimal(nudPricePerUnitGrpCreateModifyTabProductTbcMain.Value),
                                    Convert.ToDecimal(nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value),
                                    Convert.ToDecimal(nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value),
                                    Convert.ToDecimal(nudTaxRateGrpCreateModifyTabProductTbcMain.Value),
                                    txtFuelTankIDGrpFuelTankGrpCreateModifyTabProductTbcMain.Text.Trim,
                                    Convert.ToDecimal(nudCurrentVolumeGrpFuelTankGrpCreateModifyTabProductTbcMain.Value),
                                    Convert.ToDecimal(nudMaxCapacityGrpFuelTankGrpCreateModifyTabProductTbcMain.Value)
                                    )
            ElseIf radMiscGrpTypeGrpCreateModifyTabProductTbcMain.Checked Then

                theProduct = _fillingStation.createProduct(
                                    cboProductIDTabProductTbcMain.Text.Trim,
                                    Date.Now,
                                    txtTrxId1GrpCreateModifyTabProductTbcMain.Text.Trim,
                                    txtTrxId2GrpFuelTankGrpFuelTankGrpCreateModifyTabProductTbcMain.Text.Trim,
                                    txtNameGrpCreateModifyTabProductTbcMain.Text.Trim,
                                    ProductType.MISC,
                                    txtUnitGrpCreateModifyTabProductTbcMain.Text.Trim,
                                    Convert.ToDecimal(nudPricePerUnitGrpCreateModifyTabProductTbcMain.Value),
                                    Convert.ToDecimal(nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value),
                                    Convert.ToDecimal(nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value),
                                    Convert.ToDecimal(nudTaxRateGrpCreateModifyTabProductTbcMain.Value),
                                    txtFuelTankIDGrpFuelTankGrpCreateModifyTabProductTbcMain.Text.Trim,
                                    Convert.ToDecimal(nudCurrentVolumeGrpFuelTankGrpCreateModifyTabProductTbcMain.Value),
                                    Convert.ToDecimal(nudMaxCapacityGrpFuelTankGrpCreateModifyTabProductTbcMain.Value)
                                    )
            End If
        Catch ex As Exception
            MessageBox.Show("ERROR")
            cboProductIDTabProductTbcMain.SelectAll()
            cboProductIDTabProductTbcMain.Focus()
            Exit Sub
        End Try

        'reset
        cboProductIDTabProductTbcMain.Text = ""
        txtTrxId1GrpCreateModifyTabProductTbcMain.Text = ""
        txtNameGrpCreateModifyTabProductTbcMain.Text = ""
        txtUnitGrpCreateModifyTabProductTbcMain.Text = ""
        nudPricePerUnitGrpCreateModifyTabProductTbcMain.Value = 0
        nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value = 0
        nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value = 0
        nudTaxRateGrpCreateModifyTabProductTbcMain.Value = 0
        txtTrxId2GrpFuelTankGrpFuelTankGrpCreateModifyTabProductTbcMain.Text = ""
        txtFuelTankIDGrpFuelTankGrpCreateModifyTabProductTbcMain.Text = ""
        nudCurrentVolumeGrpFuelTankGrpCreateModifyTabProductTbcMain.Value = 0
        nudMaxCapacityGrpFuelTankGrpCreateModifyTabProductTbcMain.Value = 5000

    End Sub '_btnCreateProduct_Click(sender,e)

    Private Sub _btnProcessSaleTabSalesTbcMain_Click(
            ByVal sender As Object,
            ByVal e As EventArgs
            ) _
        Handles _
            btnProcessSaleTabSalesTbcMain.Click

        'initialize local variables
        Dim theLocationFound As Integer
        Dim theSelectedItemCustomer As String = CStr( 'customer validation
            cboCustomerIdTabSalesTbcMain.SelectedItem)
        Dim theCustomer As LoyaltyCustomer = _fillingStation.findCustomer(theSelectedItemCustomer)
        Dim theSelectedItemProduct As String = CStr( 'product validation
            cboProductIdTabSalesTbcMain.SelectedItem)
        Dim theProduct As Product = _fillingStation.findProduct(
            theSelectedItemProduct)
        If theProduct Is Nothing Then
            MessageBox.Show("The product is not existed, please select another item.")
            cboProductIdTabSalesTbcMain.SelectAll()
            cboProductIdTabSalesTbcMain.Focus()
            Exit Sub
        End If

        'declares textbox quantity check variable
        Dim theDecimalNum As Decimal

        'checks gui
        If txtSaleIdTabSalesTbcMain.Text.Trim = "" Then
            MessageBox.Show("Please enter sale ID.")
            txtSaleIdTabSalesTbcMain.SelectAll()
            txtSaleIdTabSalesTbcMain.Focus()
            Exit Sub
        End If
        If txtTrxIdTabSalesTbcMain.Text.Trim = "" Then
            MessageBox.Show("Please enter transaction ID.")
            txtTrxIdTabSalesTbcMain.SelectAll()
            txtTrxIdTabSalesTbcMain.Focus()
            Exit Sub
        End If
        If String.IsNullOrWhiteSpace(cboCustomerIdTabSalesTbcMain.Text) Then
        Else
            If theCustomer.securityPin.Equals(txtPINTabSalesTbcMain.Text.Trim) = False Then
                MessageBox.Show("The pin is not correct, please enter it again.")
                txtPINTabSalesTbcMain.SelectAll()
                txtPINTabSalesTbcMain.Focus()
                Exit Sub
            End If
        End If
        If (String.IsNullOrEmpty(cboProductIdTabSalesTbcMain.Text)) Then
            MessageBox.Show("Please select an item or type in.")
            cboProductIdTabSalesTbcMain.SelectAll()
            cboProductIdTabSalesTbcMain.Focus()
            Exit Sub
        End If
        If txtQtyTabSalesTbcMain.Text.Trim = "" Then
            MessageBox.Show("Please type in quantity here.")
            txtQtyTabSalesTbcMain.SelectAll()
            txtQtyTabSalesTbcMain.Focus()
            Exit Sub
        ElseIf Decimal.TryParse(txtQtyTabSalesTbcMain.Text, theDecimalNum) = False Then
            MessageBox.Show("The number you enter is not decimal number, please enter it again.")
            txtQtyTabSalesTbcMain.SelectAll()
            txtQtyTabSalesTbcMain.Focus()
            Exit Sub
        End If
        If cboProductIdTabSalesTbcMain.Text = "" Then
            MessageBox.Show("Please enter or select a Product ID.")
            cboProductIdTabSalesTbcMain.SelectAll()
            cboProductIdTabSalesTbcMain.Focus()
            Exit Sub
        End If
        'validates ID redundancy
        If _fillingStation.findSale(txtSaleIdTabSalesTbcMain.Text) IsNot Nothing Then
            MessageBox.Show("The sale ID entered has been already existed, please enter another ID.")
            txtSaleIdTabSalesTbcMain.SelectAll()
            txtSaleIdTabSalesTbcMain.Focus()
            Exit Sub
        End If

        If theProduct.productType = ProductType.FUEL Then
            If theProduct.fuelTank.currentFuelTank < CDec(txtQtyTabSalesTbcMain.Text.Trim) Then
                MessageBox.Show("The current amount of fuel tank is below the quantity you want to purchase, 
                                 please select another fuel tank or enter another quantity.")
                txtQtyTabSalesTbcMain.SelectAll()
                txtQtyTabSalesTbcMain.Focus()
                Exit Sub
            End If
        End If

        'processes create
        Try
            _fillingStation.createSale(
                txtSaleIdTabSalesTbcMain.Text.Trim,
                dtpDateTabSalesTbcMain.Value,
                txtTrxIdTabSalesTbcMain.Text.Trim,
                cboCustomerIdTabSalesTbcMain.Text.Trim,
                cboProductIdTabSalesTbcMain.Text.Trim,
                CDec(txtQtyTabSalesTbcMain.Text)
                )
        Catch ex As Exception
            MessageBox.Show("Error")
        End Try

        'reset
        txtSaleIdTabSalesTbcMain.Clear()
        txtTrxIdTabSalesTbcMain.Clear()
        cboCustomerIdTabSalesTbcMain.Text = ""
        txtPINTabSalesTbcMain.Clear()
        txtAccruedRewardGallonsTabSalesTbcMain.Clear()
        cboProductIdTabSalesTbcMain.Text = ""
        txtQtyTabSalesTbcMain.Clear()
        txtPricePerUnitTabSalesTbcMain.Clear()
        txtDiscountPerUnitTabSalesTbcMain.Clear()
        txtNetPricePerUnitTabSalesTbcMain.Clear()
        txtSubtotalPriceTabSalesTbcMain.Clear()
        txtTaxTabSalesTbcMain.Clear()
        txtTotalPriceTabSalesTbcMain.Clear()

    End Sub '_btnProcessSaleTabSalesTbcMain_Click(sender,e)

    Private Sub _btnModifyGrpCreateModifyTabProductTbcMain_Click( 'button Modify Product
            sender As Object,
            e As EventArgs
            ) _
        Handles _
            btnModifyGrpCreateModifyTabProductTbcMain.Click

        'validates
        If cboProductIDTabProductTbcMain.Text = "" Then
            MessageBox.Show("You should select or type in an item in product ID combobox.")
            cboProductIDTabProductTbcMain.SelectAll()
            cboProductIDTabProductTbcMain.Focus()
            Exit Sub
        End If
        If txtTrxId1GrpCreateModifyTabProductTbcMain.Text = "" Then
            MessageBox.Show("You should enter transaction ID.")
            txtTrxId1GrpCreateModifyTabProductTbcMain.SelectAll()
            txtTrxId1GrpCreateModifyTabProductTbcMain.Focus()
            Exit Sub
        End If
        If nudPricePerUnitGrpCreateModifyTabProductTbcMain.Value = 0 Then
            MessageBox.Show("Please set price per unit.")
            nudPricePerUnitGrpCreateModifyTabProductTbcMain.Select()
            nudPricePerUnitGrpCreateModifyTabProductTbcMain.Focus()
            Exit Sub
        End If

        'modifing
        Try
            _fillingStation.modityProduct(
                        cboProductIDTabProductTbcMain.Text,
                        Date.Now,
                        txtTrxId1GrpCreateModifyTabProductTbcMain.Text.Trim,
                        txtNameGrpCreateModifyTabProductTbcMain.Text.Trim,
                        txtUnitGrpCreateModifyTabProductTbcMain.Text.Trim,
                        Convert.ToDecimal(nudPricePerUnitGrpCreateModifyTabProductTbcMain.Value),
                        Convert.ToDecimal(nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value),
                        Convert.ToDecimal(nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value),
                        Convert.ToDecimal(nudTaxRateGrpCreateModifyTabProductTbcMain.Value)
                        )
        Catch ex As Exception
            MessageBox.Show("Error")
        End Try

    End Sub '_btnModifyGrpCreateModifyTabProductTbcMain_Click(sender,e)

    Private Sub _btnOrderGrpOrderFuelTabProductTbcMain_Click( 'button Order Fuel
            sender As Object,
            e As EventArgs
            ) _
        Handles _
            btnOrderGrpOrderFuelTabProductTbcMain.Click

        _fillingStation.orderFuel(
            cboProductIDTabProductTbcMain.Text.Trim,
            Date.Now,
            txtTrxIdGrpOrderFuelTabProductTbcMain.Text.Trim,
            CDec(txtQtyGrpOrderFuelTabProductTbcMain.Text.Trim)
            )

        'reset
        txtTrxIdGrpOrderFuelTabProductTbcMain.Clear()
        txtQtyGrpOrderFuelTabProductTbcMain.Clear()

    End Sub '_btnOrderGrpOrderFuelTabProductTbcMain_Click(sender,e)
    Private Sub _btnProcessTestData_Click(
            sender As Object,
            e As EventArgs) _
        Handles _
            btnProcessTestDataTabFilesTbcMain.Click

        _processTestData()

    End Sub '_btnCreateProduct_Click(sender, e)

    Private Sub _btnOpenFileTabFilesTbcMain_Click(
            sender As Object,
            e As EventArgs) _
        Handles _
            btnOpenFileTabFilesTbcMain.Click

        Try
            If txtFilenameTabFilesTbcMain.Text = "" Then
                MessageBox.Show("The file name must be entered.")
                txtFilenameTabFilesTbcMain.SelectAll()
                txtFilenameTabFilesTbcMain.Focus()
            Else
                _fillingStation.readFromFile(txtFilenameTabFilesTbcMain.Text.Trim)
            End If
        Catch ex As FileNotFoundException
            MessageBox.Show("ERROR: Input file not found...")
            txtFilenameTabFilesTbcMain.SelectAll()
            txtFilenameTabFilesTbcMain.Focus()
            Exit Sub
        Catch ex As Exception
            '... Device not ready exception
        End Try

    End Sub '_btnOpenFileTabFilesTbcMain_Click(sender,e)

    Private Sub _btnSaveFileTabFilesTbcMain_Click(
            sender As Object,
            e As EventArgs
            ) _
        Handles _
            btnSaveFileTabFilesTbcMain.Click

        MessageBox.Show("The transaction records are writing into txt file.")

        _fillingStation.writeToFile()

    End Sub '_btnSaveFileTabFilesTbcMain_Click(sender,e)

    Private Sub _btnClearTrxLog_Click(
            sender As Object,
            e As EventArgs) _
        Handles _
            btnClearTrxLogTabFilesTbcMain.Click

        _clearTrxLog()

    End Sub '_btnClearTrxLog_Click

    Private Sub _btnResetAll_Click(
            sender As Object,
            e As EventArgs) _
        Handles _
            btnResetAllTabFilesTbcMain.Click

        _resetAll()
        _frmMain_Load(sender, e)

    End Sub '_btnResetAll_Click(sender,e)

    Private Sub _btnDisplayFillingStation(
            sender As Object,
            e As EventArgs) _
        Handles _
            btnDisplayFillingStationStatusTabFilesTbcMain.Click

        _displayFillingStation()

    End Sub '_btnDisplayFillingStation(sender,e)

    '********** User-Interface Event Procedures
    '             - Initiated automatically by system

    Private Sub _frmMain_Load(
            sender As Object,
            e As EventArgs) _
        Handles _
            MyBase.Load

        _initializeBusinessLogic()
        _initializeUserInterface()

    End Sub '_frmMain_Load(sender,e)

    Private Sub _txtTrxLog_TextChanged(
            sender As Object,
            e As EventArgs) _
        Handles _
            txtTrxLog.TextChanged

        txtTrxLog.SelectionStart =
            txtTrxLog.TextLength
        txtTrxLog.ScrollToCaret()

    End Sub '_txtTrxLog_TextChanged(sender,e)

    '********** Business Logic Event Procedures
    '             - Initiated as a result of business logic
    '               method(s) running
    Private Sub _customerAdded(
           ByVal sender As Object,
           ByVal e As EventArgs
           ) _
       Handles _
           mFillingStation.FillingStation_CustomerAdded

        Dim theFillingStation_EventArgs_CustomerAdded _
            As FillingStation_EventArgs_loyaltyCustomerAdded
        Dim theCustomer As LoyaltyCustomer

        theFillingStation_EventArgs_CustomerAdded = CType(
            e, FillingStation_EventArgs_loyaltyCustomerAdded
            )
        theCustomer = theFillingStation_EventArgs_CustomerAdded.theCustomer

        With theCustomer
            lstLoyaltyCustomerTabSummaryTbcMain.Items.Add(.id)
            cboLoyaltyCustomerIDTabLoyaltyCustomerTbcMain.Items.Add(.id)
            cboCustomerIdTabSalesTbcMain.Items.Add(.id)
            'lblMemberAgeGrpCreateTabLoyaltyCustomerTbcMain.Text = Convert.ToString(theCustomer.membershipAge)
            lblLoyaltyCustomerCountTabSummaryTbcMain.Text = mFillingStation.numCustomers.ToString
        End With
        lstLoyaltyCustomerTabSummaryTbcMain.SelectedItem = lstLoyaltyCustomerTabSummaryTbcMain.Items.Count - 1

        txtTrxLog.Text &=
            vbCrLf _
            & _fillingStation.ToString _
            & vbCrLf _
            & theCustomer.ToString
        txtInfoTabLoyaltyCustomerTbcMain.Text &=
            vbCrLf _
            & "- LOYALTY CUSTOMER ADDED: " & theCustomer.ToString

    End Sub '_customerAdded(sender,e)

    Private Sub _productAdded(
            ByVal sender As Object,
            ByVal e As EventArgs
            ) _
        Handles _
            mFillingStation.FillingStation_ProductAdded

        Dim theFillingStation_EventArgs_ProductAdded _
            As FillingStation_EventArgs_ProductAdded
        Dim theProduct As Product

        theFillingStation_EventArgs_ProductAdded = CType(
            e, FillingStation_EventArgs_ProductAdded
            )
        theProduct = theFillingStation_EventArgs_ProductAdded.theProduct

        With theProduct
            lstProductTabSummaryTbcMain.Items.Add(.id)
            cboProductIDTabProductTbcMain.Items.Add(.id)
            cboProductIdTabSalesTbcMain.Items.Add(.id)
            lblProductCountTabSummaryTbcMain.Text = mFillingStation.numProducts.ToString
        End With

        txtTrxLog.Text &=
            vbCrLf _
            & _fillingStation.ToString _
            & vbCrLf _
            & theProduct.ToString
        txtInfoTabProductTbcMain.Text &=
            vbCrLf _
            & "- PRODUCT ADDED: " & theProduct.ToString

    End Sub '_productAdded(sender,e)

    Private Sub _saleAdded(
            ByVal sender As Object,
            ByVal e As EventArgs
            ) _
        Handles _
            mFillingStation.FillingStation_SaleAdded

        Dim theFillingStation_EventArgs_SaleAdded _
            As FillingStation_EventArgs_SaleAdded
        Dim theSale As Sale

        theFillingStation_EventArgs_SaleAdded = CType(
            e,
            FillingStation_EventArgs_SaleAdded
            )
        theSale = theFillingStation_EventArgs_SaleAdded.theSale

        With theSale
            lstSaleTabSummaryTbcMain.Items.Add(.id)
            lblSaleCountTabSummaryTbcMain.Text = mFillingStation.numSales.ToString
        End With

        txtMetricsGrpMetricsTabSummaryTbcMain.Text &=
            vbCrLf _
            & "Total Fuel Sale: " &
            _fillingStation.calProductTypeTotalSale(ProductType.FUEL).ToString("C2") _
            & "                                                    " _
            & "Total Carwash Sale: " &
            _fillingStation.calProductTypeTotalSale(ProductType.CARWASH).ToString("C2") _
            & "                                                           " _
            & "Total Misc Sale: " &
            _fillingStation.calProductTypeTotalSale(ProductType.MISC).ToString("C2") _
            & vbCrLf _
            & "Total Fuel Sale Tax Amount: " &
            _fillingStation.calProductTypeTotalTaxAmount(ProductType.FUEL).ToString("C2") _
            & "                                  " _
            & "Total Carwash Sale Tax Amount: " &
            _fillingStation.calProductTypeTotalTaxAmount(ProductType.CARWASH).ToString("C2") _
            & "                                       " _
            & "Total Misc Sale Tax Amount: " &
            _fillingStation.calProductTypeTotalTaxAmount(ProductType.MISC).ToString("C2") _
            & vbCrLf _
            & "Total Percentage of Sale in Fuel: " &
            _fillingStation.calPercentageSalePerProductType(ProductType.FUEL).ToString("P0") _
            & "                              " _
            & "Total Percentage of Sale in Carwash: " &
            _fillingStation.calPercentageSalePerProductType(ProductType.CARWASH).ToString("P0") _
            & "                                    " _
            & "Total Percentage of Sale in Misc: " &
            _fillingStation.calPercentageSalePerProductType(ProductType.MISC).ToString("P0") _
            & vbCrLf _
            & "Total Percentage of number of sale in Fuel: " &
            _fillingStation.calPercentageNumSalePerProductType(ProductType.FUEL).ToString("P0") _
            & "              " _
            & "Total Percentage of number of sale in Carwash: " &
            _fillingStation.calPercentageNumSalePerProductType(ProductType.CARWASH).ToString("P0") _
            & "                  " _
            & "Total Percentage of number of sale in Carwash: " &
            _fillingStation.calPercentageNumSalePerProductType(ProductType.MISC).ToString("P0") _
            & vbCrLf _
            & "Average Amount Per Sale: " &
            _fillingStation.calAverageAmountPerSale(_fillingStation).ToString("C2") _
            & "                                   " _
            & "The Smallest Amount of Sale: " &
            _fillingStation.getSmallestAmountInSale(_fillingStation).ToString("C2") _
            & "(ProductType: " _
            & _fillingStation.transferFromEnum(_fillingStation.getSmallestSaleProductType(_fillingStation)) & ")" _
            & "             " _
            & "The Biggest Amount of Sale: " &
            _fillingStation.getBiggestAmountInSale(_fillingStation).ToString("C2") _
            & "(ProductType: " _
            & _fillingStation.transferFromEnum(_fillingStation.getLargestSaleProductType(_fillingStation)) & ")" _


        'show total amount of customer in summary tab
        If theSale.customer IsNot Nothing Then
            txtMetricsGrpMetricsTabSummaryTbcMain.Text &=
                      vbCrLf _
                      & "Total Amount of " & theSale.customer.id & ": " _
                      & _fillingStation.calTotalAmountPerCustomer(theSale.customer).ToString _
                      & "      " _
                      & "Total Number Sale of " & theSale.customer.id & ": " _
                      & _fillingStation.calTotalNumSalePerCustomer(theSale.customer).ToString _
                      & vbCrLf
        End If


        txtTrxLog.Text &=
            vbCrLf _
            & _fillingStation.ToString _
            & vbCrLf _
            & theSale.ToString
        txtInfoTabSalesTbcMain.Text &=
            vbCrLf _
            & "- SALE ADDED: " & theSale.ToString

    End Sub '_saleAdded(sender,e)

    Private Sub _transactionAdded(
            ByVal sender As Object,
            ByVal e As EventArgs
            ) _
        Handles _
            mFillingStation.FillingStation_TransactionAdded

        Dim theFillingStation_EventArgs_TransactionAdded _
            As FillingStation_EventArgs_TransactionAdded
        Dim theTransaction As Transaction

        theFillingStation_EventArgs_TransactionAdded = CType(
            e,
            FillingStation_EventArgs_TransactionAdded
            )
        theTransaction = theFillingStation_EventArgs_TransactionAdded.theTransaction

        With theTransaction
            lstTrxTabSummaryTbcMain.Items.Add(.id)
            lblTrxCountTabSummaryTbcMain.Text = mFillingStation.numTransactions.ToString
        End With

        txtTrxLog.Text &=
            vbCrLf _
            & _fillingStation.ToString _
            & vbCrLf _
            & theTransaction.ToString
        txtInfoTabSummaryTbcMain.Text &=
            vbCrLf _
            & "- TRANSACTION ADDED: " & theTransaction.ToString

    End Sub '_TransactionAdded(sender,e)

    Private Sub _fuelTankAdded(
            ByVal sender As Object,
            ByVal e As EventArgs
            ) _
        Handles _
            mFillingStation.FillingStation_FuelTankAdded


        Dim theFillingStation_EventArgs_FuelTankAdded _
            As FillingStation_EventArgs_FuelTankAdded
        Dim theFuelTank As FuelTank

        theFillingStation_EventArgs_FuelTankAdded = CType(
            e,
            FillingStation_EventArgs_FuelTankAdded
            )
        theFuelTank = theFillingStation_EventArgs_FuelTankAdded.theFuelTank

        With theFuelTank
            lstFuelTankTabSummaryTbcMain.Items.Add(.id)
            lblFuelTankCountTabSummaryTbcMain.Text = mFillingStation.numFuelTanks.ToString
        End With

        txtTrxLog.Text &=
            vbCrLf _
            & _fillingStation.ToString _
            & vbCrLf _
            & theFuelTank.ToString
        txtInfoTabProductTbcMain.Text &=
            vbCrLf _
            & "- FUEL TANK ADDED: " & theFuelTank.ToString

    End Sub '_FuelTankAdded(sender,e)

    Private Sub _productModified(
            ByVal sender As Object,
            ByVal e As EventArgs
            ) _
        Handles _
            mFillingStation.FillingStation_ProductModified

        Dim theFillingStation_EventArgs_ProductModified _
            As FillingStation_EventArgs_ProductModified
        Dim theProduct As Product

        theFillingStation_EventArgs_ProductModified = CType(
            e, FillingStation_EventArgs_ProductModified
            )
        theProduct = theFillingStation_EventArgs_ProductModified.theProduct

        With theProduct
            txtInfoTabProductTbcMain.Text &=
                      vbCrLf & "( PRODUCT MODIFIED: " _
                      & "product id=" & cboProductIDTabProductTbcMain.Text _
                      & ", tax rate=" & nudTaxRateGrpCreateModifyTabProductTbcMain.Value.ToString _
                      & ", price per unit=" & nudPricePerUnitGrpCreateModifyTabProductTbcMain.Value.ToString _
                      & ", loyalty customer discount=" & nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value.ToString _
                      & ", reward discount per unit=" & nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value.ToString _
                      & " )"
        End With

        txtTrxLog.Text &=
            vbCrLf _
            & _fillingStation.ToString _
            & vbCrLf _
            & "( PRODUCT MODIFIED: " _
            & "product id=" & cboProductIDTabProductTbcMain.Text _
            & ", tax rate=" & nudTaxRateGrpCreateModifyTabProductTbcMain.Value.ToString _
            & ", price per unit=" & nudPricePerUnitGrpCreateModifyTabProductTbcMain.Value.ToString _
            & ", loyalty customer discount=" & nudLoyaltyDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value.ToString _
            & ", reward discount per unit=" & nudRewardDiscountPerUnitGrpCreateModifyTabProductTbcMain.Value.ToString _
            & " )"

    End Sub '_productModified(sender,e)

    Private Sub _fuelTankRefilled(
            ByVal sender As Object,
            ByVal e As EventArgs) _
        Handles _
            mFillingStation.FillingStation_FuelTankRefilled

        Dim theFillingStation_EventArgs_FuelTankRefilled _
           As FillingStation_EventArgs_FuelTankRefilled
        Dim theProduct As Product

        theFillingStation_EventArgs_FuelTankRefilled = CType(
            e, FillingStation_EventArgs_FuelTankRefilled
            )
        theProduct = theFillingStation_EventArgs_FuelTankRefilled.theProduct

        With theProduct
            txtInfoTabProductTbcMain.Text &=
                       vbCrLf & "- FUEL TANK REFILLED: " _
                       & theProduct.fuelTank.ToString
        End With

    End Sub '_fuelTankRefilled(sender,e)

#End Region 'Event Procedures

#Region "Events"
    '******************************************************************
    'Events
    '******************************************************************

    'No Events are currently defined.
    'These are all public.

#End Region 'Events

End Class 'FrmMain
