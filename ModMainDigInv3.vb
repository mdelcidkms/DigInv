Option Strict Off
Option Explicit On
Imports System.Data.OracleClient
Module ModMain
    'Program written by Yadira Cerda on 10/09/17
    'This program will be used only in JFC Mexico branch
    'This program will be called by invoice program to generate the text data for
    'Digital invoice CFDI version 3.3
    'Modified on July 31 so that when liquor invoices are not separated (entered by phone), it won't error out (ie do not separate on pdr nor xml).

    Dim XtraCustomsInfo, DTL, HDR, TOTAL, ErrMsgLog As String
    Dim ErrFlag As Boolean
    Dim IVATasa As Double
    Dim SlsTaxFlag As Boolean
    Dim SlsTaxInfo, TotalTax As String
    Dim TotalTax_New As Double
    Dim TotInvs As Short
    Dim TotInvAmt As Double
    Dim SeparateLiqTaxFlag As String
    Dim CannotContinueErrFlag, DiscExists As Boolean
    Dim DiscPer As Double
    Dim RFC As String
    Dim TotIEPS_Prt As String
    Dim IEPS_NO_SeparateXml As Boolean
    Dim TotAmtBeforeTaxes As Double
    Dim TermsDate As String
    Dim DTLCPT, COMCP, HDRCP, DTLCP, TOTCP, FolioCP As String
    Dim QtyTot As Short
    Dim TotWeight As Double
    Dim LiquorPresent As Boolean
    Dim IEPSnoIVA As Double
    Dim SorianaErr As Boolean 'Soriana Folio de Entrega

    Dim IEPSTasaGlobal, TotAmtToBeTaxed, TotAmtNotTaxable, TotIEPSAmt As Double

    Dim ArrIeps_Sep_Xml(,) As Object 'define the number of dimensions
    Dim TotPercentages As Short
    Dim HdrNo_SeparateXmlFlag As Boolean

    Const DirToOutputError As String = "C:\JETS\Spool\JMXDigInvCMLogs\"
    Dim IEPSperLiter As Double
    Const JFCID As String = "1" 'ID del Emisor (transmitter) 
    Const cdoSendUsingPickup As Short = 1 'Send message using the local SMTP service pickup directory.
    Const cdoSendUsingPort As Short = 2 'Send the message using the network (SMTP over the network).
    Const cdoAnonymous As Short = 0 'Do not authenticate
    Const cdoBasic As Short = 1 'basic (clear-text) authentication
    Const cdoNTLM As Short = 2 'NTLM
    Const AoApro As String = "2012" : Const NoApro As String = "32191" : Const FormaDePago As String = "PAGO EN UNA SOLA EXHIBICION"
    Const JFCRFC As String = "JME9707037KA"
    Const Expedido As String = "06700" '"CIUDAD DE MEXICO" '04-01-22 expedido LUGEXP is now codigo postal (zip code)
    Const JFCName As String = "JFC DE MEXICO" : Const JFCGln As String = "7504016094006"
    Const Regimen As String = "Régimen General de Ley Personas Morales"
    Dim TotImp, TotImp_Prt As Double
    Dim TotLines As Short ' Importe a imprimir may be different than total importe for cadena if no IEPS desglose.
    Dim DirToOutputData As String
    Dim DiscAmt, TotDiscAmt As Double 'discount to be used for free items as well as for freight charge
    Dim UsoCFDI As String 'per Ma Elena, USOCFDI it might be G01 or G03 so customer master will hold the value (for Invoices), for Nota de cargo, value will be PO1
    Dim ClaveFormaPago As String '99 is used for everything except nota de cargo to Comercial City Fresko want 15
    Dim ItemCatalog, DocName, UOMCatalog As String 'moved to general 

    '***********************************************************************'
    'Name: Main
    'Description: The starting point of the program that opens DB connection, gets branch information and invoice data. 
    'Params: N/A
    'Return Value: N/A
    'Precondition(s): N/A
    'Postcondition(s): DB connection is closed, invoices have been written to designated path. 
    '***********************************************************************'
    Public Sub Main()
        If Not DBOpened Then OpenDatabase()
        GetBranchName()
        RetrieveInvoices()
        End
    End Sub

    '***********************************************************************'
    'Name: RetrieveInvoices
    'Description: Retrieves the invoices to be processed, sets the details for the invoices and outputs the information for the day to the designated file
    'Params: N/A
    'Return Value: N/A
    'Precondition(s): Database connection has been successfuly opened and branch name has been set to glbal variable gs_BranchName 
    'Postcondition(s): On success the output file has been written with the digital invoice details. On error the error message has been written to the error log 
    '***********************************************************************'
    Private Sub RetrieveInvoices()

        Dim SQL, SqlTot As String
        Dim RS, RSTot As Object
        Dim t As Short
        Dim DueDate As String 'Field(8) ConditionsofPymnt Due date
        Dim ExtraCharge As String
        Dim TotalSales As String
        Dim MetodoPago, MontoLetra, Section, Email, NumDpt As String
        Dim Copies As Short
        Dim Folio, Serie, IepsID As String
        Dim TipoDoc, InvMsg, MSG123 As String
        Dim FE As String
        Dim FEPos As Short
        Dim Cita As String
        Dim CitaPos As Short
        Dim RefInvoice, RELTIP, UUID As String
        Dim TotIEPTra26_5 As Double



        Dim Exported As String
        Dim OCM, CMOrdUpd As OracleCommand
        Dim ODR As OracleDataReader
        Dim OTR As OracleTransaction
        Dim DT, DTTOT As DataTable
        Dim Adpt As OracleDataAdapter

        On Error GoTo ErrHndlr

        SQL = "select exported,moddate, moduser from sohdr where ordstateid = 'R' and (exported = 'N'  or exported = 'C' ) order by ordernum"

        CMOrdUpd = New OracleCommand(SQL, conn)
        DT = New DataTable
        DT.Load(CMOrdUpd.ExecuteReader)
        ODR = CMOrdUpd.ExecuteReader

        If DT.Rows.Count = 0 Then Exit Sub

        SQL = "select s.ordernum,s.ordstateid,s.sectioncode,s.customerid,s.custordnum,s.shiplane,s.salesmanid,c.groupid, invopt_liqtaxseparate, " & vbNewLine
        SQL = SQL & "s.exported,s.msgprintcode,s.message,s.shipdate,s.shipvia,s.orderdate,s.FRGCHARGID,s.DEPTCODE,  NVL(C.DISTANCE, 200) DISTANCE, " & vbNewLine
        SQL = SQL & "i.invhdrnum,i.totalsales,i.totaltax,i.totalmisc,i.totalredem,i.totalfreig,i.invdate, FRGHTALLOW, JMX_REGIMEN_INV, " & "n.name shipname,n.address shipaddr, n.address2 shipaddr2, n.address3 shipaddr3, " & vbNewLine
        SQL = SQL & "n.city shipcity, n.state shipstate, n.zipcode shipzip, n.country shipcountry, n.phonenum shipphone, n.GLOBALLOCNUM shipgln, " & vbNewLine
        SQL = SQL & " c.name,c.SALSTXPCNT,c.taxid RFC, c.email, c.JFCVENDORNUM, c.GLOBALLOCNUM billgln, c.invmsg1, invmsg2, invmsg3, c.memo, " & vbNewLine
        SQL = SQL & "nvl(c.invopt_liqtaxseparate,'N')separateliqtax,c.phonenum,c.CREDITCODEID, c.CODONLIQINV,c.DISCPCNT,c.NETDAYS,c.ARDAYS,c.SALSTXPCNT, nvl(C.jmx_diginv_uso, 'G01') usocfdi, " & vbNewLine
        SQL = SQL & " c.address, c.address2, c.address3, c.city, c.state, c.zipcode,c.country, st.abbreviation soldstateabbr, ss.abbreviation shipstateabbr, " & vbNewLine
        SQL = SQL & "NVL(c.JMX_PMT_METHOD_INV,'PPD')METODOPAGO, NVL(c.JMX_PMT_FORM_INV,'99') FORMAPAGO, nvl(dockid, 'N') dockid From " & "sohdr s, invhdr i, customer c, shipname n, state st , state ss Where " & vbNewLine
        SQL = SQL & "s.ordernum = i.invhdrnum and s.customerid = c.customerid and c.customerid = n.customerid (+) and c.state = st.state (+) and n.state = ss.state (+) and ordstateid = 'R' " & " and (exported = 'N' or exported = 'C') order by invhdrnum "

        OCM = New OracleCommand(SQL, conn)
        DT = New DataTable
        DT.Load(OCM.ExecuteReader)
        TotInvs = DT.Rows.Count()

        Dim SeparateIEPSAmt As Double
        Dim f As Short
        Dim SugarTaxAmt As Double
        Dim Found As Boolean
        Dim InvhdrNo As String
        Dim counter As Integer = 0

        If DT.Rows.Count() > 0 Then

            IEPSperLiter = GetSugarCuota(Format(DT.Rows(0)("invdate"), "yyyy"))

            Copies = 2
            '**********check if production or test environment ***************
            If gs_Hoststring = gs_UserName Then DirToOutputData = "\\10.39.1.31\MASFAC_ESC\DATOS\" Else DirToOutputData = "C:\JETS\Spool\JMXDigInvCM\"

            For Each DR As DataRow In DT.Rows
                InvMsg = vbNullString : TotIEPS_Prt = vbNullString : NumDpt = vbNullString : MSG123 = vbNullString
                RefInvoice = vbNullString : UUID = vbNullString : RELTIP = vbNullString 'clear only used for supermercados
                ErrFlag = False
                SlsTaxFlag = False 'If iva for any item
                SlsTaxInfo = vbNullString
                DiscExists = False : DiscPer = 0
                If DR("totalfreig") < 0 Then
                    DiscExists = True
                    DiscPer = DR("FRGHTALLOW")
                End If

                TotalTax = "0.00" : TotalTax_New = 0
                Section = vbNullString : IEPS_NO_SeparateXml = False
                SorianaErr = False
                TotAmtToBeTaxed = 0 : TotAmtNotTaxable = 0 : TotIEPSAmt = 0 : TotImp = 0 : TotImp_Prt = 0 : IEPSTasaGlobal = 0 : TotLines = 0 : TotAmtBeforeTaxes = 0
                Email = vbNullString & DR("Email").ToString()

                UsoCFDI = DR("UsoCFDI")
                TotInvAmt = DR("TotalSales") + DR("TotalTax") + DR("totalredem") + DR("totalfreig") + DR("totalmisc")
                NumDpt = vbNullString & DR("DEPTCODE")
                ClaveFormaPago = DR("formapago")

                TipoDoc = "1" '1 (ingreso) or E (Egreso) (was 3, now I)
                Folio = gs_BranchN & DR("invhdrnum")

                If DR("totalmisc") > 0 Then
                    DocName = "NOTA DE CARGO"
                    Serie = "C"
                    FolioCP = DR("invhdrnum")
                    TotalSales = Format(TotInvAmt, "##0.00") 'NOTA DE CARGO IS REPORTED AS SALES

                    If Trim(vbNullString & DR("Memo")) = "SUPERMERCADO" Then
                        RELTIP = "02"
                        FEPos = InStr(vbNullString & DR("message"), "FACTURA ")
                        If FEPos > 0 Then
                            RefInvoice = Trim(Mid(DR("message"), FEPos + 8, 9)) 'GRAB 9 CHARS AFTER FACTURA
                            UUID = GetRelUuid(RefInvoice)
                            If UUID = vbNullString Then SorianaErr = True
                        Else
                            SorianaErr = True
                        End If
                    End If
                Else
                    DocName = "FACTURA"
                    Serie = "A"
                    TotalSales = DR("TotalSales").ToString("0.00")

                    If Trim(vbNullString & DR("invmsg1")) <> vbNullString Then MSG123 = "1:  " & DR("invmsg1").ToString().TrimEnd()
                    MSG123 = MSG123 & " " & DR("invmsg2").ToString.Trim & " " & DR("invmsg3").ToString.Trim
                End If
                RFC = StripRFC(vbNullString & DR("RFC"))
                MetodoPago = DR("MetodoPago") & vbNewLine 'Sets the method payment (i.e. 'PPD' or 'PUE', defaults to 'PPD' if null) 
                DiscAmt = 0 : ExtraCharge = vbNullString
                MontoLetra = MoneyPhrase(TotInvAmt)
                TermsDate = vbNullString
                GetTermMX(DR)
                IVATasa = Val(DR("SALSTXPCNT"))
                If ((vbNullString & DR("msgprintcode")) = "B" Or (vbNullString & DR("msgprintcode")) = "I") And Trim(vbNullString & DR("message")) <> vbNullString Then
                    InvMsg = "2:  " & DR("message")
                End If
                InvMsg = Replace(InvMsg, vbTab, "")
                IepsID = "GST"
                If RFC = "NWM9709244W4" Then IepsID = vbNullString

                HDR = "E" & vbNewLine & "VERSIO" & Space(25 - Len("VERSIO")) & "4.0" & vbNewLine & "TRADPP" & Space(25 - Len("TRADPP")) & "MASTEDI" & vbNewLine
                HDR = HDR & "SERFOL" & Space(25 - Len("SERFOL")) & Serie & vbNewLine
                HDR = HDR & "CTPPRO" & Space(25 - Len("CTPPRO")) & "ZZ" & vbNewLine & "NUMFOL" & Space(25 - Len("NUMFOL")) & Folio & vbNewLine
                HDR = HDR & "SERFOL" & Space(25 - Len("SERFOL")) & Serie & vbNewLine
                HDR = HDR & "FECEXP  " & Format(Now, "yyyy-MM-ddTHH:mm:ss") & vbNewLine & "NOAPRO  " & NoApro & vbNewLine & "AOAPRO  " & AoApro & vbNewLine
                HDR = HDR & "CVEREGIMEN" & Space(15) & "601" & vbNewLine & "CVEFORPAG" & Space(16) & ClaveFormaPago & vbNewLine

                HDR = HDR & "USOCFDI     " & UsoCFDI & vbNewLine
                HDR = HDR & "CVETIPDOC       I" & vbNewLine & "METPAG     PPD" & vbNewLine
                HDR = HDR & "CODMETPAG  " & MetodoPago & vbNewLine & "REGIMEN " & Regimen & vbNewLine & "FORPAG  " & FormaDePago & vbNewLine & "NUMCHE  " & vbNewLine & "TIPDOC  " & TipoDoc & vbNewLine
                HDR = HDR & "NOMDOC  " & DocName & vbNewLine & "FUNDOC  " & "O" & vbNewLine & "TIPMON  " & "MXN" & vbNewLine & "SW_TC   " & "0" & vbNewLine & "TIPCAM  " & "1" & vbNewLine
                HDR = HDR & "DIAPAG  " & (vbNullString & DR("ardays")) & vbNewLine & "PDPPAG  " & (vbNullString & DR("DISCPCNT")) & vbNewLine & "MDPPAG  " & vbNewLine
                HDR = HDR & "NUMEOC  " & DR("custordnum") & vbNewLine & "FECHOC  " & Format(DR("orderdate"), "yyyy-MM-dd") & vbNewLine & "FECCON  " & Format(DR("ShipDate"), "yyyy-MM-dd") & vbNewLine
                HDR = HDR & "FECPAG  " & TermsDate & vbNewLine
                HDR = HDR & "LUGEXP  " & Expedido & vbNewLine & "MSG123  " & MSG123 & vbNewLine
                HDR = HDR & "NOTAS1  " & InvMsg & vbNewLine & "NOTAS2  " & vbNewLine & "NOTAS3  " & MontoLetra & vbNewLine & "AGENTE  " & (vbNullString & DR("salesmanid")) & vbNewLine
                HDR = HDR & "PEDIDO  " & gs_BranchN & "-" & DR("ordernum") & vbNewLine & "TRANSP  " & (vbNullString & DR("shipvia")) & vbNewLine
                HDR = HDR & "NUMDPT  " & NumDpt & vbNewLine & "NOMDPT  " & vbNewLine
                HDR = HDR & "NUMERO_IMP " & "1" & vbNewLine & "COPIAS  " & Copies & vbNewLine & "IEPS_ID " & IepsID & vbNewLine
                HDR = HDR & "REFFAC  " & vbNewLine & "NUMCLI  " & gstrBranchAlpha & "-" & DR("customerid") & vbNewLine & "NUMSAP  " & DR("invhdrnum") & vbNewLine
                HDR = HDR & "REMDES  " & vbNewLine

                HDRCP = HDRCP & "VERSIO  4.0" & vbNewLine & "SERFOL  CP " & vbNewLine & "NUMFOL  " & DR("invhdrnum") & vbNewLine
                HDRCP = HDRCP & "FECEXP  " & Format(Now, "yyyy-MM-ddTHH:mm:ss") & vbNewLine
                HDRCP = HDRCP & "CVETIPDOC   T" & vbNewLine & "TIPDOC  7 " & vbNewLine & "TIPMON  XXX" & vbNewLine & "TIPCAM  1" & vbNewLine
                HDRCP = HDRCP & "USOCFDI S01" & vbNewLine

                HDR = HDR & "CVEREGREC   " & DR("JMX_REGIMEN_INV") & vbNewLine

                If RFC = "TSO991022PB6" Then
                    Select Case (vbNullString & DR("groupid"))
                        Case Is = "3925"
                            HDR = HDR & "TIPADD     3" & vbNewLine
                            HDR = HDR & "CONTRA  " & vbNewLine
                        Case Else
                            FEPos = InStrRev(vbNullString & DR("message"), "FE")
                            If FEPos > 0 Then
                                FE = Trim(Mid(DR("message"), FEPos + 2)) 'GRAB EVERYTHING AFTER FE

                                If Not IsNumeric(FE) Then
                                    SorianaErr = True
                                Else
                                    HDR = HDR & "TDA_CONTRA     " & FE & vbNewLine
                                End If

                                CitaPos = InStrRev(UCase(vbNullString & DR("message")), "CITA")
                                If CitaPos > 0 Then
                                    Cita = Trim(Mid(DR("message"), CitaPos + 4, FEPos - CitaPos - 4))
                                    If Not IsNumeric(Cita) Then
                                        SorianaErr = True
                                    Else
                                        HDR = HDR & "CONTRA     " & Cita & vbNewLine
                                    End If
                                Else 'CitaPos = 0, error, no cita(date) found
                                    SorianaErr = True 'EITHER NO CITA OR NO FOLIO DE ENTRADA, DO NOT PROCESS THE ORDER!!!!!
                                End If
                            Else 'FEPos = 0, Problem, no folio de entrada found
                                SorianaErr = True
                            End If
                            HDR = HDR & "PO_NUMEOC     " & "SI" & vbNewLine
                    End Select
                Else
                    HDR = HDR & "CONTRA  " & vbNewLine
                End If
                'EMISOR DEL DOCUMENTO
                HDR = HDR & "RFCEMI  " & JFCRFC & vbNewLine & "NOMEMI  " & JFCName & vbNewLine & "EANEMI  " & JFCGln & vbNewLine & "NUMEMI  " & DR("JFCVENDORNUM").ToString().TrimEnd() & vbNewLine
                HDR = HDR & "CALEMI  " & "AV.AÑO DE JUAREZ NO. 160-B" & vbNewLine
                HDR = HDR & "NEXEMI  " & vbNewLine & "NINEMI  " & vbNewLine & "COLEMI  " & "COLONIA GRANJAS SAN ANTONIO" & vbNewLine & "LOCEMI  " & "MEXICO" & vbNewLine
                HDR = HDR & "MUNEMI  " & "IZTAPALAPA" & vbNewLine & "ESTEMI  " & "CIUDAD DE MÉXICO" & vbNewLine
                HDR = HDR & "REFEMI  " & vbNewLine & "TELEMI  " & "(55)5686-88-93" & vbNewLine

                HDRCP = HDRCP & "RFCEMI " & JFCRFC & vbNewLine & "NOMEMI  " & JFCName & vbNewLine
                HDRCP = HDRCP & "CALEMI  " & "AV.AÑO DE JUAREZ NO. 160-B" & vbNewLine
                HDRCP = HDRCP & "NEXEMI  " & vbNewLine & "NINEMI  " & vbNewLine & "COLEMI  " & "1322" & vbNewLine & "LOCEMI  " & "09" & vbNewLine
                HDRCP = HDRCP & "MUNEMI  " & "007" & vbNewLine & "ESTEMI  " & "CMX" & vbNewLine '01/11/23 use CMX instead of DIF per MasterEDI
                HDRCP = HDRCP & "PAIEMI  MEX" & vbNewLine & "CODEMI  09070" & vbNewLine & "EANEMI  " & JFCGln & vbNewLine
                HDRCP = HDRCP & "TIPCOM T " & vbNewLine & "CVEREGIMEN   601 " & vbNewLine & "USOCFDI G01" & vbNewLine

                'DATOS DEL RECEPTOR -- CLIENTE
                HDR = HDR & "RFCREC  " & RFC & vbNewLine & "NOMREC  " & Trim(DR("Name")) & vbNewLine
                HDR = HDR & "CALREC  " & DR("address").ToString.Trim & vbNewLine
                HDR = HDR & "COLREC  " & DR("address2").ToString.Trim & vbNewLine & "LOCREC  " & DR("city").ToString.Trim & vbNewLine
                HDR = HDR & "MUNREC  " & DR("address3").ToString.Trim & vbNewLine & "ESTREC  " & DR("soldstateabbr") & vbNewLine & "PAIREC  " & DR("COUNTRY").ToString.Trim & vbNewLine

                If vbNullString & DR("zipcode") = vbNullString Then HDR = HDR & "CODREC  " & "99999" & vbNewLine Else HDR = HDR & "CODREC  " & DR("zipcode") & vbNewLine

                HDR = HDR & "EANREC  " & DR("BillGLN") & vbNewLine & "TELREC  " & DR("phonenum").ToString().TrimEnd() & vbNewLine & "MAIL    " & Email.ToString() & vbNewLine

                '*******================CLIENTE, PERO USAMOS JFCRFC FOR TRASLADO =========================******
                HDRCP = HDRCP & "RFCREC  " & JFCRFC & vbNewLine
                HDRCP = HDRCP & "NOMREC  " & JFCName & vbNewLine
                HDRCP = HDRCP & "ESTREC CMX " & vbNewLine
                HDRCP = HDRCP & "CODREC 09070" & vbNewLine
                '************SHIP TO INFO ************
                HDR = HDR & "RFCENT  " & RFC & vbNewLine & "NOMENT  " & DR("shipname").ToString.Trim & vbNewLine & "CALENT  " & DR("shipaddr").ToString.Trim & vbNewLine
                HDR = HDR & "COLENT  " & DR("shipaddr2").ToString.Trim & vbNewLine & "LOCENT  " & DR("shipcity").ToString.Trim & vbNewLine & "ESTENT  " & DR("shipstateabbr") & vbNewLine
                HDR = HDR & "MUNENT  " & DR("shipaddr3").ToString.Trim & vbNewLine & "ESTENT  " & DR("shipstate") & vbNewLine & "PAIENT  " & DR("shipcountry").ToString.Trim & vbNewLine
                HDR = HDR & "CODENT  " & DR("shipzip") & vbNewLine & "EANENT  " & DR("ShipGLN") & vbNewLine

                HDRCP = HDRCP & vbNewLine & "NOMENT  " & DR("shipname").ToString.Trim & vbNewLine & "CALENT  " & DR("shipaddr").ToString.Trim & vbNewLine
                If vbNullString & DR("shipzip").ToString.Trim = vbNullString Then
                    HDRCP = HDRCP & "CODENT  " & DR("zipcode") & vbNewLine
                Else
                    HDRCP = HDRCP & "CODENT  " & DR("shipzip") & vbNewLine
                End If
                HDRCP = HDRCP & "PAIENT MEX " & vbNewLine
                HDRCP = HDRCP & "CVEREGIMEN  601" & vbNewLine & vbNewLine
                '******************* CTP ******************* 
                HDRCP = HDRCP & "NUMERO_IMP 1 " & vbNewLine & "COPIAS 2 " & vbNewLine
                HDRCP = HDRCP & "D " & vbNewLine & "CANTID 1 " & vbNewLine & "NUMLIN  1" & vbNewLine & "ESTILV " & vbNewLine & "CVEPRODSERV    50221300" & vbNewLine
                HDRCP = HDRCP & "CVEUNIDAD " & vbNewLine & "DESCRI   Productos alimenticios " & vbNewLine & "UNIDAD PZA " & vbNewLine
                HDRCP = HDRCP & "CODUPC " & vbNewLine & "CVESKU " & vbNewLine
                HDRCP = HDRCP & "VALUNI 0.00" & vbNewLine & "IMPORT   0.00" & vbNewLine & "PBRUDE   0.00" & vbNewLine & "IMPBRU  0.00" & vbNewLine
                HDRCP = HDRCP & "TDECON 0.00" & vbNewLine & "MDECON   0.00" & vbNewLine & "TDECON0 0.00" & vbNewLine & "MDECON0 0.00" & vbNewLine
                HDRCP = HDRCP & "TDECON1 0.00" & vbNewLine & "MDECON1 0.00" & vbNewLine & "TASIPE 16 " & vbNewLine & "TASIEP  0.00" & vbNewLine
                HDRCP = HDRCP & "MONIPE 0.00" & vbNewLine & "MONIEP  0.00" & vbNewLine & "OBSER " & vbNewLine & "R" & vbNewLine

                HDRCP = HDRCP & "SUBTBR 0.00 " & vbNewLine & "MONDET 0.00 " & vbNewLine & "PRCDSG 0.00 " & vbNewLine
                HDRCP = HDRCP & "SUBTOT 0.00 " & vbNewLine & "SUBTAI 0.00 " & vbNewLine & "IVATRA 0.00 " & vbNewLine
                HDRCP = HDRCP & "TOTIVA 16 " & vbNewLine & "TOTPAG 0.00 " & vbNewLine & "IVATRA1 0.00 " & vbNewLine
                HDRCP = HDRCP & "NIVATR1 IVA " & vbNewLine & "TOTIVA1 16 " & vbNewLine & "TOTTRA 0.00 " & vbNewLine
                HDRCP = HDRCP & "COM_CPT_INICPT " & vbNewLine
                HDRCP = HDRCP & "  COM_CPT_VERSIO 3.1" & vbNewLine
                HDRCP = HDRCP & "COM_CPT_IDCCP   GENERA" & vbNewLine
                HDRCP = HDRCP & " COM_CPT_TRANINT No " & vbNewLine
                HDRCP = HDRCP & "   COM_CPT_FT_CVETRANS 01 " & vbNewLine
                HDRCP = HDRCP & "   COM_CPT_TOTDIST " & DR("distance") & vbNewLine
                HDRCP = HDRCP & "   COM_CPT_TIPEST 2 " & vbNewLine & vbNewLine
                'UBICACION ORIGEN
                HDRCP = HDRCP & "   COM_CPT_INIUBI " & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_UBI_TIPUBI Origen" & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_UBI_IDUBI OR000001 " & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_UBI_RFC " & JFCRFC & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_UBI_NOM " & JFCName & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_UBI_FECHA " & Format(Now, "yyyy-MM-ddT12:00:00") & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_DOM_CAL " & "AV.AÑO DE JUAREZ NO. 160-B" & vbNewLine 'Calle de la dirección del domicilio de la ubicación.
                HDRCP = HDRCP & "      COM_CPT_DOM_COL " & "1322" & vbNewLine '"AV.AÑO DE JUAREZ NO. 160-B" & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_DOM_LOC 09 " & vbNewLine 'DR("city") 
                HDRCP = HDRCP & "      COM_CPT_DOM_MUN " & "007" & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_DOM_EST  CMX" & vbNewLine
                HDRCP = HDRCP & "       COM_CPT_DOM_PAI " & "MEX " & vbNewLine 'DR("COUNTRY") & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_DOM_CP 09070" & vbNewLine
                HDRCP = HDRCP & "   COM_CPT_FINUBI " & vbNewLine & vbNewLine 'Fin de Ubicacion origen

                HDRCP = HDRCP & "      COM_CPT_INIUBI " & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_UBI_TIPUBI Destino " & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_UBI_IDUBI " & "DE000001" & vbNewLine 'id  punto d llegada
                HDRCP = HDRCP & "      COM_CPT_UBI_RFC " & RFC & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_UBI_NOM " & DR("shipname").ToString.Trim & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_UBI_FECHA " & Format(DateAdd(Microsoft.VisualBasic.DateInterval.Hour, 6, Now), "yyyy-MM-ddT12:00:00") & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_UBI_DISTREC " & DR("distance") & vbNewLine 'DistanciaRecorrida

                HDRCP = HDRCP & "      COM_CPT_DOM_CAL " & DR("shipaddr").ToString().TrimEnd() & vbNewLine '
                If vbNullString & DR("shipstateabbr") = vbNullString Then
                    HDRCP = HDRCP & "      COM_CPT_DOM_EST " & DR("soldstateabbr") & vbNewLine
                Else
                    HDRCP = HDRCP & "      COM_CPT_DOM_EST " & DR("shipstateabbr") & vbNewLine
                End If
                HDRCP = HDRCP & "      COM_CPT_DOM_PAI MEX " & vbNewLine
                If vbNullString & DR("shipzip").ToString.Trim = vbNullString Then
                    HDRCP = HDRCP & "      COM_CPT_DOM_CP " & DR("zipcode") & vbNewLine
                Else
                    HDRCP = HDRCP & "      COM_CPT_DOM_CP " & DR("shipzip") & vbNewLine
                End If
                HDRCP = HDRCP & "COM_CPT_FINUBI " & vbNewLine & vbNewLine

                If Trim(RELTIP) <> vbNullString Then HDR = HDR & "RELTIP  " & RELTIP & vbNewLine
                If Trim(UUID) <> vbNullString Then HDR = HDR & "RELUUID1 " & UUID & vbNewLine
                InvhdrNo = DR("invhdrnum")

                Call GetDtl(InvhdrNo)
                Call GetCPVehiculo(InvhdrNo)
                If ("" & DR("FRGCHARGID")) = "Y" Then 'only if freight charge is at a header level, at detail only charges
                    If Val(DR("totalfreig")) < 0 Then 'discount
                        DiscAmt = System.Math.Abs(Val(DR("totalfreig"))) 'ASSUMING DISCOUNT AMT. ONLY COMES FROM FREIGHT (AFTER CHECKING DB)
                        TotDiscAmt = TotDiscAmt + DiscAmt
                        TotImp_Prt = TotImp_Prt + Val(DR("totalfreig")) 'DiscAmt
                    ElseIf Val(DR("totalfreig")) > 0 Then  'Freight Charge
                        ExtraCharge = Format(DR("totalfreig"), "0.00")
                        DTL = DTL & "CANTID  " & "1" & vbNewLine & "CANTID_EA  " & "1" & vbNewLine & "DESCRI  " & "SERVICIO DE ENTREGA" & vbNewLine & "CANPAQ  " & "1" & vbNewLine
                        DTL = DTL & "CANEMP  " & "1" & vbNewLine & "UNIDAD  " & "EA" & vbNewLine & vbNewLine
                        DTL = DTL & "PBRUDE  " & ExtraCharge & vbNewLine & "PBRUDE_IEPS  " & ExtraCharge & vbNewLine & "IMPBRU  " & ExtraCharge & vbNewLine & "TDECON  " & "0" & vbNewLine & "MDECON  " & "0.00" & vbNewLine 'NO DESCUENTO EN DETALLE HARD CODE 0
                        DTL = DTL & "VALUNI  " & ExtraCharge & vbNewLine & "IMPORT  " & ExtraCharge & vbNewLine & "IMPBRU_PRT  " & CDbl(ExtraCharge) & vbNewLine
                        DTL = DTL & "IMPIVAIEPS  " & CDbl(ExtraCharge) & vbNewLine
                        TotImp_Prt = TotImp_Prt + CDbl(ExtraCharge)
                    End If
                    TotLines = TotLines + 1
                End If
                TOTAL = vbNewLine & TOTAL & "R" & vbNewLine & "TOTLPF  " & TotLines & vbNewLine
                TOTCP = TOTCP & "TOTLPF    1 " & vbNewLine

                'SUBTBR value has total before ANY TAXES-  SHOULD BE SUM OF IMPORT VALUES
                If DocName = "NOTA DE CARGO" Then TotAmtBeforeTaxes = TotAmtNotTaxable + TotAmtToBeTaxed 'NOTA DE CARGO CALCULATION ONLY.  SHOULD BE SAME AMOUNT AF SUBTOT

                If Val(TotalTax) = 0 Then IVATasa = 0

                TOTAL = TOTAL & "TOTCAJ   " & vbNewLine & "TOTCAN  " & vbNewLine

                TOTAL = TOTAL & "MONDET  " & TotDiscAmt & vbNewLine & "PRCDSG  " & DiscPer & vbNewLine

                TOTAL = TOTAL & "SUBTOT_PRT  " & Format(TotImp_Prt, "#,##0.00;-#,##0.00;""") & vbNewLine

                TOTCP = TOTCP & vbNewLine

                If IVATasa > 0 Then TOTAL = TOTAL & "TOTIVA  " & IVATasa & vbNewLine & "IVATRA  " & TotalTax & vbNewLine

                If IVATasa = 0 And TotInvAmt > 0 Then TOTAL = TOTAL & "FCTTOTIVA   0.00" & vbNewLine & "TOTIVA  0 " & vbNewLine & "IVATRA 0 " & vbNewLine

                SeparateIEPSAmt = 0

                If Trim("" & DR("invopt_liqtaxseparate")) <> vbNullString Then
                    SqlTot = "select 'Tasa' as TIPFACIEP ,b.liqtax as TATIEP, nvl(b.liquorcode,0)liquorcode, " & vbNewLine
                    SqlTot = SqlTot & " b.liqtax/100 as fcttatiep,sum(lineliqtax) as IEPTRA "
                    SqlTot = SqlTot & "from invdtl d, branch_item b where d.itemcode = b.itemcode and invhdrnum = '" & DR("ordernum") & "'" & vbNewLine
                    SqlTot = SqlTot & " and lineliqtax >0 and lineprice > 0   AND B.LIQTAX > 0 group by liqtax , nvl(b.liquorcode,0) "
                    SqlTot = SqlTot & " order by tatiep "

                    If DocName = "NOTA DE CARGO" Then
                        SqlTot = "SELECT 'Tasa' as TIPFACIEP, IEPS_RATE AS TATIEP, LINELIQTAX as IEPTRA, 1 AS LIQUORCODE  "
                        SqlTot = SqlTot & "FROM INVDTL D, remark_code R WHERE INVHDRNUM = " & DR("ordernum")
                        SqlTot = SqlTot & " AND R.REMARKID = D.REMARKID   AND LINELIQTAX  > 0"
                        HDRCP = vbNullString 'Nota de cargo does no have item informaion, for the moment we do no generate carta porte.
                    End If

                    OCM = New OracleCommand(SqlTot, conn)
                    DTTOT = New DataTable
                    Adpt = New OracleDataAdapter(OCM)
                    Adpt.Fill(DTTOT)

                    t = 0 'Initialize t because it will be used in case there is sugary drinks

                    If DTTOT.Rows.Count() > 0 Then
                        For Each DRTOT As DataRow In DTTOT.Rows
                            f = InStr("" & DR("invopt_liqtaxseparate"), "C")
                            If f = 0 Then f = InStr("" & DR("invopt_liqtaxseparate"), "Y")
                            If DRTOT("tatiep") = 8 And f > 0 Then
                                t = t + 1
                                TOTAL = TOTAL & "TIPFACIEP" & t & Space(5) & DRTOT("TIPFACIEP") & vbNewLine
                                TOTAL = TOTAL & "TATIEP" & t & Space(5) & GetDecimalWOZero(DRTOT("tatiep")) & vbNewLine
                                TOTAL = TOTAL & "IEPTRA" & t & Space(5) & GetDecimalWOZero(DRTOT("IEPTRA")) & vbNewLine

                                SeparateIEPSAmt = SeparateIEPSAmt + GetDecimalWOZero(DRTOT("IEPTRA"))
                            End If

                            '---------------------------------(B) BEER 26.5--OR (Y) EVERYTHING--------------------
                            f = InStr("" & DR("invopt_liqtaxseparate"), "B")
                            If f = 0 Then f = InStr("" & DR("invopt_liqtaxseparate"), "Y")
                            If DRTOT("tatiep") = 26.5 And f > 0 And DRTOT("liquorcode") = 1 Then
                                TotIEPTra26_5 = TotIEPTra26_5 + DRTOT("IEPTRA")
                                SeparateIEPSAmt = SeparateIEPSAmt + GetDecimalWOZero(DRTOT("IEPTRA"))
                            End If
                            '---------------------------------(A) LIQUOR  26.5, 30 or 53  and liquor code (2,3) OR (Y) EVERYTHING--------------
                            '----------------------------------------ADDED SEPARATION IF 26.5 AND LIQCODE 2, (not beer)------------------------
                            f = InStr("" & DR("invopt_liqtaxseparate"), "A")
                            If f = 0 Then f = InStr("" & DR("invopt_liqtaxseparate"), "Y")
                            If DRTOT("tatiep") = 26.5 And f > 0 And DRTOT("liquorcode") > 1 Then

                                TotIEPTra26_5 = TotIEPTra26_5 + DRTOT("IEPTRA") 'combine total 26.6 amt if previous beer amt
                                SeparateIEPSAmt = SeparateIEPSAmt + GetDecimalWOZero(DRTOT("IEPTRA"))
                            End If

                            f = InStr("" & DR("invopt_liqtaxseparate"), "A")
                            If f = 0 Then f = InStr("" & DR("invopt_liqtaxseparate"), "Y")
                            If DRTOT("tatiep") = 30 And f > 0 Then '
                                t = t + 1
                                TOTAL = TOTAL & "TIPFACIEP" & t & Space(5) & DRTOT("TIPFACIEP") & vbNewLine
                                TOTAL = TOTAL & "TATIEP" & t & Space(5) & GetDecimalWOZero(DRTOT("tatiep")) & vbNewLine
                                TOTAL = TOTAL & "IEPTRA" & t & Space(5) & GetDecimalWOZero(DRTOT("IEPTRA")) & vbNewLine
                                SeparateIEPSAmt = SeparateIEPSAmt + GetDecimalWOZero(DRTOT("IEPTRA"))
                            End If

                            '---------------------------------(A) LIQUOR 30 OR (Y) EVERYTHING ------------
                            f = InStr("" & DR("invopt_liqtaxseparate"), "A")
                            If f = 0 Then f = InStr("" & DR("invopt_liqtaxseparate"), "Y")
                            If DRTOT("tatiep") = 53 And f > 0 Then
                                t = t + 1
                                TOTAL = TOTAL & "TIPFACIEP" & t & Space(5) & DRTOT("TIPFACIEP") & vbNewLine
                                TOTAL = TOTAL & "TATIEP" & t & Space(5) & GetDecimalWOZero(DRTOT("tatiep")) & vbNewLine
                                TOTAL = TOTAL & "IEPTRA" & t & Space(5) & GetDecimalWOZero(DRTOT("IEPTRA")) & vbNewLine
                                SeparateIEPSAmt = SeparateIEPSAmt + GetDecimalWOZero(DRTOT("IEPTRA"))
                            End If
                        Next
                    End If

                    '--------check if any beer ieps is pending
                    If TotIEPTra26_5 > 0 Then
                        t = t + 1
                        TOTAL = TOTAL & "TIPFACIEP" & t & Space(5) & "Tasa" & vbNewLine
                        TOTAL = TOTAL & "TATIEP" & t & Space(5) & "26.5" & vbNewLine
                        TOTAL = TOTAL & "IEPTRA" & t & Space(5) & TotIEPTra26_5 & vbNewLine
                        TotIEPTra26_5 = 0
                    End If

                    '  NOW SUGARY DRINKS  ---NOW IS 1.17 per liter, so using constant IEPSperLiter
                    f = InStr("" & DR("invopt_liqtaxseparate"), "S")
                    If f = 0 Then f = InStr("" & DR("invopt_liqtaxseparate"), "Y")
                    If f > 0 Then
                        SqlTot = "select 'Cuota' as TIPFACIEP , " & IEPSperLiter & " as TATIEP, " & IEPSperLiter & " as FCTTATIEP,sum(lineliqtax) as IEPTRA " & vbNewLine
                        SqlTot = SqlTot & "from invdtl d, branch_item b where d.itemcode = b.itemcode and b.liqtax2 > 0 and " & vbNewLine
                        SqlTot = SqlTot & "invhdrnum = '" & DR("ordernum") & "' and lineliqtax >0 and lineprice > 0 and b.liqtax2 > 0 group by 'Cuota' " 'liqtax2" group by cuota instead of liqtax2.


                        OCM = New OracleCommand(SqlTot, conn)
                        DTTOT = New DataTable
                        Adpt = New OracleDataAdapter(OCM)
                        Adpt.Fill(DTTOT)

                        SugarTaxAmt = 0
                        If DTTOT.Rows.Count > 0 Then
                            For Each DRTOT As DataRow In DTTOT.Rows
                                'In case no other types of ieps were found, sugary drinks is the only ieps, otherwise it will come up as zero
                                t = t + 1
                                TOTAL = TOTAL & "TIPFACIEP" & t & Space(5) & DRTOT("TIPFACIEP") & vbNewLine
                                TOTAL = TOTAL & "TATIEP" & t & Space(5) & GetDecimalWOZero(DRTOT("tatiep")) & vbNewLine
                                TOTAL = TOTAL & "IEPTRA" & t & Space(5) & GetDecimalWOZero(DRTOT("IEPTRA")) & vbNewLine
                                SugarTaxAmt = GetDecimalWOZero(DRTOT("IEPTRA"))

                                t = t + 1
                            Next
                        End If
                    End If
                End If

                TOTAL = TOTAL & "##############" & vbNewLine
                TOTAL = TOTAL & "SUBTBR  " & (TotImp_Prt + TotDiscAmt) & vbNewLine
                TOTAL = TOTAL & "##############" & vbNewLine

                If DocName = "NOTA DE CARGO" Then
                    If CDbl(TotalTax) = 0 Then 'TOTIVA 0%
                        TOTAL = TOTAL & "IMPIVATRA1 " & (TotAmtToBeTaxed + TotIEPSAmt) & vbNewLine & "IVATRA1 0" & vbNewLine
                        TOTAL = TOTAL & "TIPFACIVA1 Tasa " & vbNewLine & "TOTIVA1 0 " & vbNewLine
                    Else ' TOTIVA 16%
                        TOTAL = TOTAL & "IMPIVATRA1 " & (TotAmtToBeTaxed + TotIEPSAmt) & vbNewLine & "IVATRA1 " & TotalTax & vbNewLine
                        TOTAL = TOTAL & "TIPFACIVA1 Tasa " & vbNewLine & "TOTIVA1 16 " & vbNewLine
                    End If
                Else '---------END ADDITION REGARDING NOTA DE CARGO TOTAL
                    If TotAmtNotTaxable > 0 Then
                        TOTAL = TOTAL & "IVATRA1 0" & vbNewLine
                        TOTAL = TOTAL & "TIPFACIVA1 Tasa " & vbNewLine & "TOTIVA1 0 " & vbNewLine
                    Else
                        If TotInvAmt > 0 Then 'only do below if invoice amount > 0 - not for free items
                            TOTAL = TOTAL & "IMPIVATRA1 " & TotAmtToBeTaxed & vbNewLine & "IVATRA1 " & TotalTax & vbNewLine
                            TOTAL = TOTAL & "TIPFACIVA1 Tasa " & vbNewLine & "TOTIVA1 16 " & vbNewLine
                        End If
                    End If
                End If

                If TotAmtToBeTaxed > 0 And TotAmtNotTaxable > 0 Then 'two types of iva 0% and 16%
                    TOTAL = TOTAL & "IMPIVATRA2 " & TotAmtToBeTaxed & vbNewLine '
                    TOTAL = TOTAL & "IVATRA2  " & TotalTax & vbNewLine '
                    TOTAL = TOTAL & "TIPFACIVA2 Tasa " & vbNewLine & "TOTIVA2 16 " & vbNewLine
                End If

                TOTAL = TOTAL & "PRT_IEPTRA  " & TotIEPS_Prt & vbNewLine
                TOTAL = TOTAL & "IVARET  " & "0" & vbNewLine & "ISRRET  " & "0" & vbNewLine

                If IEPS_NO_SeparateXml And TotalTax_New > 0 And (Format(TotAmtNotTaxable + TotAmtToBeTaxed + TotalTax_New, "###.00")) <> Format(TotInvAmt, "###.00") Then
                    TOTAL = TOTAL & "TOTPAG  " & TotInvAmt & vbNewLine
                Else
                    TOTAL = TOTAL & "TOTPAG  " & TotInvAmt & vbNewLine
                End If

                Exported = "N"
                If Not SorianaErr Then Exported = "Y"

                SQL = "Update sohdr set moduser = 'DigInv3', moddate = '" + Format(Now, "dd-MMM-yy") + "', exported = '" + Exported + "' where ordstateid = 'R' and (exported = 'N'  or exported = 'C' ) and ordernum = '" + DR("ordernum") + "'"

                If counter = 0 Then OTR = conn.BeginTransaction(IsolationLevel.ReadCommitted) Else OTR = conn.BeginTransaction()

                CMOrdUpd = New OracleCommand(SQL, conn, OTR)
                CMOrdUpd.ExecuteNonQuery()
                OTR.Commit()
                counter += 1


                'TODO: Looks like this is where we can add the method of payment check, if it's 'PUE' then null out the HDRCP variable and carta porte does not get generated
                'Check DocName = "FACTURA" and method payment is 'PUE' 
                'Does this mean that we just completely do not write out the Invoice? Or where do we specify this is a complemento de pago? 
                If DocName = "NOTA DE CARGO" Then HDRCP = vbNullString 'Nota de cargo does no have item informaion, for the moment we do no generate carta porte.
                If Left(DR("dockid"), 1) <> "F" Then HDRCP = vbNullString
                If Not SorianaErr Then
                    If DR("exported") <> "C" Then CkFileExists(DirToOutputData, HDR & DTL & TOTAL, "INV-" & Folio & "-" & DR("RFC") & "_" & Format(Now, "yyyyMMdd_HH-mm") & ".txt")
                    If Trim(HDRCP) <> vbNullString Then CkFileExists(DirToOutputData, HDRCP & DTLCPT & DTLCP & COMCP & TOTCP, "CTP-" & DR("invhdrnum") & " - " & Format(Now, "yyyy-MM-dd") & " at " & Format(Now, "HH-mm") & ".txt")
                End If
                HDR = vbNullString : DTL = vbNullString : TOTAL = vbNullString : XtraCustomsInfo = vbNullString : TotDiscAmt = 0 : HdrNo_SeparateXmlFlag = False : TotPercentages = 0
                HDRCP = vbNullString : DTLCP = vbNullString : DTLCPT = vbNullString : TOTCP = vbNullString : COMCP = vbNullString : TotWeight = 0
                IEPSnoIVA = 0 : TotIEPTra26_5 = 0
                LiquorPresent = False
            Next

        Else
            ErrMsgLog = New String("*", 50) & vbNewLine & Format(Now, "MM/dd/yy HH:mm") & " -> ERROR:  No records found for invoices to print, but the program was run." & vbNewLine & "Error on RetrieveInvoices Routine.  Prog:  DigInv3." & vbNewLine & "PROGRAM ENDED WITH OUT PROCESSING ANYTHING!!!!" & vbNewLine & "SQL: " & vbNewLine & SQL & vbNewLine
            CkFileExists(DirToOutputError, ErrMsgLog, "ERR-Dig-Inv3_3.log")
        End If
        Exit Sub
ErrHndlr:

        ErrMsgLog = New String("*", 50) & vbNewLine & Format(Now, "MM/dd/yy HH:mm") & " -> ERROR:  " & Err.Number & "-->" & Err.Description & "." & vbNewLine & "Error on RetrieveInvoices Routine.  Prog:  DigInv3." & vbNewLine & "Total Invoices:  " & TotInvs & vbNewLine & vbNewLine & "SQL: " & vbNewLine & SQL & vbNewLine
        CkFileExists(DirToOutputError, ErrMsgLog, "ERR-Dig-Inv3_3.log")
        SendEmail(ErrMsgLog)
        End
    End Sub

    '***********************************************************************'
    'Name: GetDecimalWOZero
    'Description:Takes in a decimal value and removes trailing 0s
    'Params: 
    '   - val : the decimal value to be formatted
    'Return Value: 
    '   result: the formatted version of the decimal passed in w/ trailing 0s removed
    'Precondition(s): N/A   
    'Postcondition(s): Return value has been formatted to the "G29" format specifier 
    '***********************************************************************'
    Public Function GetDecimalWOZero(ByVal val As Decimal) As String
        Dim result As String = val.ToString("G29")
        Return result
    End Function

    '***********************************************************************'
    'Name: GetCPVehiculo
    'Description: Sets driver information to be later appended to the final output
    'Params: 
    '   InvNum : The invoice number to look up which truck driver will be delivering the shipment 
    'Return Value: N/A
    'Precondition(s): N/A
    'Postcondition(s): Truck driver information has been set to global variable COMCP 
    '***********************************************************************'
    Private Sub GetCPVehiculo(ByRef InvNum As String)
        Dim OCM As OracleCommand
        Dim dtc As DataTable
        Dim dtr As DataRow

        ' Join shiproute_info to TRUCK on TRUCKID (LEFT JOIN keeps old behavior if TRUCK row is missing)
        ' Alias TRUCK fields to avoid ambiguity with truckdriver columns.
        Dim Sql As String =
        "select c.*, " &
        "t.LICENSENUM as TRUCK_LICENSENUM, " &
        "t.YEAR_OF_VEHICLE as TRUCK_YEAR_OF_VEHICLE, " &
        "t.INSURANCE as TRUCK_INSURANCE, " &
        "t.POLICYNUM as TRUCK_POLICYNUM, " &
        "t.INSURANCE_DOWNPAYMENT as TRUCK_INSURANCE_DOWNPAYMENT " &
        "from sohdr a " &
        "join shiproute_info b on a.shipdate = b.shipdate and a.dockid = b.dockid " &
        "join truckdriver c on b.driverid = c.driverid " &
        "left join TRUCK t on b.TRUCKID = t.TRUCKID " &
        "where a.ordernum = '" & InvNum & "'"

        Try
            OCM = New OracleCommand(Sql, conn)
            dtc = New DataTable
            dtc.Load(OCM.ExecuteReader)

            ' Defaults used when TRUCK is missing or fields are null/blank
            Dim truckPlate As String = "3901CM"
            Dim truckYear As String = "2014"
            Dim truckInsurance As String = "TOKIO MARINE CIA DE SEGUROS"
            Dim truckPolicy As String = "TLJMX000244800"
            Dim truckPremium As String = "900000"

            'TRANSPORTATION
            COMCP = "COM_CPT_INIAUTO " & vbNewLine & "COM_CPT_AUT_SCT TPAF02" & vbNewLine
            COMCP = COMCP & "   COM_CPT_AUT_PSCT Permiso no contemplado en el catálogo " & vbNewLine & "COM_CPT_AUT_SUTIPREM1 " & vbNewLine
            COMCP = COMCP & "   COM_CPT_AUT_ASEGRESP Qualitas Compañía De Seguros, S.A. de C.V. " & vbNewLine & "COM_CPT_AUT_POLIRESP 0003945047 " & vbNewLine
            COMCP = COMCP & "   COM_CPT_AUT_CONVEH C2" & vbNewLine

            If dtc.Rows.Count > 0 Then
                dtr = dtc.Rows(0)

                ' TRUCK fields (aliased)
                If Trim("" & dtr("TRUCK_LICENSENUM")) <> "" Then truckPlate = Trim("" & dtr("TRUCK_LICENSENUM"))
                If Trim("" & dtr("TRUCK_YEAR_OF_VEHICLE")) <> "" Then truckYear = Trim("" & dtr("TRUCK_YEAR_OF_VEHICLE"))
                If Trim("" & dtr("TRUCK_INSURANCE")) <> "" Then truckInsurance = Trim("" & dtr("TRUCK_INSURANCE"))
                If Trim("" & dtr("TRUCK_POLICYNUM")) <> "" Then truckPolicy = Trim("" & dtr("TRUCK_POLICYNUM"))
                If Trim("" & dtr("TRUCK_INSURANCE_DOWNPAYMENT")) <> "" Then truckPremium = Trim("" & dtr("TRUCK_INSURANCE_DOWNPAYMENT"))
            End If

            ' Vehicle + insurance data (from TRUCK when available)
            COMCP = COMCP & "   COM_CPT_AUT_PLACAV " & truckPlate & " " & vbNewLine
            COMCP = COMCP & "   COM_CPT_AUT_ANIOV " & truckYear & vbNewLine
            COMCP = COMCP & "COM_CPT_AUT_ASEGCAR " & truckInsurance & "  " & vbNewLine
            COMCP = COMCP & "   COM_CPT_AUT_POLICAR " & truckPolicy & " " & vbNewLine
            COMCP = COMCP & "COM_CPT_AUT_PRIMSEG " & truckPremium & vbNewLine
            COMCP = COMCP & "COM_CPT_AUT_PESBRU  2 " & vbNewLine

            If LiquorPresent Then
                COMCP = COMCP & "   COM_CPT_AUT_ASEGMED Atlas " & vbNewLine
                COMCP = COMCP & "   COM_CPT_AUT_POLMED 1010101  " & vbNewLine
            End If

            COMCP = COMCP & "COM_CPT_FINAUTO " & vbNewLine & "COM_CPT_INIFIGTRA " & vbNewLine & "COM_CPT_FIG_TIPFIG 01" & vbNewLine

            If dtc.Rows.Count > 0 Then
                ' Driver info from truckdriver (c.*)
                dtr = dtc.Rows(0)
                COMCP = COMCP & "COM_CPT_FIG_RFCFIG " & dtr("RFC") & vbNewLine & "COM_CPT_FIG_NUMLIC " & dtr("LICENSENUM") & vbNewLine
                COMCP = COMCP & "COM_CPT_FIG_NOMFIG " & dtr("FULLNAME") & vbNewLine & "COM_CPT_FINFIGTRA " & vbNewLine & "COM_CPT_FINCPT" & vbNewLine
            Else
                COMCP = COMCP & "COM_CPT_FIG_RFCFIG " & JFCRFC & vbNewLine & "COM_CPT_FIG_NUMLIC " & "680000033715 " & vbNewLine
                COMCP = COMCP & "COM_CPT_FIG_NOMFIG " & "Jose Alberto Salas Aguilar " & vbNewLine & "COM_CPT_FINFIGTRA " & vbNewLine & "COM_CPT_FINCPT" & vbNewLine
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    '***********************************************************************'
    'Name: GetDtl
    'Description: Sets the details for an invoice and all its items
    'Params: 
    '   - InvNum : The invoice number to be looked up 
    'Return Value: N/A
    'Precondition(s): N/A
    'Postcondition(s): Details for the invoice has been set to the global variable DTL 
    '***********************************************************************'
    Private Sub GetDtl(ByRef InvNum As String)

        Dim SQL As String
        Dim RSd, RS, RSc, Rsd2 As Object
        Dim d As Short
        Dim PricePlusLiqTax As Double
        Dim IvaAmt, Units, Tst As String
        Dim CustomsInfo, FoundCustomsDefault As Boolean
        Dim IEPSTasa, IvaTasaDtl As Double
        Dim UnitPrice, UPC, Importe As String 'Importe (Price * Qty)
        Dim EaImpIvaIeps, CsImpIvaIeps As Double
        Dim FirstPortionOfName As Short
        Dim LineIVA, ItemDesc, TmpTotImp As String
        Dim CSPrice, Eaprice, LiqTax2 As Double
        Dim IEPSAmt, IEPSAmt_Prt As String
        Dim CsImporte As Double 'ONLY USED WHEN CS & EA FOR LIQUOR NO DESGLOSE CALCULATION... ON EACH
        Dim DocID, PortName, SQLInsert, DocDate As String
        Dim SQLCustom As String
        'CARTA PORTE
        Dim DtlWeight As Double
        Dim LiqPresentDtl As Boolean
        Dim Peli, CvePeli As String 'Peli Material Peligroso (Dangerous Material)
        Dim IvhdrNo, LineNo As String

        Dim OCM As OracleCommand
        Dim OTR As OracleTransaction
        Dim dt, dtc As DataTable
        Dim ItemCode As String
        Dim ODR As OracleDataReader
        Dim DRc As DataRow
        Dim Adpt As OracleDataAdapter

        On Error GoTo ErrHndlr
        '----------SIR FOR CITY FRESKO DIGITAL INVOICE IN EA ONLY -------------------
        If RFC = "CCF121101KQ4" Then 'carta porte added gross and ne weight to sql 
            SQL = "select i.unitupccode,i.jancode,i.UNIT_ML,b.liqtax,nvl(b.liqtax2,0)liqtax2,b.eapercs,b.saltaxcode,b.eapercs, nvl(GROSSWEIGHT,0)  gross, round(nvl(GROSSWEIGHT,0) /2.205 ,2) grosskg ,nvl(netweight,0) net, " & vbNewLine
            SQL = SQL & "b.mx_sat_catalog_id as satItem, b.mx_unitofmeasure as satUOM, nvl(b.liquorcode,0) liquorcode, " & vbNewLine
            SQL = SQL & "D.INVHDRNUM, D.LINENUM, D.ITEMCODE, D.REMARKID, D.LINEDESC, (D.CSSHIPPED * B.EAPERCS + D.EASHIPPED) AS EASHIPPED, 0 AS CSSHIPPED," & vbNewLine
            SQL = SQL & "ROUND(d.INVCSPRICE / B.EAPERCS * 100000000) / 100000000 As INVEAPRICE, D.LINETAX, " & vbNewLine
            SQL = SQL & "D.LINEPRICE,D.LINECOST,D.LINEREDEM,D.LINETOTAL,D.INVCSPRICE,D.LINELIQTAX, D.LINESALESTAX "
            SQL = SQL & "from invdtl d, jfcitem i, branch_item b where invhdrnum = '" & InvNum & "' " & vbNewLine
            SQL = SQL & "and d.itemcode = i.itemcode and i.itemcode = b.itemcode and (eashipped > 0 or csshipped > 0)order by d.itemcode "
        Else 'carta porte added gross and ne weight to sql 
            SQL = "select i.unitupccode,i.jancode,b.liqtax,nvl(b.liqtax2,0)liqtax2,b.eapercs,b.saltaxcode,b.eapercs, nvl(GROSSWEIGHT,0)  gross, round(nvl(GROSSWEIGHT,0) /2.205 ,2) grosskg ,nvl(netweight,0) net, " & vbNewLine
            SQL = SQL & "b.mx_sat_catalog_id as satItem, b.mx_unitofmeasure as satUOM, nvl(b.liquorcode,0) liquorcode, " & vbNewLine
            SQL = SQL & "d.* from invdtl d, jfcitem i, branch_item b where invhdrnum = '" & InvNum & "' " & vbNewLine
            SQL = SQL & "and d.itemcode = i.itemcode and i.itemcode = b.itemcode and (eashipped > 0 or csshipped > 0)order by d.itemcode "
        End If

        TotAmtToBeTaxed = 0 'Used to print next to iva msg to know total amt of taxable items

        OCM = New OracleCommand(SQL, conn)
        dt = New DataTable
        Adpt = New OracleDataAdapter(OCM)
        Adpt.Fill(dt)

        If dt.Rows.Count > 0 Then
            d = 0 ' If CS and EA, we separate into two lines, so we can't use i as detail line number.
            For Each DR As DataRow In dt.Rows
                ItemCode = DR("ItemCode")
                DiscAmt = 0 : DiscExists = False : DiscPer = 0 'Clear up discounts when free items
                IEPSTasa = Val(vbNullString & DR("liqtax"))
                LiqTax2 = Val(vbNullString & DR("LiqTax2"))
                If IEPSTasa <> 0 Then IEPSTasaGlobal = IEPSTasa ' if stmt to getlast ieps tasa in the ord. When last item did not have tasa, was upadating tasa global to 0.
                IEPSAmt = vbNullString : LineIVA = vbNullString : IvaAmt = vbNullString 'use when ea and CS exists to properly calculate line iva (in DB amt is CS and EA combined)
                IEPSAmt_Prt = vbNullString
                If RFC = "NWM9709244W4" Then
                    UPC = ("" & DR("jancode")) 'for WM use Jancode first, if not found, use UPC.
                    If Val(UPC) = 0 Then UPC = ("" & DR("unitupccode"))
                Else 'Other than WM check upc first, if null, then look for Jan code
                    UPC = ("" & DR("unitupccode"))
                    If Val(UPC) = 0 Then UPC = ("" & DR("jancode"))
                End If
                If Val(UPC) = 0 Then UPC = vbNullString 'if both upc and jan code are 0, then don't print all zeroes, leave it blank
                If RFC = "NWM9709244W4" Then UPC = Format(Val(UPC), "0000000000000") 'WM wants 13 digits field. ---WAL-MART

                ItemDesc = vbNullString & DR("linedesc").ToString().TrimEnd()
                ItemCatalog = vbNullString & DR("satItem")
                UOMCatalog = vbNullString & DR("satuom")

                '!---------------INVALID ITEM CATALOGS--------------------
                If Trim(ItemCatalog) = vbNullString Then ItemCatalog = "50171500" 'defatult item catalog 
                If Trim(ItemCatalog) = "50424800" Then ItemCatalog = "50171500" 'THIS ITEM CATALOG DOES NOT EXIST-- INVALID
                If Trim(ItemCatalog) = "50347000" Then ItemCatalog = "50171500" 'THIS ITEM CATALOG DOES NOT EXIST-- INVALID

                '------MATERIAL PELIGROSO--- If "Si" we need dangerous product info. If "No", we do not put the info
                Peli = vbNullString : CvePeli = vbNullString : LiqPresentDtl = False '01/23/24 added  LiqPresentDtl = false 02/07/24
                Select Case Trim(ItemCatalog)
                    Case "50202200", "50202201", "50202206", "50202210", "50121500", "50171708", "50121537", "14111519" 'Pescado, licor y vino, 50171708 
                        If DR("liquorcode") = "3" Then
                            Peli = "Si"
                            CvePeli = "3065"
                            LiquorPresent = True 'to send insurance info when peli is present
                            LiqPresentDtl = True
                        Else
                            Peli = "No"
                        End If
                    Case "15111505" 'Gas butano Item 60080
                        Peli = "Si"
                        CvePeli = "1011"
                        LiquorPresent = True 'to send insurance info when peli is present
                        LiqPresentDtl = True
                    Case Else
                        LiqPresentDtl = False
                End Select

                d = d + 1
                Eaprice = DR("INVEAPRICE")
                CSPrice = DR("INVCSPRICE")

                If DR("eashipped") > 0 And DR("csshipped") > 0 Then
                    Units = "CE"
                ElseIf DR("eashipped") > 0 Then
                    Units = "EA"
                    UOMCatalog = "H87" 'JMX NOW WANTS TO CHANGE UNIT OF MEASURE TO BE 'XBX' IF CASES, OR 'H87' IF EACHES.
                Else 'Either just cases, or no cases and no each.  When xtra charges in detail (AC, tax, etc)
                    Units = "CA"
                    UOMCatalog = "XBX" 'JMX NOW WANTS TO CHANGE UNIT OF MEASURE TO BE 'XBX' IF CASES, OR 'H87' IF EACHES.
                End If
                If DocName = "NOTA DE CARGO" Then
                    UOMCatalog = "ACT"
                    ItemCatalog = "84111506"
                End If

                If DR("linetax") - DR("lineliqtax") > 0 Then 'LineTax has combined tax (iva & ieps), to get iva, substract both taxes.
                    IvaTasaDtl = IVATasa
                    TotAmtToBeTaxed = Val(CStr(TotAmtToBeTaxed)) + Val(DR("lineprice")) + Val(DR("lineliqtax")) 'SUBTAI
                    TotalTax = CStr(Val(TotalTax) + (Val(DR("linetax")) - Val(DR("lineliqtax")))) 'calculate totaltax based on detail
                Else
                    IvaTasaDtl = 0
                    TotAmtNotTaxable = Val(CStr(TotAmtNotTaxable)) + Val(DR("lineprice")) 'SUBTSI
                End If

                IvhdrNo = DR("invhdrnum")
                LineNo = DR("linenum")
                SeparateLiqTaxFlag = IepsSeparate(IvhdrNo, LineNo)

                TotAmtBeforeTaxes = TotAmtBeforeTaxes + DR("lineprice")

                DTL = DTL & vbNewLine 'Just to see where detail starts on text file
                DTL = DTL & "D" & vbNewLine

                Select Case Units
                    Case Is = "EA"
                        UnitPrice = Format(Eaprice, "##0.00")
                        Importe = CDbl(DR("lineprice"))
                        TotImp = TotImp + Val(Importe)
                        If Val(CStr(IEPSTasa)) > 0 Then
                            IEPSAmt = CDbl(DR("lineliqtax"))
                        End If

                        If LiqTax2 > 0 Then
                            IEPSAmt = CDbl(DR("lineliqtax"))
                        End If

                        If (SeparateLiqTaxFlag = "Y" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "Y" And LiqTax2 > 0) Then
                            IEPSAmt_Prt = IEPSAmt 'Ieps printing
                            TotIEPS_Prt = CStr(Val(TotIEPS_Prt) + Val(IEPSAmt))
                        End If

                        If (SeparateLiqTaxFlag = "N" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "N" And LiqTax2 > 0) Then 'Price should add IEPS on inv to be sent on EXTRA FIELDS
                            If IEPSTasa > 0 Then UnitPrice = Format(Eaprice + (Eaprice * (IEPSTasa / 100)), "0.00###")

                            Importe = Format(Val(DR("lineprice")) + Val(DR("lineliqtax")), "##0.00")
                            If LiqTax2 > 0 Then UnitPrice = CStr(Val(Importe) / DR("eashipped"))

                            IEPSAmt_Prt = vbNullString
                        End If

                        If IvaTasaDtl > 0 Then
                            LineIVA = CStr(Val(DR("linetax")) - Val(DR("lineliqtax")))
                            TotalTax_New = TotalTax_New + CDbl(LineIVA) '----04/12/19 sir 1794, when ea and cs present one cent difference and it doesn't generate cfdi
                        End If

                        TotImp_Prt = TotImp_Prt + Val(Importe)

                        DTL = DTL & "CANTID  " & DR("eashipped") & vbNewLine & "CANTID_EA  " & DR("eashipped") & vbNewLine & "DESCRI  " & ItemDesc & vbNewLine
                        DTL = DTL & "CANPAQ  " & DR("eashipped") & vbNewLine
                        DTL = DTL & "CANEMP  " & DR("eashipped") & vbNewLine & "UNIDAD  " & Units & vbNewLine & "CVESKU  " & DR("itemcode") & vbNewLine
                        DTL = DTL & "ESTILV     " & DR("itemcode") & vbNewLine 'per Master EDI to be able to print item code on pdf.
                        DTL = DTL & "CVEPRODSERV     " & ItemCatalog & vbNewLine 'New SAT item code
                        DTL = DTL & "CVEUNIDAD           " & UOMCatalog & vbNewLine
                        DTL = DTL & "CODUPC  " & UPC & vbNewLine & "PIEPEM  " & "1" & vbNewLine & "PIEPEM2 " & DR("eapercs") & vbNewLine & "CODDUN  " & vbNewLine

                        DTLCP = DTLCP & "      COM_CPT_INIMER " & vbNewLine & vbNewLine 'Inicio de mercancia
                        DTLCP = DTLCP & "      COM_CPT_MER_BIENTRA " & ItemCatalog & vbNewLine 'BienesTransp (clave de producto)
                        DTLCP = DTLCP & "      COM_CPT_MER_DESCRI " & ItemDesc & vbNewLine
                        DTLCP = DTLCP & "      COM_CPT_MER_CANTID 1" & vbNewLine
                        DTLCP = DTLCP & "      COM_CPT_MER_CVUNID  " & UOMCatalog & vbNewLine

                        DtlWeight = RoundUpToDecimals((DR("eashipped") / DR("eapercs") * DR("grosskg")), 2)
                        If DtlWeight = 0 Then DtlWeight = 0.01
                        TotWeight = TotWeight + DtlWeight
                        DTLCP = DTLCP & "      COM_CPT_MER_PKG " & DtlWeight & vbNewLine

                        'FOR PRODUCTO PELIGROSO WE NEED TO READ A TABLE FOR CLAVEMATERIALPELIGROSO
                        If LiqPresentDtl Then
                            DTLCP = DTLCP & "      COM_CPT_MER_MATPEL Sí  " & vbNewLine 'Material peligroso
                            DTLCP = DTLCP & "      COM_CPT_MER_CVEMATPEL   " & CvePeli & vbNewLine '3065  drinks 24% pero no más de 70%  alcohol , or 1011 gas butano
                            DTLCP = DTLCP & "      COM_CPT_MER_EMB 4G" & vbNewLine '4C1 modified to 4G per Ma Elena
                        End If
                        If Peli = "No" Then
                            DTLCP = DTLCP & "      COM_CPT_MER_MATPEL No " & vbNewLine
                        End If
                        DTLCP = DTLCP & "      COM_CPT_INICANTRAN " & vbNewLine 'Inicio canidad trasladada
                        DTLCP = DTLCP & "         COM_CPT_CMER_CANTID " & DR("eashipped") & vbNewLine
                        DTLCP = DTLCP & "         COM_CPT_CMER_IDORI OR000001" & vbNewLine
                        DTLCP = DTLCP & "         COM_CPT_CMER_IDDES DE000001" & vbNewLine
                        DTLCP = DTLCP & "      COM_CPT_FINCANTRAN" & vbNewLine
                        DTLCP = DTLCP & "COM_CPT_FINMER" & vbNewLine & vbNewLine

                        '----------------------CHANGES FOR FREE PRODUCT-----------------
                        If Eaprice = 0 Then 'free
                            Eaprice = 0.01
                            DiscExists = True
                            DiscPer = 100
                            DiscAmt = 0.01 * DR("eashipped") 'DO NOT Accumulate total discount DiscAmt + (0.01 * DR("EASHIPPED"))
                            IEPSTasa = 0 'In case free item has ieps do not report tasa because it won't have any ieps amount
                            LiqTax2 = 0 'In case free item, do not report any tasa, no ieps amt should be reported
                            TotDiscAmt = TotDiscAmt + DiscAmt
                        End If

                        If RFC <> "PHI830429MG6" Then '-----------------Palacio de Hierro doesn't want 0 when iva is 0
                            If Trim(LineIVA) = vbNullString Then LineIVA = "0.00" : If Trim(IEPSAmt) = vbNullString Then IEPSAmt = "0.00"
                        End If

                        'NO TASIPE When 0, NO MONIPE   When 0, NO TASIEP  When 0, NO MONIEP WHEN 0
                        If Val(CStr(IvaTasaDtl)) > 0 Then 'IF TO JUST SEND THESE WHEN > 0 
                            DTL = DTL & "TASIPE  " & IvaTasaDtl & vbNewLine
                            DTL = DTL & "MONIPE  " & LineIVA & vbNewLine
                        End If

                        '--------------------OBJECTO DE IMPUESTO ------------
                        If Val(CStr(IEPSTasa)) > 0 Or LiqTax2 > 0 Then
                            If SeparateLiqTaxFlag = "N" Then
                                DTL = DTL & "OBJIMP 02" & vbNewLine 'TEMP CHANGED EVERYTHING TO 02
                            Else
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            End If
                            If Val(CStr(IvaTasaDtl)) = 0 Then DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine
                        Else
                            If Val(CStr(IvaTasaDtl)) > 0 Then
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            Else
                                If DiscExists Then
                                    DTL = DTL & "OBJIMP 01" & vbNewLine 'WHEN FREE ITEMS NO IVA SHOULD BE REPORTED
                                    DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine
                                Else
                                    DTL = DTL & "OBJIMP 02" & vbNewLine ' PER MA. ELENA  ALL ITEMS SHOULD BE 02, JFC DOES NOT HAVE 01 TYPE
                                    DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine & "FCTTASIPE   0.000000" & vbNewLine
                                    DTL = DTL & "IMPORTIPE " & Format(Eaprice * DR("eashipped"), "##0.00") & vbNewLine
                                    DTL = DTL & "TIPIPETR Tasa" & vbNewLine & "TASIEP 0" & vbNewLine
                                End If
                            End If
                        End If

                        If IEPSTasa = 8 Then '8% ieps but no iva, detail should show 0% IVA
                            DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine & "FCTTASIPE   0.000000" & vbNewLine
                            DTL = DTL & "IMPORTIPE " & Format((Eaprice * DR("eashipped")) + CDbl(IEPSAmt), "##0.00") & vbNewLine
                            DTL = DTL & "TIPIPETR Tasa" & vbNewLine
                            TotAmtNotTaxable = TotAmtNotTaxable + CDbl(IEPSAmt)
                        End If

                        DTL = DTL & "MONIEP_IEPS  " & IEPSAmt_Prt & vbNewLine
                        DTL = DTL & "IMPIVAIEPS  " & CDbl(DR("linetotal")) & vbNewLine
                        DTL = DTL & "PBRUDE_IEPS  " & Format(CDbl(UnitPrice), "##0.00") & vbNewLine
                        DTL = DTL & "IMPBRU_PRT  " & Importe & vbNewLine

                        '********************************************SIR 1794********************************************************
                        If ((SeparateLiqTaxFlag = "N" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "N" And LiqTax2 > 0)) And IEPS_NO_SeparateXml Then
                            DTL = DTL & "PBRUDE  " & Format(CDbl(UnitPrice), "##0.00") & vbNewLine
                            DTL = DTL & "VALUNI  " & Format(CDbl(UnitPrice), "##0.00") & vbNewLine
                            DTL = DTL & "IMPBRU  " & Format(CDbl(UnitPrice) * DR("eashipped"), "##0.00") & vbNewLine
                            DTL = DTL & "IMPORT  " & Format(CDbl(UnitPrice) * DR("eashipped"), "##0.00") & vbNewLine
                        Else
                            If Val(CStr(IEPSTasa)) > 0 Then DTL = DTL & "TASIEP  " & IEPSTasa & vbNewLine & "MONIEP  " & CDbl(IEPSAmt) & vbNewLine
                            DTL = DTL & "PBRUDE  " & Format(Eaprice, "##0.00") & vbNewLine
                            DTL = DTL & "VALUNI  " & Format(Eaprice, "##0.00") & vbNewLine
                            DTL = DTL & "IMPBRU  " & Format(Eaprice * DR("eashipped"), "##0.00") & vbNewLine
                            DTL = DTL & "IMPORT  " & Format(Eaprice * DR("eashipped"), "##0.00") & vbNewLine
                        End If

                        '---------SUGARY DRINKS------------------
                        If LiqTax2 > 0 Then
                            If DiscAmt = 0 And SeparateLiqTaxFlag = "Y" Then
                                DTL = DTL & "FCTTASIEP   " & IEPSperLiter & vbNewLine
                                DTL = DTL & "TIPIEPTR    Cuota" & vbNewLine

                                DTL = DTL & "IMPORTIEP    " & Format(Val(IEPSAmt) / IEPSperLiter, "0.##") & vbNewLine
                                'SPECITY THE IVA WHEN IT IS 0% like in SOME OF the sugar products which have IEPS but 0% IVA
                                DTL = DTL & "MONIEP " & CDbl(Val(IEPSAmt)) & vbNewLine

                                If Val(CStr(IvaTasaDtl)) = 0 Then
                                    DTL = DTL & "TASIPE 0 " & vbNewLine & "MONIPE 0" & vbNewLine & "FCTTASIPE 0.000000" & vbNewLine
                                    DTL = DTL & "TIPIPETR Tasa " & vbNewLine & "IMPORTIPE " & ((Eaprice * DR("eashipped"))) + CDbl(IEPSAmt) & vbNewLine
                                    IEPSnoIVA = IEPSnoIVA + CDbl(IEPSAmt)
                                Else

                                End If
                            End If
                        End If

                        If DiscExists Then
                            DTL = DTL & "TDECON    " & DiscPer & vbNewLine
                            DTL = DTL & "MDECON    " & (Eaprice * DR("eashipped")) * (DiscPer / 100) & vbNewLine
                        Else
                            DTL = DTL & "TDECON  " & "0" & vbNewLine & "MDECON  " & "0.00" & vbNewLine 'NO DESCUENTO EN DETALLE HARD CODE 0
                        End If

                        DTL = DTL & "NUMLIN    " & d & vbNewLine
                        TotIEPSAmt = TotIEPSAmt + Val(IEPSAmt)
                    Case Is = "CA"
                        UnitPrice = Format(CDbl(CSPrice), "##0.00")
                        Importe = CDbl(DR("lineprice"))
                        TotImp = TotImp + Val(Importe) 'Importe sin ieps
                        If Val(CStr(IEPSTasa)) > 0 Then
                            IEPSAmt = CDbl(DR("lineliqtax")) 'SAT IEPSAmt
                        End If
                        If LiqTax2 > 0 Then
                            IEPSAmt = CDbl(DR("lineliqtax"))
                        End If
                        If (SeparateLiqTaxFlag = "Y" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "Y" And LiqTax2 > 0) Then
                            IEPSAmt_Prt = IEPSAmt 'Ieps printing
                            TotIEPS_Prt = CStr(Val(TotIEPS_Prt) + Val(IEPSAmt))
                        End If
                        If (SeparateLiqTaxFlag = "N" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "N" And LiqTax2 > 0) Then 'Price should add IEPS on inv
                            If IEPSTasa > 0 Then UnitPrice = Format(CSPrice + (CSPrice * (IEPSTasa / 100)), "##0.00000")
                            If LiqTax2 > 0 Then UnitPrice = Format(CSPrice + LiqTax2, "##0.00000")
                            Importe = Format(Val(DR("lineprice")) + Val(DR("lineliqtax")), "##0.00")
                            IEPSAmt_Prt = vbNullString
                        End If
                        If Val(Importe) = 0 And TotInvAmt <> 0 Then 'when eaqty and csqty = 0, but detail exists, probably XTRA CHARGE!!!
                            Importe = Format(CDbl(DR("lineprice")), "##0.00")
                            UnitPrice = "0.00"
                        End If
                        If IvaTasaDtl > 0 Then
                            LineIVA = CStr(Val(DR("linetax")) - Val(DR("lineliqtax")))
                            TotalTax_New = TotalTax_New + CDbl(LineIVA)
                        End If
                        TotImp_Prt = TotImp_Prt + Val(Importe) 'Total importe a imprimir (Accumulate accordingly for printing pursposes which is diff. for SAT purposes)
                        DTL = DTL & "D" & vbNewLine
                        DTL = DTL & "CANTID  " & DR("csshipped") & vbNewLine & "CANTID_CA  " & DR("csshipped") & vbNewLine & "DESCRI  " & ItemDesc.ToString.Trim & vbNewLine
                        DTL = DTL & "CANPAQ  " & DR("csshipped") & vbNewLine
                        DTL = DTL & "CANEMP  " & DR("csshipped") & vbNewLine & "UNIDAD  " & Units & vbNewLine & "CVESKU  " & DR("itemcode") & vbNewLine
                        DTL = DTL & "ESTILV     " & DR("itemcode") & vbNewLine
                        DTL = DTL & "CVEPRODSERV     " & ItemCatalog & vbNewLine
                        DTL = DTL & "CVEUNIDAD     " & UOMCatalog & vbNewLine
                        DTL = DTL & "CODUPC  " & UPC & vbNewLine & "PIEPEM  " & DR("eapercs") & vbNewLine & "PIEPEM2 " & DR("eapercs") & vbNewLine & "CODDUN  " & vbNewLine

                        DTLCP = DTLCP & "COM_CPT_INIMER " & vbNewLine & vbNewLine 'Inicio de mercancia
                        DTLCP = DTLCP & "   COM_CPT_MER_BIENTRA " & ItemCatalog & vbNewLine 'BienesTransp (clave de producto)
                        DTLCP = DTLCP & "   COM_CPT_MER_DESCRI " & ItemDesc & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_MER_CANTID 1" & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_MER_CVUNID  " & UOMCatalog & vbNewLine '
                        DtlWeight = RoundUpToDecimals(DR("csshipped") * DR("grosskg"), 2)
                        If DtlWeight = 0 Then DtlWeight = 0.01
                        TotWeight = TotWeight + DtlWeight
                        DTLCP = DTLCP & "   COM_CPT_MER_PKG " & DtlWeight & vbNewLine & vbNewLine 'itemweight e  PesoEnKg
                        If LiqPresentDtl Then
                            DTLCP = DTLCP & "      COM_CPT_MER_MATPEL Sí  " & vbNewLine 'Material peligroso
                            DTLCP = DTLCP & "      COM_CPT_MER_CVEMATPEL   3065  " & vbNewLine ' BEBIDAS ALCOHOLICAS, 24% pero no más de 70% de alcohol en volumen
                            DTLCP = DTLCP & "      COM_CPT_MER_EMB 4C1" & vbNewLine
                        End If
                        If Peli = "No" Then
                            DTLCP = DTLCP & "      COM_CPT_MER_MATPEL No " & vbNewLine
                        End If

                        DTLCP = DTLCP & "   COM_CPT_INICANTRAN " & vbNewLine 'Inicio canidad trasladada
                        DTLCP = DTLCP & "   COM_CPT_CMER_CANTID " & DR("csshipped") & vbNewLine
                        DTLCP = DTLCP & "      COM_CPT_CMER_IDORI OR000001" & vbNewLine
                        DTLCP = DTLCP & "     COM_CPT_CMER_IDDES DE000001" & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_FINCANTRAN" & vbNewLine
                        DTLCP = DTLCP & "COM_CPT_FINMER" & vbNewLine & vbNewLine

                        '----------------------CHANGES FOR FREE PRODUCT-----------------
                        If CSPrice = 0 Then 'free
                            CSPrice = 0.01
                            DiscExists = True
                            DiscPer = 100
                            DiscAmt = 0.01 * DR("csshipped")
                            IEPSTasa = 0 'In case free item has ieps do not report tasa because it won't have any ieps amount
                            LiqTax2 = 0 'In case free item, do not report any tasa, no ieps amt should be reported
                            TotDiscAmt = TotDiscAmt + DiscAmt
                        End If

                        If RFC <> "PHI830429MG6" Then
                            If Trim(LineIVA) = vbNullString Then LineIVA = "0" : If Trim(IEPSAmt) = vbNullString Then IEPSAmt = "0"
                        End If

                        If Val(CStr(IvaTasaDtl)) > 0 Then
                            DTL = DTL & "TASIPE  " & IvaTasaDtl & vbNewLine & "MONIPE  " & LineIVA & vbNewLine
                        Else
                            DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine
                        End If

                        '--------------------OBJECTO DE IMPUESTO ------------
                        If Val(CStr(IEPSTasa)) > 0 Or LiqTax2 > 0 Then
                            If SeparateLiqTaxFlag = "N" Then
                                If DiscExists Then
                                    DTL = DTL & "OBJIMP 01" & vbNewLine
                                Else
                                    DTL = DTL & "OBJIMP 02" & vbNewLine
                                End If
                            Else
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            End If
                        Else
                            If Val(CStr(IvaTasaDtl)) > 0 Then
                                If SeparateLiqTaxFlag = "N" And Val(CStr(IEPSTasa)) > 0 Then
                                    DTL = DTL & "OBJIMP 02" & vbNewLine
                                Else
                                    DTL = DTL & "OBJIMP 02" & vbNewLine
                                End If
                            Else 'ivatasadtl = 0
                                If DiscExists Then 'IF FREE ITEM USE OBJIMP 01 (NO OBJETO DE IMPUESTO)
                                    DTL = DTL & "OBJIMP 01" & vbNewLine
                                Else
                                    DTL = DTL & "OBJIMP 02" & vbNewLine
                                End If
                                DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine
                                If DiscExists = False Then
                                    DTL = DTL & "FCTTASIPE   0.000000" & vbNewLine
                                    DTL = DTL & "IMPORTIPE " & Format(CSPrice * DR("csshipped"), "##0.00") & vbNewLine
                                    DTL = DTL & "TIPIPETR Tasa" & vbNewLine & "TASIEP 0" & vbNewLine
                                End If
                            End If
                        End If
                        If IEPSTasa = 8 Then '8% ieps but no iva, detail should show 0% IVA
                            DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine & "FCTTASIPE   0.000000" & vbNewLine
                            DTL = DTL & "IMPORTIPE " & Format((CSPrice * DR("csshipped")) + CDbl(IEPSAmt), "##0.00") & vbNewLine
                            DTL = DTL & "TIPIPETR Tasa" & vbNewLine
                            TotAmtNotTaxable = TotAmtNotTaxable + CDbl(IEPSAmt)
                        End If

                        DTL = DTL & "MONIEP_IEPS  " & IEPSAmt_Prt & vbNewLine
                        DTL = DTL & "IMPIVAIEPS  " & CDbl(DR("linetotal")) & vbNewLine
                        DTL = DTL & "PBRUDE_IEPS  " & UnitPrice & vbNewLine
                        DTL = DTL & "IMPBRU_PRT  " & Importe & vbNewLine

                        '********************************************SIR 1794********************************************************
                        If ((SeparateLiqTaxFlag = "N" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "N" And LiqTax2 > 0)) And IEPS_NO_SeparateXml Then
                            DTL = DTL & "PBRUDE  " & Format(CDbl(UnitPrice), "##0.00") & vbNewLine
                            DTL = DTL & "VALUNI  " & Format(CDbl(UnitPrice), "##0.00") & vbNewLine
                            DTL = DTL & "IMPBRU  " & Format(CDbl(UnitPrice) * DR("csshipped"), "##0.00") & vbNewLine
                            DTL = DTL & "IMPORT  " & Format(CDbl(UnitPrice) * DR("csshipped"), "##0.00") & vbNewLine
                        Else
                            If Val(CStr(IEPSTasa)) > 0 Then DTL = DTL & "TASIEP  " & IEPSTasa & vbNewLine & "MONIEP  " & CDbl(IEPSAmt) & vbNewLine
                            DTL = DTL & "PBRUDE  " & Format(CSPrice, "##0.00") & vbNewLine
                            DTL = DTL & "IMPBRU  " & Format(CSPrice * DR("csshipped"), "##0.00") & vbNewLine
                            DTL = DTL & "VALUNI  " & Format(CSPrice, "##0.00") & vbNewLine
                            DTL = DTL & "IMPORT  " & Format(CSPrice * DR("csshipped"), "##0.00") & vbNewLine
                        End If

                        TotIEPSAmt = TotIEPSAmt + Val(IEPSAmt)
                        If LiqTax2 > 0 Then '---------SUGARY DRINKS!!!!!!!!!!  1.17 per liter ------------------
                            If DiscAmt = 0 And SeparateLiqTaxFlag = "Y" Then
                                DTL = DTL & "FCTTASIEP   " & IEPSperLiter & vbNewLine
                                DTL = DTL & "TIPIEPTR    Cuota" & vbNewLine
                                DTL = DTL & "IMPORTIEP    " & Format(Val(IEPSAmt) / IEPSperLiter, "0.##") & vbNewLine
                                DTL = DTL & "MONIEP " & CDbl(Val(IEPSAmt)) & vbNewLine

                                If Val(CStr(IvaTasaDtl)) = 0 Then
                                    DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine & "FCTTASIPE   0.000000" & vbNewLine
                                    DTL = DTL & "IMPORTIPE " & Format(CSPrice * DR("csshipped") + CDbl(IEPSAmt), "##0.00") & vbNewLine 'Base to calculate Tasa 0%
                                    DTL = DTL & "TIPIPETR Tasa" & vbNewLine & "TASIEP 0" & vbNewLine
                                    IEPSnoIVA = IEPSnoIVA + CDbl(IEPSAmt)
                                End If
                            End If
                        End If
                        If DiscExists Then
                            DTL = DTL & "TDECON    " & DiscPer & vbNewLine
                            DTL = DTL & "MDECON    " & (CSPrice * DR("csshipped")) * (DiscPer / 100) & vbNewLine
                        Else
                            DTL = DTL & "TDECON  " & "0" & vbNewLine & "MDECON  " & "0.00" & vbNewLine 'NO DESCUENTO EN DETALLE HARD CODE 0
                        End If

                        DTL = DTL & "NUMLIN  " & d & vbNewLine
                    Case Is = "CE"
                        '---------------------LiqTax2 = 4.69-------------------------'
                        CsImpIvaIeps = 0 : EaImpIvaIeps = 0
                        'PRICEPLUSLIQTAX USED TO CALCULATE IVA BECAUSE IVA INCLUDES LIQ TAX ON TOP OF PRICE.
                        PricePlusLiqTax = Val(DR("INVCSPRICE")) + (Val(DR("INVCSPRICE")) * (IEPSTasa / 100))
                        If LiqTax2 > 0 Then PricePlusLiqTax = Val(DR("INVCSPRICE")) + LiqTax2 '---12/26/13
                        'this line is done here because we will separate cases line from each line. Because this is the 1st line, needs to have everything
                        UnitPrice = Format(CSPrice, "##0.00")
                        Importe = Format(CDbl(UnitPrice) * DR("csshipped"), "##0.00###")
                        TotImp = TotImp + Val(Importe)
                        If Val(CStr(IEPSTasa)) > 0 Then IEPSAmt = Format((Val(CStr(CSPrice)) * DR("csshipped")) * (IEPSTasa / 100), "##0.00###") 'SAT IEPSAmt
                        If LiqTax2 > 0 Then IEPSAmt = Format(DR("csshipped") * LiqTax2, "##0.00")
                        If (SeparateLiqTaxFlag = "Y" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "Y" And LiqTax2 > 0) Then
                            IEPSAmt_Prt = IEPSAmt
                            TotIEPS_Prt = CStr(Val(TotIEPS_Prt) + Val(IEPSAmt))
                        End If
                        If (SeparateLiqTaxFlag = "N" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "N" And LiqTax2 > 0) Then 'Price should add IEPS on inv
                            If Val(CStr(IEPSTasa)) > 0 Then UnitPrice = Format(DR("INVCSPRICE") + (DR("INVCSPRICE") * (IEPSTasa / 100)), "##0.00###")
                            If LiqTax2 > 0 Then UnitPrice = Format(DR("INVCSPRICE") + LiqTax2, "##0.00###")
                            Importe = Format(CDbl(UnitPrice) * DR("csshipped"), "##0.00###")
                            CsImporte = CDbl(Importe) 'ONLY USED WHEN CS & EA FOR LIQUOR NO DESGLOSE CALCULATION... ON EACH
                            IEPSAmt_Prt = vbNullString
                        End If
                        TotImp_Prt = TotImp_Prt + Val(Importe)
                        If Val(CStr(IvaTasaDtl)) > 0 Then
                            IvaAmt = CStr(PricePlusLiqTax * DR("csshipped") * (IVATasa / 100))
                            IvaAmt = RoundUpToDecimals(IvaAmt, 2)
                            TotalTax_New = TotalTax_New + CDbl(IvaAmt)
                        End If
                        CsImpIvaIeps = (Val(CStr(CSPrice)) * DR("csshipped")) + Val(IvaAmt) + Val(IEPSAmt)
                        Units = "CA"
                        UOMCatalog = "XBX"
                        DTL = DTL & "CANTID  " & DR("csshipped") & vbNewLine & "CANTID_CA  " & DR("csshipped") & vbNewLine & "DESCRI  " & ItemDesc.ToString.Trim & vbNewLine
                        DTL = DTL & "CANPAQ  " & DR("csshipped") & vbNewLine
                        DTL = DTL & "CANEMP  " & DR("csshipped") & vbNewLine & "UNIDAD  " & Units & vbNewLine & "CVESKU  " & DR("itemcode") & vbNewLine
                        DTL = DTL & "ESTILV     " & DR("itemcode") & vbNewLine
                        DTL = DTL & "CVEPRODSERV     " & ItemCatalog & vbNewLine
                        DTL = DTL & "CVEUNIDAD     " & UOMCatalog & vbNewLine
                        DTL = DTL & "CODUPC  " & UPC & vbNewLine & "PIEPEM  " & DR("eapercs") & vbNewLine & "PIEPEM2 " & DR("eapercs") & vbNewLine & "CODDUN  " & vbNewLine
                        DTLCP = DTLCP & "COM_CPT_INIMER " & vbNewLine & vbNewLine 'Inicio de mercancia
                        DTLCP = DTLCP & "   COM_CPT_MER_BIENTRA " & ItemCatalog & vbNewLine 'BienesTransp (clave de producto)
                        DTLCP = DTLCP & "   COM_CPT_MER_DESCRI " & ItemDesc & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_MER_CANTID 1" & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_MER_CVUNID  " & UOMCatalog & vbNewLine '

                        DtlWeight = RoundUpToDecimals(DR("csshipped") * DR("grosskg"), 2)
                        If DtlWeight = 0 Then DtlWeight = 0.01
                        TotWeight = TotWeight + DtlWeight
                        DTLCP = DTLCP & "   COM_CPT_MER_PKG " & DtlWeight & vbNewLine & vbNewLine 'itemweight  PesoEnKg
                        If LiqPresentDtl Then
                            DTLCP = DTLCP & "      COM_CPT_MER_MATPEL Sí  " & vbNewLine 'Material peligroso
                            DTLCP = DTLCP & "      COM_CPT_MER_CVEMATPEL   3065  " & vbNewLine ' BEBIDAS ALCOHOLICAS, 24% pero no más de 70% de alcohol en volumen
                            DTLCP = DTLCP & "      COM_CPT_MER_EMB 4C1" & vbNewLine
                        End If
                        If Peli = "No" Then
                            DTLCP = DTLCP & "      COM_CPT_MER_MATPEL No " & vbNewLine
                        End If
                        DTLCP = DTLCP & "   COM_CPT_INICANTRAN " & vbNewLine 'Inicio canidad trasladada
                        DTLCP = DTLCP & "   COM_CPT_CMER_CANTID " & DR("csshipped") & vbNewLine
                        DTLCP = DTLCP & "      COM_CPT_CMER_IDORI OR000001" & vbNewLine
                        DTLCP = DTLCP & "     COM_CPT_CMER_IDDES DE000001" & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_FINCANTRAN" & vbNewLine
                        DTLCP = DTLCP & "COM_CPT_FINMER" & vbNewLine & vbNewLine

                        If CSPrice = 0 Then 'FREE CASES FIRST (THIS IS WHEN CS AND EA PRESENT)
                            CSPrice = 0.01
                            DiscExists = True
                            DiscPer = 100
                            DiscAmt = 0.01 * DR("csshipped")
                            IEPSTasa = 0 'In case free item has ieps do not report tasa because it won't have any ieps amount
                            LiqTax2 = 0 'In case free item, do not report any tasa, no ieps amt should be reported
                            TotDiscAmt = TotDiscAmt + DiscAmt
                        End If
                        If RFC <> "PHI830429MG6" Then
                            If Trim(IvaAmt) = vbNullString Then IvaAmt = "0" : If Trim(IEPSAmt) = vbNullString Then IEPSAmt = "0"
                        End If

                        DTL = DTL & "TASIPE  " & IvaTasaDtl & vbNewLine
                        DTL = DTL & "MONIPE  " & IvaAmt & vbNewLine

                        '--------------------OBJECTO DE IMPUESTO --------------
                        If Val(CStr(IEPSTasa)) > 0 Or LiqTax2 > 0 Then
                            If SeparateLiqTaxFlag = "N" Then
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            Else
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            End If
                        Else
                            If Val(CStr(IvaTasaDtl)) > 0 Then
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            Else
                                If DiscExists Then
                                    DTL = DTL & "OBJIMP 01" & vbNewLine
                                    DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine
                                Else
                                    DTL = DTL & "OBJIMP 02" & vbNewLine
                                    DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine & "FCTTASIPE   0.000000" & vbNewLine
                                    DTL = DTL & "IMPORTIPE " & Format(CSPrice * DR("csshipped"), "##0.00") & vbNewLine
                                    DTL = DTL & "TIPIPETR Tasa" & vbNewLine & "TASIEP 0" & vbNewLine
                                End If
                            End If
                        End If
                        If IEPSTasa = 8 Then '8% ieps but no iva, detail should show 0% IVA
                            DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine & "FCTTASIPE   0.000000" & vbNewLine
                            DTL = DTL & "IMPORTIPE " & Format((CSPrice * DR("csshipped")) + CDbl(IEPSAmt), "##0.00") & vbNewLine
                            DTL = DTL & "TIPIPETR Tasa" & vbNewLine
                            TotAmtNotTaxable = TotAmtNotTaxable + CDbl(IEPSAmt)
                        End If

                        DTL = DTL & "MONIEP_IEPS  " & IEPSAmt_Prt & vbNewLine
                        DTL = DTL & "IMPIVAIEPS  " & CsImpIvaIeps & vbNewLine
                        DTL = DTL & "PBRUDE_IEPS  " & UnitPrice & vbNewLine
                        DTL = DTL & "IMPBRU_PRT  " & CDbl(Importe) & vbNewLine

                        '********************************************SIR 1794********************************************************
                        If ((SeparateLiqTaxFlag = "N" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "N" And LiqTax2 > 0)) And IEPS_NO_SeparateXml Then
                            DTL = DTL & "PBRUDE  " & Format(CDbl(UnitPrice), "##0.00###") & vbNewLine
                            DTL = DTL & "VALUNI  " & Format(CDbl(UnitPrice), "##0.00###") & vbNewLine
                            DTL = DTL & "IMPBRU  " & Format(CDbl(UnitPrice) * DR("csshipped"), "##0.00###") & vbNewLine
                            DTL = DTL & "IMPORT  " & Format(CDbl(UnitPrice) * DR("csshipped"), "##0.00###") & vbNewLine
                        Else
                            If Val(CStr(IEPSTasa)) > 0 Then DTL = DTL & "TASIEP  " & IEPSTasa & vbNewLine & "MONIEP  " & CDbl(IEPSAmt) & vbNewLine
                            DTL = DTL & "PBRUDE  " & Format(CSPrice, "##0.00") & vbNewLine
                            DTL = DTL & "IMPBRU  " & Format(CSPrice * DR("csshipped"), "##0.00") & vbNewLine
                            DTL = DTL & "VALUNI  " & Format(CSPrice, "##0.00") & vbNewLine
                            DTL = DTL & "IMPORT  " & Format(CSPrice * DR("csshipped"), "##0.00") & vbNewLine
                        End If

                        If DiscExists Then
                            DTL = DTL & "TDECON    " & DiscPer & vbNewLine
                            DTL = DTL & "MDECON    " & (CSPrice * DR("csshipped")) * (DiscPer / 100) & vbNewLine
                        Else
                            DTL = DTL & "TDECON  " & "0" & vbNewLine & "MDECON  " & "0.00" & vbNewLine 'NO DESCUENTO EN DETALLE HARD CODE 0
                        End If

                        'addenda does not get created because when ea and cs present, I did not have the NUMLIN for this 1stline created (CA), only for EA.
                        DTL = DTL & "NUMADU  " & DocID & vbNewLine & "FECADU  " & DocDate & vbNewLine & "ADUANA  " & PortName & vbNewLine & "EANADU  " & vbNewLine
                        DTL = DTL & "NUMPED " & Trim(DocID) & vbNewLine

                        If LiqTax2 > 0 Then '---------SUGARY DRINKS--------------
                            If DiscAmt = 0 And SeparateLiqTaxFlag = "Y" Then ' don't do this bellow for free items --07/14/23  ADDED SEPARATELIQTAXFLAG
                                DTL = DTL & "FCTTASIEP   " & IEPSperLiter & vbNewLine
                                DTL = DTL & "TIPIEPTR    Cuota" & vbNewLine
                                DTL = DTL & "IMPORTIEP    " & Format(Val(IEPSAmt) / IEPSperLiter, "0.##") & vbNewLine
                                DTL = DTL & "MONIEP " & CDbl(Val(IEPSAmt)) & vbNewLine
                                If Val(CStr(IvaTasaDtl)) = 0 Then
                                    DTL = DTL & "TASIPE 0 " & vbNewLine & "MONIPE 0" & vbNewLine & "FCTTASIPE 0.000000" & vbNewLine
                                    DTL = DTL & "TIPIPETR Tasa " & vbNewLine & "IMPORTIPE " & ((CSPrice * DR("csshipped")) + CDbl(IEPSAmt)) & vbNewLine
                                    IEPSnoIVA = IEPSnoIVA + CDbl(IEPSAmt)
                                End If
                            End If
                        End If
                        DTL = DTL & "NUMLIN  " & d & vbNewLine
                        DTL = DTL & "D " & vbNewLine
                        TotIEPSAmt = TotIEPSAmt + Val(IEPSAmt)

                        d = d + 1
                        PricePlusLiqTax = Val(DR("INVEAPRICE"))
                        If IEPSTasa > 0 Then PricePlusLiqTax = Val(DR("INVEAPRICE")) + (Val(DR("INVEAPRICE")) * (IEPSTasa / 100))
                        If LiqTax2 > 0 Then PricePlusLiqTax = Val(DR("INVEAPRICE")) + (LiqTax2 / DR("eapercs"))

                        UnitPrice = Format(Eaprice, "##0.00")
                        Importe = Format(CDbl(UnitPrice) * DR("eashipped"), "##0.00###")
                        TotImp = TotImp + Val(Importe)
                        IEPSAmt = vbNullString
                        If Val(CStr(IEPSTasa)) > 0 Then IEPSAmt = Format((Val(CStr(Eaprice)) * DR("eashipped")) * (IEPSTasa / 100), "##0.00###") 'SAT IEPSAmt
                        If LiqTax2 > 0 Then IEPSAmt = Format(DR("eashipped") * (LiqTax2 / DR("eapercs")), "##0.00###")
                        If SeparateLiqTaxFlag = "N" Then
                            If IEPSTasa > 0 Then UnitPrice = Format(DR("INVEAPRICE") + (DR("INVEAPRICE") * (IEPSTasa / 100)), "##0.00###")
                            If LiqTax2 > 0 Then UnitPrice = Format(Eaprice + (LiqTax2 / DR("eapercs")), "##0.00###")
                            If IEPSTasa = 0 And LiqTax2 = 0 Then
                                Importe = Format(CDbl(UnitPrice) * DR("eashipped"), "##0.00###")
                            Else
                                Importe = Format(Val(DR("lineprice")) + Val(DR("lineliqtax")) - CsImporte, "##0.00###")
                            End If
                        End If
                        If (SeparateLiqTaxFlag = "Y" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "Y" And LiqTax2 > 0) Then
                            IEPSAmt_Prt = IEPSAmt
                            TotIEPS_Prt = CStr(Val(TotIEPS_Prt) + Val(IEPSAmt))
                        End If

                        TotImp_Prt = TotImp_Prt + Val(Importe)
                        If SeparateLiqTaxFlag = "Y" And LiqTax2 > 0 Then IEPSAmt = Format((LiqTax2 / DR("eapercs")) * (DR("eashipped")), "##0.00")
                        If Val(CStr(IvaTasaDtl)) > 0 Then
                            IvaAmt = CStr(PricePlusLiqTax * DR("eashipped") * (IVATasa / 100))

                            IvaAmt = RoundUpToDecimals(IvaAmt, 2)
                            TotalTax_New = TotalTax_New + CDbl(IvaAmt)
                        End If
                        EaImpIvaIeps = (Val(CStr(Eaprice)) * DR("eashipped")) + Val(IvaAmt) + Val(IEPSAmt)
                        Units = "EA"
                        UOMCatalog = "H87"

                        '--------------------OBJECTO DE IMPUESTO WHEN EA AND CS PRESENT CS SHOULD HAVE THEIR OBJIMP ALREADY------------
                        If Val(CStr(IEPSTasa)) > 0 Or LiqTax2 > 0 Then
                            If SeparateLiqTaxFlag = "N" Then
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            Else
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            End If
                        Else
                            If Val(CStr(IvaTasaDtl)) > 0 Then
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            Else
                                If DiscExists Then
                                    DTL = DTL & "OBJIMP 01" & vbNewLine
                                    DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine
                                Else
                                    DTL = DTL & "OBJIMP 02" & vbNewLine
                                    DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine & "FCTTASIPE   0.000000" & vbNewLine
                                    DTL = DTL & "IMPORTIPE " & Format(CDbl(UnitPrice) * DR("eashipped"), "##0.00") & vbNewLine
                                    DTL = DTL & "TIPIPETR Tasa" & vbNewLine & "TASIEP 0" & vbNewLine
                                End If
                            End If
                        End If
                        If IEPSTasa = 8 Then '8% ieps but no iva, detail should show 0% IVA
                            DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine & "FCTTASIPE   0.000000" & vbNewLine
                            DTL = DTL & "IMPORTIPE " & Format(CDbl(UnitPrice) * DR("eashipped"), "##0.00") & vbNewLine
                            DTL = DTL & "TIPIPETR Tasa" & vbNewLine '---------------- & "TASIEP 0" & vbNewLine
                            TotAmtNotTaxable = TotAmtNotTaxable + CDbl(IEPSAmt)
                        End If

                        DTL = DTL & "CANTID  " & DR("eashipped") & vbNewLine & "CANTID_EA  " & DR("eashipped") & vbNewLine & "DESCRI  " & ItemDesc.Trim & vbNewLine
                        DTL = DTL & "CANPAQ  " & DR("eashipped") & vbNewLine
                        DTL = DTL & "CANEMP  " & DR("eashipped") & vbNewLine & "UNIDAD  " & Units & vbNewLine & "CVESKU  " & DR("itemcode") & vbNewLine
                        DTL = DTL & "ESTILV     " & DR("itemcode") & vbNewLine
                        DTL = DTL & "CVEPRODSERV     " & ItemCatalog & vbNewLine
                        DTL = DTL & "CVEUNIDAD     " & UOMCatalog & vbNewLine
                        DTL = DTL & "CODUPC  " & UPC & vbNewLine & "PIEPEM  " & "1" & vbNewLine & "PIEPEM2 " & DR("eapercs") & vbNewLine & "CODDUN  " & vbNewLine
                        DTLCP = DTLCP & "COM_CPT_INIMER " & vbNewLine & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_MER_BIENTRA " & ItemCatalog & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_MER_DESCRI " & ItemDesc.Trim & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_MER_CANTID 1" & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_MER_CVUNID  " & UOMCatalog & vbNewLine '

                        DtlWeight = RoundUpToDecimals((DR("eashipped") / DR("eapercs")) * DR("grosskg"), 2)
                        If DtlWeight = 0 Then DtlWeight = 0.01
                        TotWeight = TotWeight + DtlWeight
                        DTLCP = DTLCP & "   COM_CPT_MER_PKG " & DtlWeight & vbNewLine & vbNewLine 'itemweight e PesoEnKg
                        If LiqPresentDtl Then
                            DTLCP = DTLCP & "      COM_CPT_MER_MATPEL Sí  " & vbNewLine 'Material peligroso
                            DTLCP = DTLCP & "      COM_CPT_MER_CVEMATPEL   3065  " & vbNewLine ' BEBIDAS ALCOHOLICAS, 24% pero no más de 70% de alcohol en volumen
                            DTLCP = DTLCP & "      COM_CPT_MER_EMB 4C1" & vbNewLine
                        End If
                        If Peli = "No" Then
                            DTLCP = DTLCP & "      COM_CPT_MER_MATPEL No " & vbNewLine
                        End If
                        DTLCP = DTLCP & "   COM_CPT_INICANTRAN " & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_CMER_CANTID " & DR("eashipped") & vbNewLine
                        DTLCP = DTLCP & "      COM_CPT_CMER_IDORI OR000001" & vbNewLine
                        DTLCP = DTLCP & "     COM_CPT_CMER_IDDES DE000001" & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_FINCANTRAN" & vbNewLine
                        DTLCP = DTLCP & "COM_CPT_FINMER" & vbNewLine & vbNewLine

                        If Eaprice = 0 Then 'FREE EA (WHEN CS AND EA PRESENT) 
                            Eaprice = 0.01
                            DiscExists = True
                            DiscPer = 100
                            DiscAmt = 0.01 * DR("eashipped")
                            IEPSTasa = 0 'In case free item has ieps do not report tasa because it won't have any ieps amount
                            LiqTax2 = 0 'In case free item, do not report any tasa, no ieps amt should be reported
                            TotDiscAmt = TotDiscAmt + DiscAmt
                        End If

                        If RFC <> "PHI830429MG6" Then
                            If Trim(IvaAmt) = vbNullString Then IvaAmt = "0" : If Trim(IEPSAmt) = vbNullString Then IEPSAmt = "0"
                        End If

                        DTL = DTL & "TASIPE  " & IvaTasaDtl & vbNewLine & "MONIPE  " & IvaAmt & vbNewLine

                        '--------------------OBJECTO DE IMPUESTO ------------
                        If Val(CStr(IEPSTasa)) > 0 Or LiqTax2 > 0 Then
                            If SeparateLiqTaxFlag = "N" Then
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            Else
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            End If
                        Else
                            If Val(CStr(IvaTasaDtl)) > 0 Then
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            Else
                                If DiscExists Then
                                    DTL = DTL & "OBJIMP 01" & vbNewLine
                                    DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine
                                Else
                                    DTL = DTL & "OBJIMP 02" & vbNewLine
                                    DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine & "FCTTASIPE   0.000000" & vbNewLine
                                    DTL = DTL & "IMPORTIPE " & Format(Eaprice * DR("eashipped"), "##0.00") & vbNewLine
                                    DTL = DTL & "TIPIPETR Tasa" & vbNewLine & "TASIEP 0" & vbNewLine
                                End If
                            End If
                        End If
                        If IEPSTasa = 8 Then '8% ieps but no iva, detail should show 0% IVA
                            DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine & "FCTTASIPE   0.000000" & vbNewLine
                            DTL = DTL & "IMPORTIPE " & Format((Eaprice * DR("eashipped")) + CDbl(IEPSAmt), "##0.00") & vbNewLine
                            DTL = DTL & "TIPIPETR Tasa" & vbNewLine
                            TotAmtNotTaxable = TotAmtNotTaxable + CDbl(IEPSAmt)
                        End If

                        DTL = DTL & "MONIEP_IEPS  " & IEPSAmt_Prt & vbNewLine
                        DTL = DTL & "IMPIVAIEPS  " & EaImpIvaIeps & vbNewLine
                        DTL = DTL & "PBRUDE_IEPS  " & UnitPrice & vbNewLine
                        DTL = DTL & "IMPBRU_PRT  " & CDbl(Importe) & vbNewLine

                        If ((SeparateLiqTaxFlag = "N" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "N" And LiqTax2 > 0)) And IEPS_NO_SeparateXml Then
                            DTL = DTL & "PBRUDE  " & Format(CDbl(UnitPrice), "##0.00###") & vbNewLine
                            DTL = DTL & "VALUNI  " & Format(CDbl(UnitPrice), "##0.00###") & vbNewLine
                            DTL = DTL & "IMPBRU  " & Format(CDbl(UnitPrice) * DR("eashipped"), "##0.00###") & vbNewLine
                            DTL = DTL & "IMPORT  " & Format(CDbl(UnitPrice) * DR("eashipped"), "##0.00###") & vbNewLine
                        Else
                            If Val(CStr(IEPSTasa)) > 0 Then DTL = DTL & "TASIEP  " & IEPSTasa & vbNewLine & "MONIEP  " & CDbl(IEPSAmt) & vbNewLine
                            DTL = DTL & "PBRUDE  " & Format(Eaprice, "##0.00") & vbNewLine
                            DTL = DTL & "VALUNI  " & Format(Eaprice, "##0.00") & vbNewLine
                            DTL = DTL & "IMPBRU  " & Format(Eaprice * DR("eashipped"), "##0.00") & vbNewLine
                            DTL = DTL & "IMPORT  " & Format(Eaprice * DR("eashipped"), "##0.00") & vbNewLine
                        End If

                        If DiscExists Then
                            DTL = DTL & "TDECON    " & DiscPer & vbNewLine
                            DTL = DTL & "MDECON    " & (Eaprice * DR("eashipped")) * (DiscPer / 100) & vbNewLine
                        Else
                            DTL = DTL & "TDECON  " & "0" & vbNewLine & "MDECON  " & "0.00" & vbNewLine
                        End If
                        If LiqTax2 > 0 Then '0---------SUGARY DRINKS----------
                            If DiscAmt = 0 And SeparateLiqTaxFlag = "Y" Then
                                DTL = DTL & "FCTTASIEP   " & IEPSperLiter & vbNewLine
                                DTL = DTL & "TIPIEPTR    Cuota" & vbNewLine
                                DTL = DTL & "IMPORTIEP    " & Format(Val(IEPSAmt) / IEPSperLiter, "0.##") & vbNewLine
                                DTL = DTL & "MONIEP " & CDbl(Val(IEPSAmt)) & vbNewLine
                                DTL = DTL & "MONIEP " & CDbl(Val(IEPSAmt)) & vbNewLine
                                If Val(CStr(IvaTasaDtl)) = 0 Then
                                    DTL = DTL & "TASIPE 0 " & vbNewLine & "MONIPE 0" & vbNewLine & "FCTTASIPE 0.000000" & vbNewLine
                                    DTL = DTL & "TIPIPETR Tasa " & vbNewLine & "IMPORTIPE " & ((Eaprice * DR("eashipped")) + CDbl(IEPSAmt)) & vbNewLine
                                    IEPSnoIVA = IEPSnoIVA + CDbl(IEPSAmt)
                                End If
                            End If
                        End If
                        TotIEPSAmt = TotIEPSAmt + Val(IEPSAmt)
                End Select
                SQL = "select * from custom_doc where itemcode = '" & DR("itemcode") & "'"

                OCM = New OracleCommand(SQL, conn)
                dtc = New DataTable
                dtc.Load(OCM.ExecuteReader)
                DocID = vbNullString : DocDate = vbNullString : PortName = vbNullString
                If dtc.Rows.Count > 0 Then
                    DRc = dtc.Rows(0)
                    PortName = vbNullString & DRc("Port")
                    FirstPortionOfName = InStr(PortName, ",")
                    If FirstPortionOfName > 0 Then PortName = Left(PortName, FirstPortionOfName - 1)
                    PortName = Left(PortName, 11) 'restrict port name to a max of 11 per e-mail
                    DocID = StripChars(vbNullString & DRc("doc_id"))

                    If Len(Trim(DocID)) <> 15 Then DocID = vbNullString

                    DocDate = Format(DRc("doc_date"), "yyyy-MM-dd")

                    If Not SorianaErr Then
                        If Trim(DocID) <> vbNullString Then
                            SQLInsert = "insert into custom_doc_invoice values ('" & InvNum & "','" & DR("itemcode") & "','" & DocID & "','" & Format(DRc("doc_date"), "dd-MMM-yyyy") & "','" & PortName & "')"
                            OCM = New OracleCommand(SQLInsert, conn)
                            OCM.ExecuteNonQuery()
                        End If
                    End If
                End If

                DTL = DTL & "NUMADU  " & DocID & vbNewLine & "FECADU  " & DocDate & vbNewLine & "ADUANA  " & PortName & vbNewLine & "EANADU  " & vbNewLine
                DTL = DTL & "NUMPED " & Trim(DocID) & vbNewLine
                DTL = DTL & "NUMLIN  " & d & vbNewLine
            Next
            DTLCP = DTLCP & "COM_CPT_FINMER" & vbNewLine & "COM_CPT_FINMERS " & vbNewLine
            DTLCPT = DTLCPT & "COM_CPT_INIMERS " & vbNewLine
            DTLCPT = DTLCPT & "COM_CPT_MER_PESBRU " & Format(TotWeight, "##0.00") & vbNewLine ''PESO BRUTO 
            DTLCPT = DTLCPT & " COM_CPT_MER_PESONET " & Format(TotWeight, "##.00") & vbNewLine ' 'PESO NETO, using peso bruto as not all items have peso neto.
            DTLCPT = DTLCPT & "COM_CPT_MER_UNIPES KGM" & vbNewLine
            DTLCPT = DTLCPT & "COM_CPT_MER_NUMTOT  " & d & vbNewLine & vbNewLine
        End If
        'NO ITEM CODE SO JFCITEM AND BRANCHITEM PRODUCE NOTHING, NO CS, no EA.. some times we get only remark code no item....  ie:  39-0128122,39-0210747, specially notas de cargo
        'OR IT MIGHT BE ITEMCODE PRESENT BUT NO EA NOR CS SHIPPED, MIGHT BE ADDITIONAL CHARGE (ERROR ON PRICE OR SOMETHING ELSE) ie: 39-0356084
        NoItemCdRoutine(InvNum, IEPSTasa, IEPSAmt)
        Exit Sub
ErrHndlr:
        ErrMsgLog = New String("*", 50) & vbNewLine & Format(Now, "MM/dd/yy HH:mm") & " -> ERROR:  " & Err.Number & "-->" & Err.Description & "." & vbNewLine & "Error on GetDtl Routine.  Prog:  DigInv3." & vbNewLine & "Total Invoices:  " & TotInvs & vbNewLine & vbNewLine & "Last SQL ran : " & vbNewLine & SQL & vbNewLine & "Item being processed for line: " & d & " " & ItemCode & vbNewLine & "Order being processed: " & InvNum & vbNewLine
        CkFileExists(DirToOutputError, ErrMsgLog, "ERR-Dig-Inv3_3.log")
        SendEmail(ErrMsgLog)
        End
    End Sub

    '***********************************************************************'
    'Name: GetRelUuid
    'Description: Retrieves the UUID (universal unique identifier) based on invoice number
    'Params: 
    '   - RefInvNo : The invoice number used to retrieve the UUID 
    'Return Value: String value representing the UUID tied to the invoice number passed in
    'Precondition(s): N/A
    'Postcondition(s): N/A
    '***********************************************************************'
    Private Function GetRelUuid(ByRef RefInvNo As String) As String

        Dim SQL As String
        Dim OCM As OracleCommand
        Dim OTR As OracleTransaction
        Dim DT As DataTable
        On Error GoTo ErrHndlr
        SQL = "select r." & Chr(34) & "rep_UUID" & Chr(34) & " UUID " & "from dbo.reporte@jmxinv r " & "Where r." & Chr(34) & "rep_folio" & Chr(34) & " = '" & RefInvNo & "' "
        OCM = New OracleCommand(SQL, conn)
        DT = New DataTable
        DT.Load(OCM.ExecuteReader)

        If DT.Rows.Count > 0 Then
            For Each DR As DataRow In DT.Rows
                GetRelUuid = DR("UUID")
            Next
        Else
            GetRelUuid = vbNullString
        End If
        Exit Function
ErrHndlr:
        ErrMsgLog = New String("*", 50) & vbNewLine & Format(Now, "MM/dd/yy HH:mm") & " -> ERROR:  " & Err.Number & "-->" & Err.Description & "." & vbNewLine & "Error on GetDtl Routine.  Prog:  DigInv3." & vbNewLine & "Total Invoices:  " & TotInvs & vbNewLine & vbNewLine & "Last SQL ran : " & vbNewLine & SQL & vbNewLine
        CkFileExists(DirToOutputError, ErrMsgLog, "ERR-Dig-Inv3_3.log")
        SendEmail(ErrMsgLog)
    End Function

    '***********************************************************************'
    'Name: NoItemCdRoutine
    'Description: Adds details for items w/o item code or no EA, CS Shipped 
    'Params: 
    '   - InvNum : Invoice number being processed 
    '   - IEPSTasa : The tax rate (Tasa = tax rate) 
    '   - IEPSAmt : The tax amount 
    'Return Value: N/A
    'Precondition(s): N/A
    'Postcondition(s): Details for items w/o item code have been appended to global variable DTL
    '***********************************************************************'
    Private Sub NoItemCdRoutine(ByRef InvNum As String, ByRef IEPSTasa As Double, ByRef IEPSAmt As String)

        Dim Rsd2, Rsd3 As Object
        Dim Created1stPart, Created2ndPart As Boolean
        Dim IEPSAmt_Prt, SQL, TmpTotImp As String
        Dim j As Short
        Dim LineImpIvaIeps As Double
        Dim UnitPrice As Double
        Dim FinishedDtl As Boolean

        Dim OCM As OracleCommand
        Dim ODR As OracleDataReader
        Dim dtd2, dtd3 As DataTable
        Dim DRd3 As DataRow
        Dim Adpt As OracleDataAdapter

        On Error GoTo ErrorHandler

        SQL = "SELECT * FROM INVDTL D where invhdrnum = '" & InvNum & "' and (itemcode is null or (csshipped = 0 and eashipped = 0 and linetotal > 0) ) ORDER BY LINENUM"
        OCM = New OracleCommand(SQL, conn)
        Adpt = New OracleDataAdapter(OCM)
        dtd2 = New DataTable
        Adpt.Fill(dtd2)
        If dtd2.Rows.Count > 0 Then
            j = 1
            For Each DR2 As DataRow In dtd2.Rows
                SeparateLiqTaxFlag = IepsSeparate(InvNum, DR2("linenum"))
                Select Case (vbNullString & DR2("remarkid"))
                    Case Is = "MC", "AC" 'Nota de Cargo, Aditional charge
                        If j > 1 Then DTL = DTL & "IMPIVAIEPS  " & LineImpIvaIeps & vbNewLine & "NUMLIN  " & TotLines & vbNewLine 'if more than one line of nota de cargo, finish the dtl line, otherwise, it will start wiht 2nd line without finishing 1st.
                        Create1stPartofDtl(DR2)
                        Created1stPart = True : Created2ndPart = True 'if nota de cargo, and iva, ieps, NC will create the 2nd part of dtl, so we make flag = T.
                        TotImp_Prt = TotImp_Prt + DR2("lineprice")
                        TotAmtToBeTaxed = DR2("lineprice")
                        UnitPrice = DR2("lineprice")
                        LineImpIvaIeps = DR2("lineprice")
                        DTL = DTL & "CODUPC  " & "001234567890" & vbNewLine
                        DTL = DTL & "VALUNI  " & DR2("lineprice") & vbNewLine & "IMPORT  " & DR2("lineprice") & vbNewLine
                        DTL = DTL & "PBRUDE  " & DR2("lineprice") & vbNewLine
                        DTL = DTL & "IMPBRU  " & DR2("lineprice") & vbNewLine
                        DTL = DTL & "CVEPRODSERV " & "84111506" & vbNewLine
                        DTL = DTL & "CVEUNIDAD " & "ACT" & vbNewLine
                        If SeparateLiqTaxFlag = "Y" Then
                            DTL = DTL & "PBRUDE_IEPS  " & DR2("lineprice") & vbNewLine
                            DTL = DTL & "IMPBRU_PRT  " & CDbl(DR2("lineprice")) & vbNewLine
                        End If
                    Case Is = "FC"
                        If j > 1 Then DTL = DTL & "IMPIVAIEPS  " & LineImpIvaIeps & vbNewLine & "NUMLIN  " & TotLines & vbNewLine '--if more than one line of nota de cargo, finish the dtl line, otherwise, it will start wiht 2nd line without finishing 1st.
                        Create1stPartofDtl(Rsd2)
                        DTL = DTL & "PBRUDE  " & DR2("lineprice") & vbNewLine & "PBRUDE_IEPS  " & DR2("lineprice") & vbNewLine
                        DTL = DTL & "IMPBRU  " & DR2("lineprice") & vbNewLine & "IMPBRU_PRT  " & CDbl(DR2("lineprice")) & vbNewLine
                        DTL = DTL & "VALUNI  " & DR2("lineprice") & vbNewLine & "IMPORT  " & DR2("lineprice") & vbNewLine
                        If Val(CStr(IEPSTasa > 0)) Then DTL = DTL & "TASIEP  " & IEPSTasa & vbNewLine & "MONIEP  " & CDbl(IEPSAmt) & vbNewLine 'VERSION 4 ADDED IF ONLY IF >0
                        LineImpIvaIeps = DR2("lineprice")
                        TotImp_Prt = TotImp_Prt + DR2("lineprice")
                    Case Is = "ST"
                        If Created1stPart = False Then
                            Create1stPartofDtl(Rsd2) 'only one line, create dtl  'MC should be read first and if exists, flag will be true.
                            Created1stPart = True
                        End If
                        If Created2ndPart = False Then
                            Create2ndPartOfDtl()
                            Created2ndPart = True
                        End If
                        TotalTax = CStr(Val(TotalTax) + Val(DR2("lineprice")))
                        TotalTax_New = TotalTax_New + CDbl(TotalTax)
                        LineImpIvaIeps = LineImpIvaIeps + Val(DR2("lineprice"))

                        If Val(CStr(IVATasa)) > 0 Then
                            DTL = DTL & "MONIPE  " & DR2("lineprice") & vbNewLine
                            DTL = DTL & "TASIPE  " & IVATasa & vbNewLine
                        End If
                        FinishedDtl = False
                    Case Else
                        SQL = "select nvl(ieps_rate, 0)ieps_rate from remark_code where remarkid = '" & DR2("remarkid") & "'"
                        OCM = New OracleCommand(SQL, conn)
                        dtd3 = New DataTable
                        dtd3.Load(OCM.ExecuteReader)

                        If dtd3.Rows.Count > 0 Then
                            DRd3 = dtd3.Rows(0)
                            If (DRd3("ieps_rate")) >= 0 Then 'printing ieps in separate line.  should be alone only if there is only one line.
                                IEPSTasa = DRd3("ieps_rate")
                                IEPSAmt = DR2("lineliqtax")
                                TotIEPSAmt = TotIEPSAmt + CDbl(IEPSAmt)
                                If SeparateLiqTaxFlag = "Y" And Val(CStr(IEPSTasa)) > 0 Then
                                    TotIEPS_Prt = CStr(Val(TotIEPS_Prt) + CDbl(IEPSAmt))
                                    IEPSAmt_Prt = IEPSAmt
                                End If
                                If IEPSTasa <> 0 Then IEPSTasaGlobal = IEPSTasa
                                If Created1stPart = False Then
                                    Create1stPartofDtl(Rsd2) : Created1stPart = True
                                End If
                                If Created2ndPart = False Then
                                    Create2ndPartOfDtl()
                                    Created2ndPart = True
                                End If
                                If Val(CStr(IEPSTasa)) > 0 Then DTL = DTL & "TASIEP  " & IEPSTasa & vbNewLine & "MONIEP  " & CDbl(IEPSAmt) & vbNewLine
                                DTL = DTL & "MONIEP_IEPS  " & IEPSAmt_Prt & vbNewLine
                                If DR2("lineprice") = 0 Then 'Line price is 0, use the amount in line total
                                    TmpTotImp = DR2("linetotal")
                                Else
                                    TmpTotImp = DR2("lineprice")
                                End If
                                If SeparateLiqTaxFlag = "Y" Then
                                    LineImpIvaIeps = LineImpIvaIeps + CDbl(TmpTotImp)
                                Else 'DO NOT SEPARATE, ADD IEPS TO UNIT PRICE.
                                    'Now we create the part that was not created because it it only created when we separate ieps.
                                    LineImpIvaIeps = LineImpIvaIeps + CDbl(TmpTotImp)
                                    TotImp_Prt = TotImp_Prt + CDbl(TmpTotImp)
                                    DTL = DTL & "PBRUDE_IEPS  " & TotImp_Prt & vbNewLine
                                    DTL = DTL & "IMPBRU_PRT  " & TotImp_Prt & vbNewLine
                                End If
                                FinishedDtl = True
                                TotAmtToBeTaxed = TotAmtToBeTaxed + DR2("lineprice")
                            End If
                        End If
                End Select
                j = j + 1
            Next

            If j > 1 Then
                If Val(IEPSAmt) = 0 Then
                    DTL = DTL & "TASIEP " & "0" & vbNewLine
                    DTL = DTL & "MONIEP " & "0" & vbNewLine
                End If
                If CDbl(TotalTax) = 0 Then
                    DTL = DTL & "TASIPE " & "0" & vbNewLine
                    DTL = DTL & "MONIPE " & "0" & vbNewLine
                End If

            End If
            If j > 1 And FinishedDtl = False Then
                DTL = DTL & "PBRUDE_IEPS  " & UnitPrice & vbNewLine & "IMPBRU_PRT  " & UnitPrice & vbNewLine '
            End If
            DTL = DTL & "IMPIVAIEPS  " & LineImpIvaIeps & vbNewLine & "NUMLIN  " & TotLines & vbNewLine 'This is the last line of detail, after reading iva ieps, then print total import as one line.
        End If
        Exit Sub
ErrorHandler:  'write error file.
        ErrMsgLog = Err.Number & "-->" & Err.Description & "." & vbCrLf & "Error on NoItemCdRoutine sub."
        Call CkFileExists(DirToOutputError, ErrMsgLog, "ERR-DigInv3_3.log")
        SendEmail(ErrMsgLog)
    End Sub

    '***********************************************************************'
    'Name: Create1stPartofDtl
    'Description: Adds in details such as description, discount, unit for an item 
    'Params: 
    '   - DR2 : The Data Row being passed in to create the 1st part of details for 
    'Return Value: N/A
    'Precondition(s): N/A   
    'Postcondition(s): 1st part of detail has been added to global variable DTL for the data row passed in 
    '***********************************************************************'
    Private Sub Create1stPartofDtl(ByRef DR2 As Object)
        DTL = DTL & "CANTID  " & "1" & vbNewLine & "CANTID_EA  " & "1" & vbNewLine & "DESCRI  " & DR2("linedesc").ToString.Trim & vbNewLine
        DTL = DTL & "CANEMP  " & vbNewLine & "UNIDAD  " & "EA" & vbNewLine
        DTL = DTL & "TDECON  " & "0" & vbNewLine & "MDECON  " & "0.00" & vbNewLine 'NO DESCUENTO EN DETALLE HARD CODE 0
        TotLines = TotLines + 1
    End Sub

    '***********************************************************************'
    'Name: Create2ndPartofDtl
    'Description: Adds in additional details to the data row being processed
    'Params: N/A
    'Return Value: N/A
    'Precondition(s): N/A
    'Postcondition(s): Additional details have been added to the global variable DTL 
    '***********************************************************************'
    Private Sub Create2ndPartOfDtl()
        DTL = DTL & "PBRUDE  " & vbNewLine & "PBRUDE_IEPS  " & vbNewLine
        DTL = DTL & "IMPBRU  " & vbNewLine & "IMPBRU_PRT  " & vbNewLine
        DTL = DTL & "VALUNI  " & "0" & vbNewLine & "IMPORT  " & "0" & vbNewLine
    End Sub

    '***********************************************************************'
    'Name: IepsSeparate
    'Description: Determines if an item has a separate liquor tax
    'Params: 
    '   - InvNo : The invoice number tied to the item to be looked up
    '   - LineNo : The line number
    'Return Value: Either a 'Y' indicating that the line number has separate liquor tax, or 'N' indicating it does not
    'Precondition(s): N/A
    'Postcondition(s): N/A
    '***********************************************************************'
    Private Function IepsSeparate(ByRef InvNo As String, ByRef LineNo As String) As String
        Dim SQL As String
        Dim OCM As OracleCommand
        Dim DT As DataTable
        Dim DR As DataRow
        On Error GoTo ErrorHandler

        SQL = "select IepsSeparate('" & InvNo & "', '" & LineNo & "') separate from dual"

        OCM = New OracleCommand(SQL, conn)
        DT = New DataTable
        DT.Load(OCM.ExecuteReader)

        If DT.Rows.Count > 0 Then
            DR = DT.Rows(0)
            If Val("" & DR("SEPARATE")) = 1 Then
                IepsSeparate = "Y"
                IEPS_NO_SeparateXml = False
            Else
                IepsSeparate = "N"
                IEPS_NO_SeparateXml = True
            End If
        Else
            IepsSeparate = "N"
            IEPS_NO_SeparateXml = True
        End If
        Exit Function
ErrorHandler:  'write error file.
        ErrMsgLog = Err.Number & "-->" & Err.Description & "." & vbCrLf & "Error on call stored procedure."
        Call CkFileExists(DirToOutputError, ErrMsgLog, "ERR-DigInv3_3.log")
        SendEmail(ErrMsgLog)
    End Function

    '***********************************************************************'
    'Name: StripRFC
    'Description: Strips comma, periods, spaces from the string passed in 
    'Params: 
    '   - RFC : The string to be stripped 
    'Return Value: A string with the above mentioned characters removed
    'Precondition(s): N/A
    'Postcondition(s): A stripped version of the string passed in 
    '***********************************************************************'
    Private Function StripRFC(ByRef RFC As String) As String
        Dim TmpStr As String
        TmpStr = RFC
        TmpStr = Replace(TmpStr, "R.F.C.", "")
        TmpStr = Replace(TmpStr, ".", "")
        TmpStr = Replace(TmpStr, " ", "")
        StripRFC = TmpStr
    End Function

    '***********************************************************************'
    'Name: MoneyPhrase
    'Description: Converts a decimal version of a monetary amount into the legal amount (in words)
    'Params: 
    '   - Amount : The amount to be converted into words
    'Return Value: String value containing the legal amount
    'Precondition(s): N/A
    'Postcondition(s): N/A
    '***********************************************************************'
    Function MoneyPhrase(ByVal Amount As Double) As String
        Dim buff, done As String '( 'buff without "," and with ".")
        Static Units(10) As String
        Static teens(10) As String
        Static tens(10) As String
        Static denoms(4) As String
        Dim Index, Length As Short
        Dim i, passes As Short
        Dim Temp, remains As Double
        Dim EndofPesos As Boolean
        Dim originalamt As Object
        Dim before As Object
        Units(0) = "0" : Units(1) = "uno" : Units(2) = "dos"
        Units(3) = "tres" : Units(4) = "cuatro" : Units(5) = "cinco"
        Units(6) = "seis" : Units(7) = "siete" : Units(8) = "ocho"
        Units(9) = "nueve" : teens(0) = "diez"

        teens(1) = "once" : teens(2) = "doce" : teens(3) = "trece"
        teens(4) = "catorce" : teens(5) = "quince" : teens(6) = "dieciseis"
        teens(7) = "diecisiete" : teens(8) = "dieciocho" : teens(9) = "diecinueve"

        tens(1) = "diez" : tens(2) = "veinte" : tens(3) = "treinta"
        tens(4) = "cuarenta" : tens(5) = "cincuenta" : tens(6) = "sesenta"
        tens(7) = "setenta" : tens(8) = "ochenta" : tens(9) = "noventa"

        denoms(1) = "cien" : denoms(2) = " mil " : denoms(3) = "millon "

        buff = Format(Amount, "#########.00")
        originalamt = buff
        Length = Len(buff) - 3
        passes = Length
        Index = 0
        Temp = 0
        remains = 0
        EndofPesos = False
        Do While (Length > -1)
            Select Case (Length Mod 3)
                Case 0
                    If Mid(buff, 1, 1) = "." Then
                        If done = "un " Then
                            done = done & "peso "
                        Else
                            done = done & " pesos "
                        End If
                        Length = 3
                        EndofPesos = True
                    Else
                        If Length > 0 Then
                            If Len(done) > 0 Then
                                If before = "1" Then
                                    done = done & denoms(Length / 3 + 1)
                                Else
                                    If denoms(Length / 3 + 1) <> " mil " Then
                                        done = done & denoms(Length / 3 + 1) & "es "
                                    Else
                                        done = done & denoms(Length / 3 + 1)
                                    End If
                                End If
                            End If
                            If Mid(buff, 1, 1) <> "0" Then
                                i = Val(Mid(buff, 1, 1))
                                If Units(i) <> "uno" Then
                                    Select Case i
                                        Case Is = 5
                                            done = done & "quinientos "
                                        Case Is = 7
                                            done = done & "sete"
                                        Case Is = 9
                                            done = done & "nove"
                                        Case Else
                                            done = done & Units(i)
                                    End Select
                                    If i <> 5 Then
                                        done = done & denoms(1) & "tos "
                                    End If
                                Else
                                    done = done & denoms(1) & "to "
                                End If
                            End If
                        Else
                            i = Len(done) - 2
                        End If
                    End If
                    Length = Length - 1
                Case 1
                    If Mid(buff, 1, 1) <> "0" Then
                        i = Val(Mid(buff, 1, 1))
                        If i <> CDbl("1") Then
                            If Trim(done) <> "" Then
                                If before <> "0" Then
                                    If before = 2 Then 'If added to change 21 to veintiun not for 31 that still shuld be: treinta y un
                                        done = Mid(done, 1, Len(done) - 1) & "i" & Units(i)
                                    Else
                                        done = done & " y " & Units(i)
                                    End If
                                Else
                                    done = done & Units(i)
                                End If
                            Else
                                done = done & Units(i)
                            End If
                        Else
                            If Trim(done) <> "" Then
                                If before = 2 Then 'If added to change 21 to veintiun not for 31 that still shuld be: treinta y un
                                    done = Mid(done, 1, Len(done) - 1) & "iun"
                                Else
                                    done = done & " y un "
                                End If
                            Else
                                done = done & "un "
                            End If
                        End If
                    End If
                    If Mid(buff, 1, 1) = "0" And passes = 1 Then
                    Else
                        If Length = 1 And Len(done) = 1 Then
                            done = done & "peso"
                        End If
                    End If
                    Length = Length - 1
                Case 2
                    If Mid(buff, 1, 1) = "1" Then
                        i = Val(Mid(buff, 2, 1))
                        If EndofPesos Then
                            done = done & buff & "/100 M.N."
                            Length = -2 'to finish the loop
                        Else
                            done = done & teens(i)
                            buff = Mid(buff, 2)
                            Length = Length - 2
                        End If
                    Else
                        i = Val(Mid(buff, 1, 1))
                        If EndofPesos Then
                            done = done & buff & "/100 M.N."
                            Length = -2 'to finish the loop
                        Else
                            done = done & tens(i)
                            'If Trim(tens(i)) = "" And Mid(buff, 2, 1) = "0" Then '1,100 comes out as: un mil ciento
                            If Trim(tens(i)) = "" And Mid(buff, 2, 1) = "0" And before = "1" Then '1,100 comes out as: un mil ciento ==> added before because 300 comes out as: tres cient
                                done = Left(done, Len(done) - 3) & " "
                            End If
                            Length = Length - 1
                        End If
                    End If
            End Select
            before = Mid(buff, 1, 1)
            buff = Mid(buff, 2)
        Loop
        Mid(done, 1, 1) = UCase(Mid(done, 1, 1))
        MoneyPhrase = done
    End Function

    '***********************************************************************'
    'Name: CkFileExists
    'Description: Checks whether file already exists then writes the data passed in to the designated file
    'Params: 
    '   - StrDir : The directory to output the file to 
    '   - StrData : The data that will be written into the file
    '   - StrFile : The name of the file to be written
    'Return Value: N/A
    'Precondition(s): N/A
    'Postcondition(s): Data has been written or appended to the designated file, file handler has been closed
    '***********************************************************************'
    Sub CkFileExists(ByRef StrDir As String, ByRef StrData As String, ByRef StrFile As String)
        Dim filenum As Object
        Dim strFilePath As String
        Dim strFileName As String
        Dim strDirExist As String
        On Error GoTo ErrHndlr
        filenum = FreeFile()
        Dim tmpdir As String = StrDir
        If tmpdir.EndsWith("\") Then tmpdir = tmpdir.Substring(0, tmpdir.Length - 1)

        strDirExist = Dir(StrDir, 16)
        If strDirExist = "" Then
            StrData = "ERROR:  Directory not found:  " & StrDir & vbNewLine & StrData
            StrDir = DirToOutputError
            StrFile = "Error-JetsInvoice.txt"
            SendEmail(StrData)
        End If
        strFileName = StrDir & StrFile
        strFilePath = Dir(strFileName)

        If strFilePath = "" Then FileOpen(filenum, strFileName, OpenMode.Output) Else FileOpen(filenum, strFileName, OpenMode.Append)

        Print(filenum, StrData)
        FileClose(filenum)
        Exit Sub
ErrHndlr:
        ErrMsgLog = New String("*", 50) & vbNewLine & "Today: " & Format(Now, "MM/dd/yy HH:mm") & "---> Eror # " & Err.Number & "-->" & Err.Description & "." & vbNewLine & "Error on DigInv3 program, CkFileExists sub" & vbNewLine
        ErrMsgLog = ErrMsgLog & "PROGRAM ENDED WITHOUT PROCESSING ANYTHING!!!!" & vbNewLine
        Call CkFileExists(DirToOutputError, ErrMsgLog, "ERR-Dig-Inv3_3.log")
        SendEmail(ErrMsgLog)
    End Sub

    '***********************************************************************'
    'Name: GetTermMX
    'Description: Sets the terms date to the global variable TermsDate
    'Params: 
    '   - DR : The data row being processed
    'Return Value: N/A 
    'Precondition(s): N/A
    'Postcondition(s): The terms date has been set appropriately
    '***********************************************************************'
    Private Sub GetTermMX(ByRef DR As DataRow)
        If "" & DR("CreditCodeId") = "A" Or "" & "" & DR("CreditCodeId") = "B" Then
            TermsDate = Format(DR("invdate"), "yyyy-MM-dd")
        ElseIf "" & DR("codonliqinv") = "1" And "" & DR("sectioncode") = "5" Then
            TermsDate = Format(DR("invdate"), "yyyy-MM-dd")
        Else
            If DR("DISCPCNT") = 0 Then
                If DR("netdays") = 0 And DR("ardays") = 0 Then
                    TermsDate = Format(DateAdd(DateInterval.Day, 30, CDate(DR("invdate"))), "yyyy-MM-dd")
                Else
                    If DR("ardays") > DR("netdays") And DR("netdays") <> 0 Then
                        TermsDate = Format(DateAdd(DateInterval.Day, CInt(DR("netdays")), DR("invdate")), "yyyy-MM-dd")
                    ElseIf DR("ardays") = 0 And DR("netdays") > 0 Then
                        TermsDate = Format(DateAdd(DateInterval.Day, CInt(DR("netdays")), DR("invdate")), "yyyy-MM-dd")
                    Else
                        If DR("ardays") < 30 Then
                            TermsDate = Format(DateAdd(DateInterval.Day, CInt(DR("ardays")), DR("invdate")), "yyyy-MM-dd")
                        Else
                            TermsDate = Format(DateAdd(DateInterval.Day, 30, CDate(DR("invdate"))), "yyyy-MM-dd")
                        End If
                    End If
                End If
            Else
                If DR("netdays") <> 0 Then
                    TermsDate = Format(DateAdd(DateInterval.Day, CInt(DR("netdays")), CDate(DR("invdate"))), "yyyy-MM-dd")
                Else
                    TermsDate = Format(DateAdd(DateInterval.Day, 30, CDate(DR("invdate"))), "yyyy-MM-dd")
                End If
            End If
        End If
    End Sub

    '***********************************************************************'
    'Name: SendEmail
    'Description: Sends email to the error email account
    'Params: 
    '   - MSG : The message to be sent in the email
    'Return Value: N/A
    'Precondition(s): N/A
    'Postcondition(s): The email has been sent out to the designated email account, or in case of an error the error message has been written to the error log
    '***********************************************************************'
    Private Sub SendEmail(ByRef MSG As String)
        Dim Body, EmailAddress, Subject As String
        On Error GoTo ErrHndlr
        Dim ObjMessage As Object

        EmailAddress = "jmx_diginv_error@kmsnet.com" 'Remove "_" in production | JBS Mex

        ObjMessage = CreateObject("CDO.Message")
        With ObjMessage
            .From = """JMX-SERVER"" <notice@kmsnet.com>"
            .To = EmailAddress
            .Subject = "ERROR - JMX - DigInv3 - "
            Body = Subject & vbNewLine & "Today: " & Format(Now, "MM/dd/yyyy HH:mm") & vbNewLine & "Check error folder on JMX: " & DirToOutputError & vbNewLine & vbNewLine & MSG
            .textBody = Body
            .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.kmsnet.com"
            .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
            .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
            .Configuration.Fields.Update()
        End With
        ObjMessage = Nothing
        Exit Sub
ErrHndlr:
        ErrMsgLog = "Today: " & Format(Now, "MM/dd/yy HH:mm") & "---> Eror # " & Err.Number & "-->" & Err.Description & "." & vbNewLine & "Error on DigInv2 program, SendEmail sub" & vbNewLine
        Call CkFileExists(DirToOutputError, ErrMsgLog, "ERR-DigInv3_3-SendEmail.txt")
    End Sub

    '***********************************************************************'
    'Name: GetSugarCuota
    'Description: Retrieves the sugar tax rate according to the year passed in
    'Params: 
    '   - Year : The year that the sugar tax rate is related to 
    'Return Value: The sugar tax rate for the year passed in 
    'Precondition(s): N/A
    'Postcondition(s): Sugar tax rate has been returned 
    '***********************************************************************'
    Private Function GetSugarCuota(Year As String) As Double
        Dim SQL As String, RS As Object, RS1 As Object
        Dim OC As OracleCommand
        Dim DT As DataTable

        On Error GoTo ErrHndlr
        SQL = "select * from sugartaxrate where year = '" & Year & "'"

        OC = New OracleCommand(SQL, conn)
        DT = New DataTable
        DT.Load(OC.ExecuteReader)

        If DT.Rows.Count > 0 Then
            GetSugarCuota = DT.Rows(0)("SUGARRATE")
        Else
            SQL = "select * from sugartaxrate order by year desc "
            OC = New OracleCommand(SQL, conn)
            DT = New DataTable
            DT.Load(OC.ExecuteReader)
            If DT.Rows.Count > 0 Then
                GetSugarCuota = DT.Rows(0)("SUGARRATE")
            End If
        End If
        Exit Function
ErrHndlr:
        ErrMsgLog = String.Format(50, "*") & vbNewLine & "Today: " & Format(Now, "mm/dd/yy hh:mm") & "---> Eror # " & Err.Number & "-->" & Err.Description & "." & vbNewLine &
"Error on DigInv3 program, GetSugarCuota Function" & vbNewLine & "SQL: " & SQL & vbNewLine
        ErrMsgLog = ErrMsgLog & "PROGRAM ENDED.  Year: " & Year & vbNewLine
        Call CkFileExists(DirToOutputError, ErrMsgLog, "ERR-Dig-Inv3_3.log")
        SendEmail(ErrMsgLog)
        End
    End Function

    '***********************************************************************'
    'Name: StripChars
    'Description: Removes , . - and " " from the string passed in 
    'Params: 
    '   - WithChars : The string to be stripped 
    'Return Value: A string with the specified characters removed
    'Precondition(s): N/A
    'Postcondition(s): The stripped version of the string passed in has been returned 
    '***********************************************************************'
    Private Function StripChars(ByRef WithChars As String) As String
        Dim TmpStr As String
        TmpStr = WithChars
        TmpStr = Replace(TmpStr, ",", " ")
        TmpStr = Replace(TmpStr, ".", "")
        TmpStr = Replace(TmpStr, "-", "")
        TmpStr = Replace(TmpStr, " ", "")
        StripChars = TmpStr
    End Function

    '***********************************************************************'
    'Name: RoundUpToDecimals
    'Description: Rounds the value passed in to the desired digit position decimal form
    'Params: 
    '   - value : The value to be rounded up 
    '   - digitPosition : The desired digit position to be rounded up to 
    'Return Value: The rounded up decimal value
    'Precondition(s): N/A
    'Postcondition(s): N/A
    '***********************************************************************'
    Function RoundUpToDecimals(value As Double, digitPosition As Integer) As Double
        Try
            ' Calculate the scaling factor based on the desired number of decimal places
            Dim scale As Decimal = Convert.ToDecimal(Math.Pow(10, digitPosition + 1))

            ' Multiply by the scaling factor to bring the relevant digit to the integer part
            Dim scaledValue As Decimal = Convert.ToDecimal(value) * scale

            Dim scaledValueInt = CInt(Math.Floor(scaledValue))

            ' Determine if the next digit (one more than desired) is 5 or greater
            Dim nextDigit As Integer = scaledValueInt Mod 10

            ' Truncate the scaled value to remove the extra digit
            scaledValue = Math.Floor(scaledValue / 10)

            ' Add 1 to the last kept decimal place if the next digit is 5 or greater
            If nextDigit >= 5 Then
                scaledValue += 1
            End If

            ' Scale back to the original number of decimal places
            Return scaledValue / Math.Pow(10, digitPosition)

        Catch ex As Exception
            Return value
        End Try
    End Function

End Module

