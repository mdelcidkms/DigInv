Option Strict Off
Option Explicit On
Imports System.Data.OracleClient
Module ModMain
    'Program written by Yadira Cerda on 10/09/17
    'This program will be used only in JFC Mexico branch
    'This program will be called by invoice program to generate the text data for
    'Digital invoice CFDI version 3.3
    'Modified on July 31 so that when liquor invoices are not separated (entered by phone), it won't error out (ie do not separate on pdr nor xml).
    '-----------------
    'April 2022 started modification for version 4.0
    '----------------
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
    Dim DiscPer As Double 'V 3.3 Discount diff calculation now
    Dim RFC As String
    Dim TotIEPS_Prt As String
    Dim IEPS_NO_SeparateXml As Boolean 'Added IEPSeparate on 04-09-2019 SIR 1794
    Dim TotAmtBeforeTaxes As Double
    Dim TermsDate As String 'Terms date added 11/21/17
    '-----2021-september add for new CFDI CARTA PORTE
    Dim DTLCPT, COMCP, HDRCP, DTLCP, TOTCP, FolioCP As String
    Dim SERIECP As String 'CP'
    Dim QtyTot As Short
    Dim TotWeight As Double
    Dim LiquorPresent As Boolean 'GLOBAL INVOICE LEVEL
    Dim IEPSnoIVA As Double '------version 4.0 need to add to the amount of IEPS TO THE NO-TAXABLE AMOUNT
    Dim SorianaErr As Boolean 'Soriana Folio de Entrega
    '50202200 is item catalog for liquor, mx_sat_catalog_id - will use for carta porte for carga peligrosa--
    '-----
    '**
    Dim IEPSTasaGlobal, TotAmtToBeTaxed, TotAmtNotTaxable, TotIEPSAmt As Double

    Dim ArrIeps_Sep_Xml(,) As Object '--> 04/05/24 define the number of dimensions
    Dim TotPercentages As Short
    Dim HdrNo_SeparateXmlFlag As Boolean '----->>>>>>>>07/26/19--flag will be true if at least one item qualifies for no separation---->>>>>>>>>>>>>>>>>>>>>>>>>

    Const DirToOutputError As String = "C:\JETS\Spool\JMXDigInvCMLogs\"
    Const IEPSperLiter As Double = 1.17 '05/09/18
    '*****************************S
    Const JFCID As String = "1" '--> Field(14) ID del Emisor
    'Added for SSL e-mail
    Const cdoSendUsingPickup As Short = 1 'Send message using the local SMTP service pickup directory.
    Const cdoSendUsingPort As Short = 2 'Send the message using the network (SMTP over the network).
    Const cdoAnonymous As Short = 0 'Do not authenticate
    Const cdoBasic As Short = 1 'basic (clear-text) authentication
    Const cdoNTLM As Short = 2 'NTLM
    Const AoApro As String = "2012" : Const NoApro As String = "32191" : Const FormaDePago As String = "PAGO EN UNA SOLA EXHIBICION"
    Const JFCRFC As String = "JME9707037KA"
    Const Expedido As String = "06700" '"CIUDAD DE MEXICO" '04-01-22 expedido LUGEXP is now codigo postal
    'Const JFCName = "JFC DE MEXICO, S.A. DE C.V.": Const JFCGln = "7504016094006"
    Const JFCName As String = "JFC DE MEXICO" : Const JFCGln As String = "7504016094006" 'VERSION 4.0 JUST NAME, NO S.A.........
    Const Regimen As String = "Régimen General de Ley Personas Morales"
    Dim TotImp, TotImp_Prt As Double
    Dim TotLines As Short ' Importe a imprimir may be different than total importe for cadena if no IEPS desglose.
    Dim DirToOutputData As String '04/17/15
    Dim DiscAmt, TotDiscAmt As Double '02/09/18 discount to be used for free items as well as for freight charge
    Dim UsoCFDI As String '09/12/18 per Ma Elena, USOCFDI it might be G01 or G03 so customer master will hold the value (for Invoices), for Nota de cargo, value will be PO1
    Dim ClaveFormaPago As String '99 is used for everything except nota de cargo to Comercial City Fresko want 15
    Dim ItemCatalog, DocName, UOMCatalog As String 'moved to general 02/26/19
    Public Sub Main()
        If Not DBOpened Then OpenDatabase()
        GetBranchName()
        RetrieveInvoices()
        End
    End Sub
    '==============================
    Private Sub RetrieveInvoices()
        '==============================
        Dim SQL, SqlTot As String
        Dim RS, RSTot As Object
        Dim s, i, t, x As Short '06/19/23 added x
        Dim DueDate As String 'Field(8) ConditionsofPymnt Due date ---Check where this field will come from
        Dim ExtraCharge As String ', RsnExtraCharge As String 'Freight only because this is in invhdr MC will be for nota de cargo only. If MC, invdtl will show up on dtl desc. an
        Dim TotalSales As String ', DiscAmt As Double 'DiscAmt moved General declarations
        Dim MetodoPago, MontoLetra, Section, Email, NumDpt As String 'NumDpt used for addenda comercial mexicana
        Dim Copies As Short
        Dim Folio, Printer, Serie, IepsID As String ', FolioCP As String
        Dim TipoDoc, InvMsg, MSG123 As String
        Dim FE As String
        Dim FEPos As Short ', SorianaErr As Boolean  'Soriana Folio de Entrega
        Dim Cita As String
        Dim CitaPos As Short '=========>>>>>>>>>>>>>>>>>>>07/03/18 now we need to send cita number as well as folio de entrada
        Dim RefInvoice, RELTIP, UUID As String 'Added for SIR 1702 NOTAS DE CARGO 02/21/19



        Dim Exported As String
        Dim OCM, CMOrdUpd As OracleCommand ' 04/25/2024 add for vb.net upgrade
        Dim ODR As OracleDataReader '04/25/2024 add for vb.net upgrade
        Dim OTR As OracleTransaction '04/25/2024 add for vb.net upgrade
        Dim DT, DTTOT As DataTable '04/25/2024 add for vb.net upgrade
        Dim Adpt As OracleDataAdapter

        'Impuesto trasladado siempre 0
        On Error GoTo ErrHndlr
        'PRODUCION EXPORTED FLAG N.............
        'sql for order update because the program for whatever reason sometimes is called by invoice prog. simulstaneously several times.
        SQL = "select exported,moddate, moduser from sohdr where ordstateid = 'R' and (exported = 'N'  or exported = 'C' ) order by ordernum"
        'For update
        'RSOrdUpd = Db.CreateDynaset(SQL, &H0) '04/24/2024 replaced for vb.net upgrade
        'If RSOrdUpd.RecordCount = 0 Then Exit Sub 'Nothing to run.

        '        OTR = conn.BeginTransaction(IsolationLevel.ReadCommitted) 'Begin the whole transaction here!!!
        CMOrdUpd = New OracleCommand(SQL, conn)
        DT = New DataTable
        DT.Load(CMOrdUpd.ExecuteReader)
        ODR = CMOrdUpd.ExecuteReader

        If DT.Rows.Count = 0 Then Exit Sub

        'Se.begintrans() 'Begin the whole transaction here!!! 04/25/2024 replaced for vb.net upgrade
        'PRODUCTION ********************************
        'get Folio de Entrada from Message (FE). 11/22/17 ADDED FRGHTALLOW
        '03/12/18 added GroupID to the SQL will be used to differentiate between Soriana and TCM
        '09/12/18 ADDED jmx_diginv_uso TO THE SQL-- 07/17/20 ADDED METODOPAGO ON INV
        '--------10/13/21  ---mod sql nvl(usocfdi) to default to G01
        SQL = "select s.ordernum,s.ordstateid,s.sectioncode,s.customerid,s.custordnum,s.shiplane,s.salesmanid,c.groupid, invopt_liqtaxseparate, " & "s.exported,s.msgprintcode,s.message,s.shipdate,s.shipvia,s.orderdate,s.FRGCHARGID,s.DEPTCODE,  NVL(C.DISTANCE, 200) DISTANCE, " & "i.invhdrnum,i.totalsales,i.totaltax,i.totalmisc,i.totalredem,i.totalfreig,i.invdate, FRGHTALLOW, JMX_REGIMEN_INV, " & "n.name shipname,n.address shipaddr, n.address2 shipaddr2, n.address3 shipaddr3, " & "n.city shipcity, n.state shipstate, n.zipcode shipzip, n.country shipcountry, n.phonenum shipphone, n.GLOBALLOCNUM shipgln, " & " c.name,c.SALSTXPCNT,c.taxid RFC, c.email, c.JFCVENDORNUM, c.GLOBALLOCNUM billgln, c.invmsg1, invmsg2, invmsg3, c.memo, " & "nvl(c.invopt_liqtaxseparate,'N')separateliqtax,c.phonenum,c.CREDITCODEID, c.CODONLIQINV,c.DISCPCNT,c.NETDAYS,c.ARDAYS,c.SALSTXPCNT, nvl(C.jmx_diginv_uso, 'G01') usocfdi, " & " c.address, c.address2, c.address3, c.city, c.state, c.zipcode,c.country, st.abbreviation soldstateabbr, ss.abbreviation shipstateabbr, " & "NVL(c.JMX_PMT_METHOD_INV,'PPD')METODOPAGO, NVL(c.JMX_PMT_FORM_INV,'99') FORMAPAGO, nvl(dockid, 'N') dockid From " & "sohdr s, invhdr i, customer c, shipname n, state st , state ss Where " & "s.ordernum = i.invhdrnum and s.customerid = c.customerid and c.customerid = n.customerid (+) and c.state = st.state (+) and n.state = ss.state (+) and ordstateid = 'R' " & " and (exported = 'N' or exported = 'C') order by invhdrnum "
        '10/07/21 adeed FormaPago to sql, used to be hard-coded 99, now read from customer master
        'Debug.Print SQL

        OCM = New OracleCommand(SQL, conn)
        'ODR = OCM.ExecuteReader()
        DT = New DataTable
        DT.Load(OCM.ExecuteReader)
        TotInvs = DT.Rows.Count()

        'RS = Db.CreateDynaset(SQL, &H4)
        'TotInvs = RS.RecordCount

        Dim SeparateIEPSAmt As Double
        Dim f As Short
        Dim SugarTaxAmt As Double '---------------!!!!!!!!!!!!!!!!!!!!!!!!!06/20/23!!!!!! '07/21/23
        Dim n As Short
        Dim Found As Boolean '07/21/23 'to add to subtbr at the end ....
        Dim InvhdrNo As String '04/22/24 add for upgrade VB.net
        Dim counter As Integer = 0 'JBS added for loop and take first

        If DT.Rows.Count() > 0 Then
            Copies = 2
            Printer = "1"
            SERIECP = "CP" '11/12/21
            '**********04/1715 modified to check if production or test environment ***************
            If gs_Hoststring = gs_UserName Then
                DirToOutputData = "\\JMXS2INV01\MASFAC_ESC\DATOS\" 'MASTEREDI REINSTALLED, AFTER THAT PREVIOUS PATH NOT WRKNG, WILL USE FULL PATH.
            Else
                'DirToOutputData = "\\JMXS2INVDEV01\MASFAC_ESC\DATOS\"
                DirToOutputData = "C:\JETS\Spool\JMXDigInvCM\"
            End If

            'For i = 1 To RS.RecordCount
            For Each DR As DataRow In DT.Rows
                InvMsg = vbNullString : TotIEPS_Prt = vbNullString : NumDpt = vbNullString : MSG123 = vbNullString
                RefInvoice = vbNullString : UUID = vbNullString : RELTIP = vbNullString 'clear only used for supermercados
                ErrFlag = False
                SlsTaxFlag = False 'If iva for any item
                SlsTaxInfo = vbNullString
                DiscExists = False : DiscPer = 0 'V 3.3
                If DR("totalfreig") < 0 Then 'JBS
                    DiscExists = True 'V 3.3c
                    DiscPer = DR("FRGHTALLOW") 'JBS
                End If

                'mod TotalTax from nullstring to "0.00" per Master EDI -- Some invoices were not printing due to this not being sent over when it was no tax.
                'TotalTax = vbNullString 'IVA Amt to be printed in the importe column next to slstaxinfo = (precio*cantidad)+IEPS%) * IVA%
                TotalTax = "0.00" : TotalTax_New = 0 'IVA Amt to be printed in the importe column next to slstaxinfo = (precio*cantidad)+IEPS%) * IVA%
                Section = vbNullString : IEPS_NO_SeparateXml = False
                SorianaErr = False '08/23/17  -->> MOVED UP HERE ON 02/21/19, WILL BE USING SAME ERROR FLAG WHEN NOTA DE CARGO, SUPEMARKET AND NO UUID FOUND
                TotAmtToBeTaxed = 0 : TotAmtNotTaxable = 0 : TotIEPSAmt = 0 : TotImp = 0 : TotImp_Prt = 0 : IEPSTasaGlobal = 0 : TotLines = 0 : TotAmtBeforeTaxes = 0
                Email = vbNullString & DR("Email").ToString()


                'VERSION 4.0 per email from Ma. Elena on 5/12/22 :  CFDI 4.0 solo debemos atender lo que diga su constancia de situación fiscal
                'commenting bellow for the above reason.  XML from now on should be separated all the time.
                '-------------------------------------------
                'If InStr(("" & RS!invopt_liqtaxseparate.Value), "X") Then IEPS_NO_SeparateXml = True Else IEPS_NO_SeparateXml = False  '--->>04/09/19 SIR #1794
                '---------------------------------------------------
                'UsoCFDI = RS!UsoCFDI.Value '***********09/12/18 --->>>> 02/21/19 sir # 1702 nota de cargo usocfdi = "P01" code bellow for Nota de Cargo.
                'TotInvAmt = (Val(DR("TotalSales")) + Val(RS!TotalTax.Value) + Val(RS!totalredem.Value) + Val(RS!totalfreig.Value) + Val(RS!totalmisc.Value))
                UsoCFDI = DR("UsoCFDI") '***********09/12/18 --->>>> 02/21/19 sir # 1702 nota de cargo usocfdi = "P01" code bellow for Nota de Cargo.
                TotInvAmt = DR("TotalSales") + DR("TotalTax") + DR("totalredem") + DR("totalfreig") + DR("totalmisc")
                NumDpt = vbNullString & DR("DEPTCODE")
                ClaveFormaPago = DR("formapago") '  "99"   '10/03/18  '---- 10/07/21 no more hard-code read from customer master
                If DR("totalmisc") > 0 Then '
                    'ACCORDING TO ED, TOTALMISC WILL BE USED ONLY FOR NOTAS DE CARGO, MIGHT HAVE MULTIPLE LINES
                    'talked to JFC people and they said, that ALL notas de cargo generate iva
                    'Notas de cargo are for items that are not JFC items
                    'ie: fletes, dried ice, saporo promotions, fletera messed up merchandise.
                    'TipoDoc = "I"  '---10/11/17 Now, TipoDoc will have either I (ingreso) or E (Egreso) (was 3, now I)
                    TipoDoc = "1" '---11/18/20 Now, TipoDoc will have either 1 (ingreso) or E (Egreso) (was 3, now I)
                    DocName = "NOTA DE CARGO"
                    Serie = "C"
                    Folio = gs_BranchN & DR("invhdrnum") 'GetDigInvNum("N")
                    FolioCP = DR("invhdrnum")
                    TotalSales = Format(TotInvAmt, "##0.00") 'NOTA DE CARGO IS REPORTED AS SALES
                    '----VERSION 4 CLAVEFORMAPAGO IS READ FROM DB,SHOULD NOT THIS HARD CODE ANY MORE COMMENTED 5/22/22
                    'If (vbNullString & RS!RFC.Value) = "CCF121101KQ4" Then ClaveFormaPago = "15" '10/03/18 EMAIL FROM MA. ELENA
                    'UsoCFDI = "P01"  '--->>>> 02/21/19 sir # 1702 nota de cargo usocfdi = "P01" 'VERSION 4 COMMENTED OUT , PO1 IS NOT IN CATALOG......
                    If Trim(vbNullString & DR("Memo")) = "SUPERMERCADO" Then '-------------------02/21/19  SIR 1702 --
                        RELTIP = "02"
                        FEPos = InStr(vbNullString & DR("message"), "FACTURA ") '
                        If FEPos > 0 Then
                            ''Debug.Print RS!message
                            RefInvoice = Trim(Mid(DR("message"), FEPos + 8, 9)) 'GRAB 9 CHARS AFTER FACTURA
                            UUID = GetRelUuid(RefInvoice)
                            If UUID = vbNullString Then SorianaErr = True
                        Else
                            SorianaErr = True
                        End If
                    End If
                Else
                    'TipoDoc = "I"  '---10/11/17 Now, TipoDoc will have either I (ingreso) or E (Egreso) (Was 1, now I)
                    TipoDoc = "1" '---11/18/20 Now, TipoDoc will have either 1 (ingreso) or E (Egreso) (Was 1, now I)

                    DocName = "FACTURA"
                    Serie = "A"
                    Folio = gs_BranchN & DR("invhdrnum") 'GetDigInvNum("I")
                    TotalSales = DR("TotalSales").ToString("0.00") 'Total sales is in the colum total sales???? because it separates other charges!!!, SAT???? //
                    '052124 JBS An: trimend to match VB result. 
                    'If Trim(vbNullString & DR("invmsg1")) <> vbNullString Then MSG123 = "1:  " & DR("invmsg1") 'added on 02/08/13
                    If Trim(vbNullString & DR("invmsg1")) <> vbNullString Then MSG123 = "1:  " & DR("invmsg1").ToString().TrimEnd() 'added on 02/08/13
                    MSG123 = MSG123 & " " & DR("invmsg2").ToString.Trim & " " & DR("invmsg3").ToString.Trim
                    ''ClaveFormaPago = "99"   '10/03/18   'now depends if it is NC to Fresko, we use 15, otherwise 99
                End If
                RFC = StripRFC(vbNullString & DR("RFC")) 'RFC is now in TaxID column
                'MetodoPago = "PPD" & vbNewLine 'ALWAYS USE PAGO POSTERIOR FOR ALL CUSTOMERS. -07/17/20 NO MORE
                MetodoPago = DR("MetodoPago") & vbNewLine
                DiscAmt = 0 : ExtraCharge = vbNullString '
                MontoLetra = MoneyPhrase(TotInvAmt)
                TermsDate = vbNullString '09/07/18  when COD, not date is calculated for now.
                DueDate = GetTermMX(DR) 'DueDate NOT BEING USED, WE USE TERMSDATE
                IVATasa = Val(DR("SALSTXPCNT"))
                If ((vbNullString & DR("msgprintcode")) = "B" Or (vbNullString & DR("msgprintcode")) = "I") And Trim(vbNullString & DR("message")) <> vbNullString Then
                    InvMsg = "2:  " & DR("message")
                End If
                InvMsg = Replace(InvMsg, vbTab, "") '-- 10/14/15 When copy and paste, it seems that a tab also gets pasted, and it cause msg not printed.
                IepsID = "GST" '01/09/13
                If RFC = "NWM9709244W4" Then IepsID = vbNullString '
                'ENCABEZADO -- now versio 3.3 not 2.0 ---STARTING ON JULY 1ST, VERSION 4.0 NOT 3.3
                'HDR = "E" & vbNewLine & "VERSIO" & Space(25 - Len("VERSIO")) & "3.3" & vbNewLine & "TRADPP" & Space(25 - Len("TRADPP")) & "MASTEDI" & vbNewLine
                HDR = "E" & vbNewLine & "VERSIO" & Space(25 - Len("VERSIO")) & "4.0" & vbNewLine & "TRADPP" & Space(25 - Len("TRADPP")) & "MASTEDI" & vbNewLine
                HDR = HDR & "SERFOL" & Space(25 - Len("SERFOL")) & Serie & vbNewLine
                HDR = HDR & "CTPPRO" & Space(25 - Len("CTPPRO")) & "ZZ" & vbNewLine & "NUMFOL" & Space(25 - Len("NUMFOL")) & Folio & vbNewLine
                HDR = HDR & "SERFOL" & Space(25 - Len("SERFOL")) & Serie & vbNewLine
                HDR = HDR & "FECEXP  " & Format(Now, "yyyy-MM-ddTHH:mm:ss") & vbNewLine & "NOAPRO  " & NoApro & vbNewLine & "AOAPRO  " & AoApro & vbNewLine
                'NEW ----------
                HDR = HDR & "CVEREGIMEN" & Space(15) & "601" & vbNewLine & "CVEFORPAG" & Space(16) & ClaveFormaPago & vbNewLine '10/16/17 --99 means 'por definir'--10/03/18 mod from "99" to claveformapago
                'Debug.Print HDR
                If vbNullString & DR("COUNTRY") <> "MEX" Then
                    'HDR = HDR & "RESFISC      " & RS!country.Value & vbNewLine     'si receptor es extranjero, clave del pais.
                    'HDR = HDR & "TAXID     " & RS!taxid & vbNewLine  'COMMENT UNTIL CLEAR IF WE WILL HAVE TAX ID AND WHERE!!!!!!!!!!!!
                End If
                HDR = HDR & "USOCFDI     " & UsoCFDI & vbNewLine ''"G01" & vbNewLine''**********09/12/18 NOW WE READ USO CDFI FROM CUSTOMER, IT MIGHT BE G01, OR G03
                HDR = HDR & "CVETIPDOC       I" & vbNewLine & "METPAG     PPD" & vbNewLine '11/14/17 'it seems that codmetpag is not being used for cfdi
                HDR = HDR & "CODMETPAG  " & MetodoPago & vbNewLine & "REGIMEN " & Regimen & vbNewLine & "FORPAG  " & FormaDePago & vbNewLine & "NUMCHE  " & vbNewLine & "TIPDOC  " & TipoDoc & vbNewLine
                HDR = HDR & "NOMDOC  " & DocName & vbNewLine & "FUNDOC  " & "O" & vbNewLine & "TIPMON  " & "MXN" & vbNewLine & "SW_TC   " & "0" & vbNewLine & "TIPCAM  " & "1" & vbNewLine
                HDR = HDR & "DIAPAG  " & (vbNullString & DR("ardays")) & vbNewLine & "PDPPAG  " & (vbNullString & DR("DISCPCNT")) & vbNewLine & "MDPPAG  " & vbNewLine
                HDR = HDR & "NUMEOC  " & DR("custordnum") & vbNewLine & "FECHOC  " & Format(DR("orderdate"), "yyyy-MM-dd") & vbNewLine & "FECCON  " & Format(DR("ShipDate"), "yyyy-MM-dd") & vbNewLine
                HDR = HDR & "FECPAG  " & TermsDate & vbNewLine
                'HDR = HDR & "FECPAG  " & DueDate & vbNewLine '11/21/17
                HDR = HDR & "LUGEXP  " & Expedido & vbNewLine & "MSG123  " & MSG123 & vbNewLine 'added msg123 02/07/133
                HDR = HDR & "NOTAS1  " & InvMsg & vbNewLine & "NOTAS2  " & vbNewLine & "NOTAS3  " & MontoLetra & vbNewLine & "AGENTE  " & (vbNullString & DR("salesmanid")) & vbNewLine
                HDR = HDR & "PEDIDO  " & gs_BranchN & "-" & DR("ordernum") & vbNewLine & "TRANSP  " & (vbNullString & DR("shipvia")) & vbNewLine
                HDR = HDR & "NUMDPT  " & NumDpt & vbNewLine & "NOMDPT  " & vbNewLine
                'HDR = HDR & "CONTRA  " & vbNewLine '==>>>>>>>>>07/03/18, NOW WE HANDLE THIS AFTER BELLOW CHECK OF SORIANA/TCM
                HDR = HDR & "NUMERO_IMP " & "1" & vbNewLine & "COPIAS  " & Copies & vbNewLine & "IEPS_ID " & IepsID & vbNewLine
                HDR = HDR & "REFFAC  " & vbNewLine & "NUMCLI  " & gstrBranchAlpha & "-" & DR("customerid") & vbNewLine & "NUMSAP  " & DR("invhdrnum") & vbNewLine
                HDR = HDR & "REMDES  " & vbNewLine
                '*******==========OCTOBER 2021===================******
                'HDRCP = HDRCP & "VERSIO  3.3" & vbNewLine & "SERFOL  CP " & vbNewLine & "NUMFOL  " & RS!invhdrnum.Value & vbNewLine
                '01/10/23 added version 4.0 to CARTA PORTE --- COMMENTED ABOVE, REPLACED WITH BELLOW...
                HDRCP = HDRCP & "VERSIO  4.0" & vbNewLine & "SERFOL  CP " & vbNewLine & "NUMFOL  " & DR("invhdrnum") & vbNewLine '----01/10/22
                HDRCP = HDRCP & "FECEXP  " & Format(Now, "yyyy-MM-ddTHH:mm:ss") & vbNewLine
                HDRCP = HDRCP & "CVETIPDOC   T" & vbNewLine & "TIPDOC  7 " & vbNewLine & "TIPMON  XXX" & vbNewLine & "TIPCAM  1" & vbNewLine
                'HDRCP = HDRCP & "USOCFDI P01" & vbNewLine  '01/10/23
                HDRCP = HDRCP & "USOCFDI S01" & vbNewLine '01/19/23 confirmed by Ma. Elena
                '------------
                '04/26/22 version 4
                'HDR = HDR & "CVEREGIMEN   " & RS!JMX_REGIMEN_INV & vbNewLine  '07/19/22
                HDR = HDR & "CVEREGREC   " & DR("JMX_REGIMEN_INV") & vbNewLine 'CVEREGREC is the correct nemonic for the customer.
                '*******=========================================******
                'SorianaErr = False  '08/23/17   -->> MOVED UP ON 02/21/19
                If RFC = "TSO991022PB6" Then '08/10/17 ADDED FOR SORIANA ADDENDA TO INCLUDE FOLIO O NOTA DE ENTRADA---->>>>mod on 07/03/18 now to also include numero de cita
                    Select Case (vbNullString & DR("groupid")) 'Select Case RS!customerid.Value   '--------------->>>>>>>>>>>>>>03/12/18
                        Case Is = "3925" '-Case Is = "7723" '11/09/17 case added because Soriana and TCM now share same RFC, but diff addendad and diff Portal --->>>>03/12/18, now we use groupid instead of customer
                            HDR = HDR & "TIPADD     3" & vbNewLine '11/09/17 to separate addenda Soriana de TCM
                            HDR = HDR & "CONTRA  " & vbNewLine
                        Case Else
                            FEPos = InStrRev(vbNullString & DR("message"), "FE") '08/23/17 to get Folio Nota de Entrada from message looking from end of msg.
                            If FEPos > 0 Then
                                'FE = Trim$(Mid$(RS!message, FEPos + 2, 7)) 'Read 7 only -->03/09/18 ADDED TRIM$, BECAUSE WHEN SPACE BETWEEN FE AND #, WAS ACCEPTING SPACE
                                FE = Trim(Mid(DR("message"), FEPos + 2)) '05/04/18 GRAB EVERYTHING AFTE FE
                                ''Debug.Print RS!message
                                If Not IsNumeric(FE) Then ' Or Len(FE) <> 7 Then '7 digits ---05/04/18 DONT' CK LEN ANYMORE
                                    SorianaErr = True '
                                Else
                                    HDR = HDR & "TDA_CONTRA     " & FE & vbNewLine
                                End If
                                'Debug.Print RS!MESSAGE
                                CitaPos = InStrRev(UCase(vbNullString & DR("message")), "CITA") '================>>>>>>>>>>>>>>>07/03/18
                                If CitaPos > 0 Then
                                    Cita = Trim(Mid(DR("message"), CitaPos + 4, FEPos - CitaPos - 4)) '-->>>07/03/18 TO GET THE LENGHT OF CITA.
                                    If Not IsNumeric(Cita) Then
                                        SorianaErr = True '===============>>>>07/03/18 EITHER NO CITA OR NO FOLIO DE ENTRADA, DO NOT PROCESS THE ORDER!!!!!!!!!!!!!!!!!
                                    Else
                                        HDR = HDR & "CONTRA     " & Cita & vbNewLine
                                    End If
                                Else 'CitaPos = 0, error, no cita found
                                    SorianaErr = True '===============>>>>07/03/18 EITHER NO CITA OR NO FOLIO DE ENTRADA, DO NOT PROCESS THE ORDER!!!!!
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
                '052124 JBS An: Added trim
                'HDR = HDR & "RFCEMI  " & JFCRFC & vbNewLine & "NOMEMI  " & JFCName & vbNewLine & "EANEMI  " & JFCGln & vbNewLine & "NUMEMI  " & DR("JFCVENDORNUM") & vbNewLine
                HDR = HDR & "RFCEMI  " & JFCRFC & vbNewLine & "NOMEMI  " & JFCName & vbNewLine & "EANEMI  " & JFCGln & vbNewLine & "NUMEMI  " & DR("JFCVENDORNUM").ToString().TrimEnd() & vbNewLine
                HDR = HDR & "CALEMI  " & "AV.AÑO DE JUAREZ NO. 160-B" & vbNewLine
                HDR = HDR & "NEXEMI  " & vbNewLine & "NINEMI  " & vbNewLine & "COLEMI  " & "COLONIA GRANJAS SAN ANTONIO" & vbNewLine & "LOCEMI  " & "MEXICO" & vbNewLine
                HDR = HDR & "MUNEMI  " & "IZTAPALAPA" & vbNewLine & "ESTEMI  " & "CIUDAD DE MÉXICO" & vbNewLine
                HDR = HDR & "REFEMI  " & vbNewLine & "TELEMI  " & "(55)5686-88-93" & vbNewLine
                '*******================OCTOBER 2021=========================******
                HDRCP = HDRCP & "RFCEMI " & JFCRFC & vbNewLine & "NOMEMI  " & JFCName & vbNewLine
                HDRCP = HDRCP & "CALEMI  " & "AV.AÑO DE JUAREZ NO. 160-B" & vbNewLine
                HDRCP = HDRCP & "NEXEMI  " & vbNewLine & "NINEMI  " & vbNewLine & "COLEMI  " & "1322" & vbNewLine & "LOCEMI  " & "09" & vbNewLine
                'HDRCP = HDRCP & "MUNEMI  " & "007" & vbNewLine & "ESTEMI  " & "DIF" & vbNewLine ' "CIUDAD DE MÉXICO" & vbNewLine '01/11/23
                HDRCP = HDRCP & "MUNEMI  " & "007" & vbNewLine & "ESTEMI  " & "CMX" & vbNewLine '01/11/23 use CMX instead of DIF per MasterEDI
                HDRCP = HDRCP & "PAIEMI  MEX" & vbNewLine & "CODEMI  09070" & vbNewLine & "EANEMI  " & JFCGln & vbNewLine
                HDRCP = HDRCP & "TIPCOM T " & vbNewLine & "CVEREGIMEN   601 " & vbNewLine & "USOCFDI G01" & vbNewLine
                '*******=========================================******
                'DATOS DEL RECEPTOR -- CLIENTE
                HDR = HDR & "RFCREC  " & RFC & vbNewLine & "NOMREC  " & Trim(DR("Name")) & vbNewLine
                HDR = HDR & "CALREC  " & DR("address").ToString.Trim & vbNewLine
                HDR = HDR & "COLREC  " & DR("address2").ToString.Trim & vbNewLine & "LOCREC  " & DR("city").ToString.Trim & vbNewLine
                HDR = HDR & "MUNREC  " & DR("address3").ToString.Trim & vbNewLine & "ESTREC  " & DR("soldstateabbr") & vbNewLine & "PAIREC  " & DR("COUNTRY").ToString.Trim & vbNewLine
                If vbNullString & DR("zipcode") = vbNullString Then 'VERSION 4 DO NOT LET ZIP CODE TO BE NULL BECAUSE ERROR REPORT IS NOT CLEAR
                    HDR = HDR & "CODREC  " & "99999" & vbNewLine 'IF USING 99999 WHEN NULL, THEN ERROR WILL SAY DOMICILIO (ADDRESS INCORRECT)
                Else
                    HDR = HDR & "CODREC  " & DR("zipcode") & vbNewLine
                End If
                '05202024 JBS An: remove trimend since VB has no trim logic on this field. 
                'HDR = HDR & "EANREC  " & DR("BillGLN") & vbNewLine & "TELREC  " & DR("phonenum") & vbNewLine & "MAIL    " & Email.ToString.TrimEnd & vbNewLine
                HDR = HDR & "EANREC  " & DR("BillGLN") & vbNewLine & "TELREC  " & DR("phonenum").ToString().TrimEnd() & vbNewLine & "MAIL    " & Email.ToString() & vbNewLine

                '*******================CLIENTE, PERO USAMOS JFCRFC FOR TRASLADO =========================******
                HDRCP = HDRCP & "RFCREC  " & JFCRFC & vbNewLine '& "NOMREC  " & DR("Name") & vbNewLine '& "CALREC  " & "AV.AÑO DE JUAREZ NO. 160-B" & vbNewLine ' 01/19/23  per MasterEDI  same as JFC
                HDRCP = HDRCP & "NOMREC  " & JFCName & vbNewLine '& "CALREC  " & DR("address") & vbNewLine
                HDRCP = HDRCP & "ESTREC CMX " & vbNewLine '-----01/11/23 --------------
                HDRCP = HDRCP & "CODREC 09070" & vbNewLine '01/19/23 as per Master EDI should be same as CODEMI
                '************SHIP TO INFO ************
                HDR = HDR & "RFCENT  " & RFC & vbNewLine & "NOMENT  " & DR("shipname").ToString.Trim & vbNewLine & "CALENT  " & DR("shipaddr").ToString.Trim & vbNewLine
                HDR = HDR & "COLENT  " & DR("shipaddr2").ToString.Trim & vbNewLine & "LOCENT  " & DR("shipcity").ToString.Trim & vbNewLine & "ESTENT  " & DR("shipstateabbr") & vbNewLine
                HDR = HDR & "MUNENT  " & DR("shipaddr3").ToString.Trim & vbNewLine & "ESTENT  " & DR("shipstate") & vbNewLine & "PAIENT  " & DR("shipcountry").ToString.Trim & vbNewLine
                HDR = HDR & "CODENT  " & DR("shipzip") & vbNewLine & "EANENT  " & DR("ShipGLN") & vbNewLine
                '*******================OCTOBER 2021=========================******
                HDRCP = HDRCP & vbNewLine & "NOMENT  " & DR("shipname").ToString.Trim & vbNewLine & "CALENT  " & DR("shipaddr").ToString.Trim & vbNewLine
                If vbNullString & DR("shipzip").ToString.Trim = vbNullString Then
                    HDRCP = HDRCP & "CODENT  " & DR("zipcode") & vbNewLine
                Else
                    HDRCP = HDRCP & "CODENT  " & DR("shipzip") & vbNewLine
                End If
                HDRCP = HDRCP & "PAIENT MEX " & vbNewLine
                HDRCP = HDRCP & "CVEREGIMEN  601" & vbNewLine & vbNewLine
                '******************* ------------CTP ----NOVEMBER 2021
                HDRCP = HDRCP & "NUMERO_IMP 1 " & vbNewLine & "COPIAS 2 " & vbNewLine 'Added to be able to print 12/29/21 -- changed copies from 1 to 2 12/31/21
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
                HDRCP = HDRCP & "COM_CPT_INICPT " & vbNewLine 'delimitador  inicio de la carta porte
                HDRCP = HDRCP & "  COM_CPT_VERSIO 3.1" & vbNewLine '2.0 " & vbNewLine 12/28/23 carta porte version 3.0
                HDRCP = HDRCP & "COM_CPT_IDCCP   GENERA" & vbNewLine 'VERSION 3 GENERATE UUID
                HDRCP = HDRCP & " COM_CPT_TRANINT No " & vbNewLine 'Mercancia no sale de territorio nacional
                HDRCP = HDRCP & "   COM_CPT_FT_CVETRANS 01 " & vbNewLine 'CveTransporte Clave de transporte 01-> auto transporte federal
                HDRCP = HDRCP & "   COM_CPT_TOTDIST " & DR("distance") & vbNewLine '  "200 " & vbNewLine 'TotalDistRec Total distancia recorrida en Km
                HDRCP = HDRCP & "   COM_CPT_TIPEST 2 " & vbNewLine & vbNewLine 'TipoEstacion
                'UBICACION ORIGEN
                HDRCP = HDRCP & "   COM_CPT_INIUBI " & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_UBI_TIPUBI Origen" & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_UBI_IDUBI OR000001 " & vbNewLine 'IDOrigen ID elige la compania. ---MODIFY TO ORIGEN -COULD BE 1 SINCE ORIGEN SAME---
                HDRCP = HDRCP & "      COM_CPT_UBI_RFC " & JFCRFC & vbNewLine 'RFC remitente (emisor)
                HDRCP = HDRCP & "      COM_CPT_UBI_NOM " & JFCName & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_UBI_FECHA " & Format(Now, "yyyy-MM-ddT12:00:00") & vbNewLine 'FechaHoraSalida
                HDRCP = HDRCP & "      COM_CPT_DOM_CAL " & "AV.AÑO DE JUAREZ NO. 160-B" & vbNewLine 'Calle de la dirección del domicilio de la ubicación.
                HDRCP = HDRCP & "      COM_CPT_DOM_COL " & "1322" & vbNewLine '"AV.AÑO DE JUAREZ NO. 160-B" & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_DOM_LOC 09 " & vbNewLine 'DR("city") & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_DOM_MUN " & "007" & vbNewLine '& DR("address3") & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_DOM_EST  CMX" & vbNewLine 'DIF " & vbNewLine  ----01/11/23
                HDRCP = HDRCP & "       COM_CPT_DOM_PAI " & "MEX " & vbNewLine 'DR("COUNTRY") & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_DOM_CP 09070" & vbNewLine
                HDRCP = HDRCP & "   COM_CPT_FINUBI " & vbNewLine & vbNewLine 'Fin de Ubicacion origen
                'UBICACION DESTINO
                HDRCP = HDRCP & "      COM_CPT_INIUBI " & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_UBI_TIPUBI Destino " & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_UBI_IDUBI " & "DE000001" & vbNewLine 'id  punto d llegada
                HDRCP = HDRCP & "      COM_CPT_UBI_RFC " & RFC & vbNewLine ''RFC de destino
                HDRCP = HDRCP & "      COM_CPT_UBI_NOM " & DR("shipname").ToString.Trim & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_UBI_FECHA " & Format(DateAdd(Microsoft.VisualBasic.DateInterval.Hour, 6, Now), "yyyy-MM-ddT12:00:00") & vbNewLine
                HDRCP = HDRCP & "      COM_CPT_UBI_DISTREC " & DR("distance") & vbNewLine 'DistanciaRecorrida
                '052124, JBS An: added trim
                'HDRCP = HDRCP & "      COM_CPT_DOM_CAL " & DR("shipaddr") & vbNewLine '
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
                HDRCP = HDRCP & "COM_CPT_FINUBI " & vbNewLine & vbNewLine '& "COM_CPT_INIMERS " & vbNewLine
                '*******************
                If Trim(RELTIP) <> vbNullString Then HDR = HDR & "RELTIP  " & RELTIP & vbNewLine '===02/21/19 SIR 1702 FOR NOTAS DE CARGO
                If Trim(UUID) <> vbNullString Then HDR = HDR & "RELUUID1 " & UUID & vbNewLine
                InvhdrNo = DR("invhdrnum")
                'Call GetDtl(i, InvhdrNo) 'GET DETAIL INFORMATION
                'Call GetCPVehiculo(i, DR("invhdrnum")) 'GET INFORMATION REGARDING TRANSPORTATION VEHICLE ---OCTOBER 2021
                Call GetDtl(i, InvhdrNo) 'GET DETAIL INFORMATION
                Call GetCPVehiculo(i, InvhdrNo) 'GET INFORMATION REGARDING TRANSPORTATION VEHICLE ---OCTOBER 2021
                If ("" & DR("FRGCHARGID")) = "Y" Then 'only if freight charge is at a header level, at detail only charges.
                    If Val(DR("totalfreig")) < 0 Then 'discount
                        DiscAmt = System.Math.Abs(Val(DR("totalfreig"))) 'ASUMING DISCOUNT AMT. ONLY COMES FROM FREIGHT (AFTER CHECKING DB)
                        TotDiscAmt = TotDiscAmt + DiscAmt '02/09/18
                        TotImp_Prt = TotImp_Prt + Val(DR("totalfreig")) '---------DiscAmt
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
                '*********************************************
                'SUBTBR value has total before ANY TAXES-  SHOULD BE SUM OF IMPORT VALUES
                'added if condition for totamtbeforetaxes because when no items ie nota de cargo, this amt was 0.  Was not being calculated.
                'TotamtbeforeTaxes is used in xml as subTotal (SUBTBR). When NC, subtotal was written as 0.
                'If TotAmtBeforeTaxes = 0 Then TotAmtBeforeTaxes = TotAmtNotTaxable + TotAmtToBeTaxed '--->>>>commented on 01/23/18
                If DocName = "NOTA DE CARGO" Then
                    TotAmtBeforeTaxes = TotAmtNotTaxable + TotAmtToBeTaxed 'NOTA DE CARGO CALCULATION ONLY.  SHOULD BE SAME AMOUNT AF SUBTOT
                End If
                If Val(TotalTax) = 0 Then IVATasa = 0
                TOTAL = TOTAL & "TOTCAJ   " & vbNewLine & "TOTCAN  " & vbNewLine
                'HERE!!!!!!!!!!!!!! CHECK THIS SECTION ------------- IF COMBINED IEPS SEPARATE WIHT NO SEPARATE------------HOW TO WORK THIS CALCUALTION!!!!!
                ''If IEPS_NO_SeparateXml And TotAmtToBeTaxed > 0 Then
                '-----------------------------------
                'comment this out not working version 4 where if customer does not separate --06/19/23
                'If HdrNo_SeparateXmlFlag And TotAmtToBeTaxed > 0 Then    '07/29/19
                'TOTAL = TOTAL & "SUBTBR  " & TotAmtToBeTaxed & vbNewLine '------>>><<04/09/19 SIR 1794 ADDED IF 06/19/23 comment
                'Else' 06/19/23 commented
                'VERSION 4 NEED TO ADD DISCOUNT TO SUBTBR..........05/22/22 -- IF LATER WE USE THE NO SEPARATE XML THEN ADD TOTDISCAMT FOR THE SUBTBR -- IGNORE FOR NOW SINCE EVERYTHING WILL BE SEPARATED ON XML
                'TOTAL = TOTAL & "SUBTBR  " & TotAmtBeforeTaxes & vbNewLine 'Version 4.0 commented this one to add Discount...
                'TOTAL = TOTAL & "SUBTBR  " & (TotAmtBeforeTaxes + TotDiscAmt) & vbNewLine  'VERSION 4.0 ADD TotDiscAmt  .....
                ''''''''''''''''''''!!!!!!!!!temporary will put this total after the ieps!!!!!!!!!!!!!!
                'TOTAL = TOTAL & "SUBTBR  " & (TotImp_Prt + TotDiscAmt) & vbNewLine  '06/19/23 COMMENT ABOVE TO USE SAME AS SUBTOT_PRT - MOD ON NO SEPARATE TAX -
                ''''''''TOTAL = IEPSnoIVA
                'End If '06/19/23 commented
                '--------------------------------------
                TOTAL = TOTAL & "MONDET  " & TotDiscAmt & vbNewLine & "PRCDSG  " & DiscPer & vbNewLine '"0" & vbNewLine 'added 0 to nemonico PRCDSG requerido en addenda.
                'MOVED THIS CODE AFTER FINDIND IEPS SEPARATE TAXES...................
                ' '********VERSION 4.0 -- WE DON'T NEED SUBTSI VALUE ANY MORE -EVERYTHING IS TAXABLE EVEN IF IT IS IVA TASA 0%
                ''TOTAL = TOTAL & "SUBTSI  " & TotAmtNotTaxable & vbNewLine '-  NOW IS BEING USED AS IMPIVATRA1
                '*****----------------------------
                'VERSION 4 COMMENT OUT THE SUBTOT , seems not necccesary- want to avoid any issues of calculations if not necessary.
                'TOTAL = TOTAL & "SUBTOT  " & TotAmtNotTaxable + TotAmtToBeTaxed & vbNewLine '---May 20, 22-------------------------------
                TOTAL = TOTAL & "SUBTOT_PRT  " & Format(TotImp_Prt, "#,##0.00;-#,##0.00;""") & vbNewLine
                '01/10/18 modified because if IVATasa is 0, customer Palacio de Hierro was rejecting invoices because MasMasfacura creates the following entry
                '<cfdi:Traslado Importe="0.00" TasaOCuota="0.000000" Impuesto="002" TipoFactor="Tasa" />
                'TOTAL = TOTAL & "TOTIVA  " & IVATasa & vbNewLine & "IVATRA  " & TotalTax & vbNewLine '01/10/18 only use total iva if tasa > 0
                'If IVATasa > 0 Then TOTAL = TOTAL & "TOTIVA  " & IVATasa & vbNewLine & "IVATRA  " & TotalTax & vbNewLine
                '04-02-18 Other customers wants 0 reported even if no tax, Palacio de Hierro does not want 0 reported when no tax.
                TOTCP = TOTCP & vbNewLine
                '--------04/26/22 version 4.0 if ivatasa > 0 then write the value on total, otherwise do not--------
                ' If RFC = "PHI830429MG6" Then 'Palacio de Hierro  '-----version 4.0 all customes not only palacio so I comment out
                If IVATasa > 0 Then
                    TOTAL = TOTAL & "TOTIVA  " & IVATasa & vbNewLine & "IVATRA  " & TotalTax & vbNewLine
                End If
                ' Else
                '05/09/18 when free products, it errors out if we put tasa iva and ivatra 0, so I added if condition.  Free products, discount percent is 100.
                '-VERSION 4.0 NEED TO CHECK FOR FREE PRODUCTS TO SEE HOW IT BEHAVES, FOR NOW I AM COMMENTING TO AVOID THE 0 VALUES ON TOTAL
                'If DiscPer <> 100 Then TOTAL = TOTAL & "TOTIVA  " & IVATasa & vbNewLine & "IVATRA  " & Format(TotalTax_New, "###0.00") & vbNewLine  ''TotalTax & vbNewLine '--04/12/19 01/10/18 only use total iva if tasa > 0
                ' End If
                '--------04/26/22 version 4.0 FCTTOTIVA SHOULD BE ZERO
                ' If IVATasa = 0 Then '05/22/22
                If IVATasa = 0 And TotInvAmt > 0 Then 'Version 4.0 when only free items, do use bellow... 05/22/22
                    '06/06/22 modify, it was giving error:  El certificado no cumple con algunos de los valores permitidos... took out the TIPFACIVA Tasa...
                    'TOTAL = TOTAL & "FCTTOTIVA   0.00" & vbNewLine & "TOTIVA  0 " & vbNewLine & "IVATRA 0 " & vbNewLine & "TIPFACIVA Tasa " & vbNewLine
                    TOTAL = TOTAL & "FCTTOTIVA   0.00" & vbNewLine & "TOTIVA  0 " & vbNewLine & "IVATRA 0 " & vbNewLine '& "TIPFACIVA Tasa " & vbNewLine
                End If
                SeparateIEPSAmt = 0
                'added bellow for 3.3 version to separate the totals by tasa
                If Trim("" & DR("invopt_liqtaxseparate")) <> vbNullString Then ' 07/21/23  added if to get the total ieps only if we need to separate something-  null means do not separate anything
                    'SqlTot = "select 'Tasa' as TIPFACIEP ,b.liqtax as TATIEP, sum(lineliqtax)/sum(lineprice) as FCTTATIEP,sum(lineliqtax) as IEPTRA , b.liquorcode " & vbNewLine
                    'SqlTot = SqlTot & "from invdtl d, branch_item b where d.itemcode = b.itemcode and invhdrnum = " & DR("ordernum & vbNewLine
                    'SqlTot = SqlTot & "and lineliqtax >0 and lineprice > 0   AND B.LIQTAX > 0 " & vbNewLine 'Do not get the sugar beberage, separate SQL for that.
                    'SqlTot = SqlTot & "group by liqtax, b.liquorcode  " '02/02/24 added b.liquorcode
                    '02/06/24 modify sql  some items might have null or 0 in liq code and have same liqtax cannot trust the data
                    SqlTot = "select 'Tasa' as TIPFACIEP ,b.liqtax as TATIEP, nvl(b.liquorcode,0)liquorcode, " & vbNewLine
                    SqlTot = SqlTot & " b.liqtax/100 as fcttatiep,sum(lineliqtax) as IEPTRA "
                    SqlTot = SqlTot & "from invdtl d, branch_item b where d.itemcode = b.itemcode and invhdrnum = '" & DR("ordernum") & "'" & vbNewLine
                    SqlTot = SqlTot & " and lineliqtax >0 and lineprice > 0   AND B.LIQTAX > 0 group by liqtax , nvl(b.liquorcode,0) "

                    '10/03/18 ADDED FOR NOTA DE CARGO. DIFFERENT SQL NEEDED.
                    If DocName = "NOTA DE CARGO" Then
                        SqlTot = "SELECT 'Tasa' as TIPFACIEP, IEPS_RATE AS TATIEP, LINELIQTAX as IEPTRA, nvl(b.liquorcode,0)liquorcode  " & vbNewLine '02/06/24 added liquorcode
                        SqlTot = SqlTot & "FROM INVDTL D, branch_item b ,remark_code R WHERE INVHDRNUM = '" & DR("ordernum") & "'" & vbNewLine
                        SqlTot = SqlTot & " AND R.REMARKID = D.REMARKID  and d.itemcode = b.itemcode AND LINELIQTAX  > 0 " & vbNewLine
                        HDRCP = vbNullString 'Nota de cargo does no have item informaion, for the moment we do no generate carta porte.
                    End If
                    'Debug.Print SqlTot
                    'RSTot = Db.CreateDynaset(SqlTot, &H4) '01/05/2024 replaced by OracleConection for upgrade in vb.net
                    OCM = New OracleCommand(SqlTot, conn)
                    DTTOT = New DataTable
                    'DTTOT.Load(OCM.ExecuteReader) '052224 JBS An: different way to read data for better performance
                    Adpt = New OracleDataAdapter(OCM)
                    Adpt.Fill(DTTOT)

                    t = 0 'Initialize t because it will be used in case there is sugary drinks
                    ''If IEPS_NO_SeparateXml = False Then  '--->>>>>SIR 1794 DO NOT SEPARATE ON XML added if,dont do this if xml no desglose--------<<<<comment out on 07/29/19<<<<<
                    'Found = False '07/21/23 not needed used for the totpercentages
                    'If RSTot.RecordCount > 0 Then '01/05/2024 replaced by OracleConection for upgrade in vb.net
                    If DTTOT.Rows.Count() > 0 Then
                        'For t = 1 To RSTot.RecordCount  '06/19/23
                        'For x = 1 To RSTot.RecordCount '06/19/23 '01/05/2024 replaced by OracleConection for upgrade in vb.net
                        For Each DRTOT As DataRow In DTTOT.Rows
                            '---******************07/21/23 not using the array anymore because now xml and pdf should be the same----
                            'Found = False  '06/21/23 was not generating total for the second tasa ie 30 no desglose 8 si desglose , detail is correct, but in tot was not generating the total for desglose
                            'For n = 1 To TotPercentages '------------------------------------
                            '    If ArrIeps_Sep_Xml(1, (n * 2 - 1)) = Val(DRTOT("tatiep")) Then '---------------
                            '       n = TotPercentages '--  -----------------------------Found not continue searching
                            '       Found = True 'Found on the array which means this tasa is not separated on xml, so it should no be in total either.
                            '    End If
                            ' Next n
                            'If Not Found Then --07/21/23 comment out
                            '---*********************************************************
                            '-----07/20/23 ck if iep was separated i detail, if so then put in total, if not, ignore these lines
                            '--------------------------------------------------------------------------------------
                            f = InStr("" & DR("invopt_liqtaxseparate"), "C") '07/21/23
                            If f = 0 Then f = InStr("" & DR("invopt_liqtaxseparate"), "Y")
                            If DRTOT("tatiep") = 8 And f > 0 Then '= "C" Then   'C is 8%  '07/20/23
                                t = t + 1 '06/19/23 ..SIR 2536.. xml error out if we do not use secuential iepx ... because not all types have separation on xml now I needed to change. version 4 we started separating everything on xml, bur new request says go according to customer setup
                                TOTAL = TOTAL & "TIPFACIEP" & t & Space(5) & DRTOT("TIPFACIEP") & vbNewLine
                                '052124 JBS An: get decimal value by removing decimal zero
                                'TOTAL = TOTAL & "TATIEP" & t & Space(5) & DRTOT("tatiep") & vbNewLine
                                'TOTAL = TOTAL & "TATIEP" & t & Space(5) & DRTOT("IEPTRA") & vbNewLine

                                'SeparateIEPSAmt = SeparateIEPSAmt + DRTOT("IEPTRA") '07/21/23
                                TOTAL = TOTAL & "TATIEP" & t & Space(5) & GetDecimalWOZero(DRTOT("tatiep")) & vbNewLine
                                TOTAL = TOTAL & "IEPTRA" & t & Space(5) & GetDecimalWOZero(DRTOT("IEPTRA")) & vbNewLine
                                '---
                                SeparateIEPSAmt = SeparateIEPSAmt + GetDecimalWOZero(DRTOT("IEPTRA")) '07/21/23
                                '---
                                'here------------if handle this on detail do not handle in this by total-- doesn't seem right, so commented out for now ?????????????????????
                                ' If DRTOT("tatiep = 8 Then
                                'IEPSnoIVA = IEPSnoIVA + DRTOT("IEPTRA  'IEPSnoIVA  add to amt not taxable for 0% IVA traslado, only 8% is that case
                                ' End If
                            End If '07/20/23
                            '---------------------------------(B) BEER 26.5--OR (Y) EVERYTHING--------------------'07/21/23
                            f = InStr("" & DR("invopt_liqtaxseparate"), "B")
                            If f = 0 Then f = InStr("" & DR("invopt_liqtaxseparate"), "Y")
                            If DRTOT("tatiep") = 26.5 And f > 0 And DRTOT("liquorcode") = 1 Then ' 02/02/24 added and DRTOT("liquorcode = 1 it could be 26.5%, but sake not beer
                                t = t + 1 '06/19/23 ..SIR 2536.. xml error out if we do not use secuential iepx ... because not all types have separation on xml now I needed to change. version 4 we started separating everything on xml, bur new request says go according to customer setup
                                TOTAL = TOTAL & "TIPFACIEP" & t & Space(5) & DRTOT("TIPFACIEP") & vbNewLine
                                '052124 JBS An: get decimal value by removing decimal zero
                                TOTAL = TOTAL & "TATIEP" & t & Space(5) & GetDecimalWOZero(DRTOT("tatiep")) & vbNewLine
                                TOTAL = TOTAL & "IEPTRA" & t & Space(5) & GetDecimalWOZero(DRTOT("IEPTRA")) & vbNewLine
                                '---
                                SeparateIEPSAmt = SeparateIEPSAmt + GetDecimalWOZero(DRTOT("IEPTRA")) '07/21/23
                                '---
                            End If
                            '---------------------------------(A) LIQUOR  26.5, 30 or 53  and liquor code (2,3) OR (Y) EVERYTHING--------------
                            '--------------------------------2024-02-02------ADDED SEPARATION IF 26.5 AND LIQCODE 2, (not beer)--- I DID MISSED THIS ORIGINALLY---
                            f = InStr("" & DR("invopt_liqtaxseparate"), "A")
                            If f = 0 Then f = InStr("" & DR("invopt_liqtaxseparate"), "Y")
                            If DRTOT("tatiep") = 26.5 And f > 0 And DRTOT("liquorcode") > 1 Then '
                                t = t + 1 'xml error out if we do not use secuential iepx ... because not all types have separation on xml . version 4 we started separating everything on xml, bur new request says go according to customer setup
                                TOTAL = TOTAL & "TIPFACIEP" & t & Space(5) & DRTOT("TIPFACIEP") & vbNewLine
                                '052124 JBS An: get decimal value by removing decimal zero
                                TOTAL = TOTAL & "TATIEP" & t & Space(5) & GetDecimalWOZero(DRTOT("tatiep")) & vbNewLine
                                TOTAL = TOTAL & "IEPTRA" & t & Space(5) & GetDecimalWOZero(DRTOT("IEPTRA")) & vbNewLine
                                '---
                                SeparateIEPSAmt = SeparateIEPSAmt + GetDecimalWOZero(DRTOT("IEPTRA"))
                                '---
                            End If
                            '--------------
                            f = InStr("" & DR("invopt_liqtaxseparate"), "A")
                            If f = 0 Then f = InStr("" & DR("invopt_liqtaxseparate"), "Y")
                            If DRTOT("tatiep") = 30 And f > 0 Then '
                                t = t + 1 '06/19/23 ..SIR 2536.. xml error out if we do not use secuential iepx ... because not all types have separation on xml now I needed to change. version 4 we started separating everything on xml, bur new request says go according to customer setup
                                TOTAL = TOTAL & "TIPFACIEP" & t & Space(5) & DRTOT("TIPFACIEP") & vbNewLine
                                '052124 JBS An: get decimal value by removing decimal zero
                                TOTAL = TOTAL & "TATIEP" & t & Space(5) & GetDecimalWOZero(DRTOT("tatiep")) & vbNewLine
                                TOTAL = TOTAL & "IEPTRA" & t & Space(5) & GetDecimalWOZero(DRTOT("IEPTRA")) & vbNewLine
                                SeparateIEPSAmt = SeparateIEPSAmt + GetDecimalWOZero(DRTOT("IEPTRA")) '07/21/23
                            End If
                            '---------------------------------(A) LIQUOR 30 OR (Y) EVERYTHING ------------
                            f = InStr("" & DR("invopt_liqtaxseparate"), "A")
                            If f = 0 Then f = InStr("" & DR("invopt_liqtaxseparate"), "Y")
                            If DRTOT("tatiep") = 53 And f > 0 Then 'Mid$("" & DR("invopt_liqtaxseparate, 1) = "A") Or (Mid$("" & DR("invopt_liqtaxseparate, 1) = "Y") Then
                                t = t + 1 '06/19/23 ..SIR 2536.. xml error out if we do not use secuential iepx ... because not all types have separation on xml now I needed to change. version 4 we started separating everything on xml, bur new request says go according to customer setup
                                TOTAL = TOTAL & "TIPFACIEP" & t & Space(5) & DRTOT("TIPFACIEP") & vbNewLine
                                TOTAL = TOTAL & "TATIEP" & t & Space(5) & GetDecimalWOZero(DRTOT("tatiep")) & vbNewLine
                                TOTAL = TOTAL & "IEPTRA" & t & Space(5) & GetDecimalWOZero(DRTOT("IEPTRA")) & vbNewLine
                                SeparateIEPSAmt = SeparateIEPSAmt + GetDecimalWOZero(DRTOT("IEPTRA")) '07/21/23
                            End If
                            'End If  '07/21/23 comment out end if of the not found... not used any more
                            'RSTot.movenext() 05/01/2024 replaced by OracleConnection for upgrade in vb.net
                        Next
                    End If
                    '  NOW SUGARY DRINKS!!!!!!!!!!  ---NOW IS 1.17 pe liter, so using constant IEPSperLiter
                    f = InStr("" & DR("invopt_liqtaxseparate"), "S")
                    If f = 0 Then f = InStr("" & DR("invopt_liqtaxseparate"), "Y")
                    If f > 0 Then 'Mid("" & DR("invopt_liqtaxseparate, 1) = "S" Then '07/14/23 don't get total sugar beverage if is not S, which is separate sugar
                        'SqlTot = "select 'Cuota' as TIPFACIEP ,1 as TATIEP, 1 as FCTTATIEP,sum(lineliqtax) as IEPTRA " & vbNewLine '05/09/18
                        SqlTot = "select 'Cuota' as TIPFACIEP , " & IEPSperLiter & " as TATIEP, " & IEPSperLiter & " as FCTTATIEP,sum(lineliqtax) as IEPTRA " & vbNewLine '05/09/18
                        SqlTot = SqlTot & "from invdtl d, branch_item b where d.itemcode = b.itemcode and b.liqtax2 > 0 and " & vbNewLine
                        '052624 JBS An: no need group by since it's not a column. 
                        SqlTot = SqlTot & "invhdrnum = '" & DR("ordernum") & "' and lineliqtax >0 and lineprice > 0 and b.liqtax2 > 0 group by 'Cuota' " 'liqtax2" 01/24/18 group by cuota instead of liqtax2.
                        'SqlTot = SqlTot & "invhdrnum = '" & DR("ordernum") & "' and lineliqtax >0 and lineprice > 0 and b.liqtax2 > 0; "
                        'RSTot = Db.CreateDynaset(SqlTot, &H4)

                        OCM = New OracleCommand(SqlTot, conn)
                        DTTOT = New DataTable
                        Adpt = New OracleDataAdapter(OCM)
                        Adpt.Fill(DTTOT)

                        'Debug.Print SqlTot
                        SugarTaxAmt = 0 '06/20/23
                        If DTTOT.Rows.Count > 0 Then
                            For Each DRTOT As DataRow In DTTOT.Rows
                                'If t = 0 Then t = t + 1 '06/19/23 comment out SIR # 2536--In case no other types of ieps were found, sugary drinks is the only ieps, otherwise it will come up as zero
                                t = t + 1 '06/19/23
                                TOTAL = TOTAL & "TIPFACIEP" & t & Space(5) & DRTOT("TIPFACIEP") & vbNewLine
                                TOTAL = TOTAL & "TATIEP" & t & Space(5) & GetDecimalWOZero(DRTOT("tatiep")) & vbNewLine
                                TOTAL = TOTAL & "IEPTRA" & t & Space(5) & GetDecimalWOZero(DRTOT("IEPTRA")) & vbNewLine
                                SugarTaxAmt = GetDecimalWOZero(DRTOT("IEPTRA"))
                                '------------------VERSION 4 ADD THE TOTAL IEP TO THE IVA 0%- only if no iva was reported for the item
                                ' 05/22/22 move this at detail level for 1.17 cuota because some items have iva some don't so i cannot do this at total.....
                                'IEPSnoIVA = IEPSnoIVA + DRTOT("IEPTRA  'IEPSnoIVA to be added to amount not taxable for 0% IVA traslado
                                'TotAmtNotTaxable = TotAmtNotTaxable + IEPSnoIVA
                                t = t + 1 'use t instead of s in case there is liquor items we increment from last liqour total
                            Next
                        End If
                        'RSTot.Close()
                    End If '07/14/23
                End If '                            07/21/23  end if of  DR("invopt_liqtaxseparate <> vbnullstring then
                '!!!!!!!!!!!!!!!--------------
                '................ 06/20/23 --- modification for no desglose --- somehow  subtbr is not bein calculated properly
                TOTAL = TOTAL & "##############" & vbNewLine
                TOTAL = TOTAL & "SUBTBR  " & (TotImp_Prt + TotDiscAmt) & vbNewLine '+ SeparateIEPSAmt) & vbNewLine  ' - SugarTaxAmt) & vbNewLine  '07/21/23  -DO NOT ADD SUGAR TAX GENERATES ISSUES ON SOME -----06/19/23 COMMENT ABOVE TO USE SAME AS SUBTOT_PRT - MOD ON NO SEPARATE TAX -
                TOTAL = TOTAL & "##############" & vbNewLine
                '********VERSION 4.0 -- WE DON'T NEED SUBTSI VALUE ANY MORE -EVERYTHING IS TAXABLE EVEN IF IT IS IVA TASA 0%-NEED TO CHECK MORE INVOICES
                'TOTAL = TOTAL & "SUBTSI  " & TotAmtNotTaxable & vbNewLine '-  NOW IS BEING USED AS IMPIVATRA1
                'FOR IVA 0% TASA VERSION 4.0   BELLOW
                If DocName = "NOTA DE CARGO" Then '05/25/22 'VERSION 4 ---------------
                    If CDbl(TotalTax) = 0 Then 'TOTIVA 0%
                        TOTAL = TOTAL & "IMPIVATRA1 " & (TotAmtToBeTaxed + TotIEPSAmt) & vbNewLine & "IVATRA1 0" & vbNewLine
                        TOTAL = TOTAL & "TIPFACIVA1 Tasa " & vbNewLine & "TOTIVA1 0 " & vbNewLine
                    Else ' TOTIVA 16%
                        TOTAL = TOTAL & "IMPIVATRA1 " & (TotAmtToBeTaxed + TotIEPSAmt) & vbNewLine & "IVATRA1 " & TotalTax & vbNewLine
                        TOTAL = TOTAL & "TIPFACIVA1 Tasa " & vbNewLine & "TOTIVA1 16 " & vbNewLine
                    End If
                Else '---------END ADDITION REGARDING NOTA DE CARGO TOTAL...... 05/25/22
                    If TotAmtNotTaxable > 0 Then
                        '------------05/20/22-------------------
                        'TOTAL = TOTAL & "IMPIVATRA1 " & (TotAmtNotTaxable) & vbNewLine  '10/24/22 when 8% this amt is wrong, do not use this nemonic.... example invoice:  390963701
                        TOTAL = TOTAL & "IVATRA1 0" & vbNewLine
                        TOTAL = TOTAL & "TIPFACIVA1 Tasa " & vbNewLine & "TOTIVA1 0 " & vbNewLine
                    Else
                        If TotInvAmt > 0 Then 'only do bellow if invoice amount > 0 - not for free items
                            TOTAL = TOTAL & "IMPIVATRA1 " & TotAmtToBeTaxed & vbNewLine & "IVATRA1 " & TotalTax & vbNewLine
                            TOTAL = TOTAL & "TIPFACIVA1 Tasa " & vbNewLine & "TOTIVA1 16 " & vbNewLine
                        End If
                    End If
                End If
                'If TotAmtToBeTaxed > 0 Then  'IGNORE SUBTAI FOR NOW, SEEMS NOT NECESSARY -MAY 2022
                '     'TOTAL = TOTAL & "SUBTAI  " & TotAmtToBeTaxed & vbNewLine
                'Else
                '     'TOTAL = TOTAL & "SUBTAI  " & TotAmtNotTaxable & vbNewLine
                'End If
                'Debug.Print TOTAL
                If TotAmtToBeTaxed > 0 And TotAmtNotTaxable > 0 Then 'two types of iva 0% and 16%
                    TOTAL = TOTAL & "IMPIVATRA2 " & TotAmtToBeTaxed & vbNewLine '
                    TOTAL = TOTAL & "IVATRA2  " & TotalTax & vbNewLine '
                    TOTAL = TOTAL & "TIPFACIVA2 Tasa " & vbNewLine & "TOTIVA2 16 " & vbNewLine
                End If
                '*****----------------------------
                TOTAL = TOTAL & "PRT_IEPTRA  " & TotIEPS_Prt & vbNewLine '04/20/18 per MasterEDI, instead of IEPTRA_PRT, now use PRT_IEPTRA
                TOTAL = TOTAL & "IVARET  " & "0" & vbNewLine & "ISRRET  " & "0" & vbNewLine
                '04/26/22 ---version 4.0
                If IEPS_NO_SeparateXml And TotalTax_New > 0 And (Format(TotAmtNotTaxable + TotAmtToBeTaxed + TotalTax_New, "###.00")) <> Format(TotInvAmt, "###.00") Then
                    'TOTAL = TOTAL & "TOTPAG  " & TotAmtNotTaxable + TotAmtToBeTaxed + TotalTax_New
                    TOTAL = TOTAL & "TOTPAG  " & TotInvAmt & vbNewLine '08/25/23  -- I think it should be always totinvamt,  just comment above for now pending further testing.
                Else
                    TOTAL = TOTAL & "TOTPAG  " & TotInvAmt & vbNewLine
                End If

                'SQL = "select exported,moddate, moduser from sohdr where ordstateid = 'R' and (exported = 'N'  or exported = 'C' ) order by ordernum for update"

                Exported = "N"
                If Not SorianaErr Then Exported = "Y"


                SQL = "Update sohdr set moduser = 'DigInv3', moddate = '" + Format(Now, "dd-MMM-yy") + "', exported = '" + Exported + "' where ordstateid = 'R' and (exported = 'N'  or exported = 'C' ) and ordernum = '" + DR("ordernum") + "'"

                If counter = 0 Then
                    OTR = conn.BeginTransaction(IsolationLevel.ReadCommitted)
                Else
                    OTR = conn.BeginTransaction()
                End If
                CMOrdUpd = New OracleCommand(SQL, conn, OTR)
                CMOrdUpd.ExecuteNonQuery()
                OTR.Commit()
                counter += 1


                'OTR = conn.BeginTransaction(IsolationLevel.ReadCommitted)
                'CMOrdUpd.Transaction = OTR
                'CMOrdUpd = New OracleCommand(SQL, conn, OTR)
                'CMOrdUpd.ExecuteNonQuery()
                'OTR.Commit()
                'RSOrdUpd.edit()
                'If Not SorianaErr Then
                '    RSOrdUpd!exported.Value = "Y" '08/23/17 ADDED IF
                'End If
                'RSOrdUpd!moduser.Value = "DigInv3"
                'RSOrdUpd!moddate.Value = Format(Now, "MM/dd/yyyy HH:mm:ss")
                'RSOrdUpd.Update()
                'RSOrdUpd.movenext()

                If DocName = "NOTA DE CARGO" Then HDRCP = vbNullString 'Nota de cargo does no have item informaion, for the moment we do no generate carta porte.
                If Left(DR("dockid"), 1) <> "F" Then HDRCP = vbNullString
                If Not SorianaErr Then
                    If DR("exported") <> "C" Then CkFileExists(DirToOutputData, HDR & DTL & TOTAL, "INV-" & Folio & "-" & DR("RFC") & "_" & Format(Now, "yyyyMMdd_HH-mm") & ".txt")
                    If Trim(HDRCP) <> vbNullString Then CkFileExists(DirToOutputData, HDRCP & DTLCPT & DTLCP & COMCP & TOTCP, "CTP-" & DR("invhdrnum") & " - " & Format(Now, "yyyy-MM-dd") & " at " & Format(Now, "HH-mm") & ".txt")
                Else '09/15/17 changed to stop sending e-mail when error
                    'SendEmail "SORIANA INVOICE FOUND:  " & DR("ORDERNUM") & ", BUT NOT FOLIO DE ENTRADA FOUND " & vbNewLine & "THIS ORDER WON'T PROCESS."
                End If
                HDR = vbNullString : DTL = vbNullString : TOTAL = vbNullString : XtraCustomsInfo = vbNullString : TotDiscAmt = 0 : HdrNo_SeparateXmlFlag = False : TotPercentages = 0
                HDRCP = vbNullString : DTLCP = vbNullString : DTLCPT = vbNullString : TOTCP = vbNullString : COMCP = vbNullString : TotWeight = 0
                IEPSnoIVA = 0
                LiquorPresent = False
            Next
            'Se.committrans() 'JBS

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
    Public Function GetDecimalWOZero(ByVal val As Decimal) As String
        Dim result As String = val.ToString("G29")
        Return result
    End Function
    '============================================
    Private Sub GetCPVehiculo(ByRef H As Short, ByRef InvNum As String)
        Dim OCM As OracleCommand
        Dim dtc As DataTable
        Dim dtr As DataRow
        Dim Sql As String = "select c.* from sohdr a, shiproute_info b, truckdriver c where a.shipdate = b.shipdate and a.dockid = b.dockid and b.driverid = c.driverid  and a.ordernum = '" + InvNum + "'"


        Try
            OCM = New OracleCommand(Sql, conn)
            dtc = New DataTable
            dtc.Load(OCM.ExecuteReader)
            'RSc = Db.CreateDynaset(SQL, &H4)
            '============================================
            'TRANSPORTATION...........
            COMCP = "COM_CPT_INIAUTO " & vbNewLine & "COM_CPT_AUT_SCT TPAF02" & vbNewLine
            COMCP = COMCP & "   COM_CPT_AUT_PSCT Permiso no contemplado en el catálogo " & vbNewLine & "COM_CPT_AUT_SUTIPREM1 " & vbNewLine 'CTR004" & vbNewLine
            COMCP = COMCP & "   COM_CPT_AUT_ASEGRESP Qualitas Compañía De Seguros, S.A. de C.V. " & vbNewLine & "COM_CPT_AUT_POLIRESP 0003945047 " & vbNewLine
            COMCP = COMCP & "   COM_CPT_AUT_CONVEH C2" & vbNewLine
            'HERE!!!!!! PLACA.... OTHER READ FROM TABLE LATER..... FOR NOW, DEFAULT VALUES............
            COMCP = COMCP & "   COM_CPT_AUT_PLACAV " & "3901CM " & vbNewLine
            COMCP = COMCP & "   COM_CPT_AUT_ANIOV 2014" & vbNewLine & "COM_CPT_AUT_ASEGCAR TOKIO MARINE CIA DE SEGUROS  " & vbNewLine
            COMCP = COMCP & "   COM_CPT_AUT_POLICAR TLJMX000244800 " & vbNewLine & "COM_CPT_AUT_PRIMSEG 900000" & vbNewLine
            COMCP = COMCP & "COM_CPT_AUT_PESBRU  2 " & vbNewLine 'New for version 3 peso del vehiculo sin mercderia en toneladas.
            '!!!!!!!!   CK HERE WHICH VALUES NEED TO BE CHANGED  !!!!!!!!!!!!!!!
            If LiquorPresent Then
                COMCP = COMCP & "   COM_CPT_AUT_ASEGMED Atlas " & vbNewLine 'Ask for correcr information
                COMCP = COMCP & "   COM_CPT_AUT_POLMED 1010101  " & vbNewLine
            End If


            COMCP = COMCP & "COM_CPT_FINAUTO " & vbNewLine & "COM_CPT_INIFIGTRA " & vbNewLine & "COM_CPT_FIG_TIPFIG 01" & vbNewLine

            If dtc.Rows.Count > 0 Then
                'If there is operator data, use information  17/Jul/2024
                dtr = dtc.Rows(0)
                COMCP = COMCP & "COM_CPT_FIG_RFCFIG " & dtr("RFC") & vbNewLine & "COM_CPT_FIG_NUMLIC " & dtr("LICENSENUM") & vbNewLine
                COMCP = COMCP & "COM_CPT_FIG_NOMFIG " & dtr("FULLNAME") & vbNewLine & "COM_CPT_FINFIGTRA " & vbNewLine & "COM_CPT_FINCPT" & vbNewLine
            Else
                COMCP = COMCP & "COM_CPT_FIG_RFCFIG " & JFCRFC & vbNewLine & "COM_CPT_FIG_NUMLIC " & "680000033715 " & vbNewLine
                COMCP = COMCP & "COM_CPT_FIG_NOMFIG " & "Jose Alberto Salas Aguilar " & vbNewLine & "COM_CPT_FINFIGTRA " & vbNewLine & "COM_CPT_FINCPT" & vbNewLine
            End If
            'Debug.Print HDRCP & DTLCPT & DTLCP & COMCP
            'Debug.Print DTLCP
        Catch ex As Exception

            Throw New Exception(ex.Message)
        End Try
    End Sub

    '=====================================================================
    Private Sub GetDtl(ByRef H As Short, ByRef InvNum As String)
        '=====================================================================
        'Will have (I) if the item is taxable to be printed on invoice
        '12/28/10 change to change qty on detalle, and eaqty goes to extra field just like case qty
        Dim SQL As String
        Dim RSd, RS, RSc, Rsd2 As Object
        Dim c, i, d, j As Short
        Dim PricePlusLiqTax As Double
        Dim IvaAmt, Units, Tst As String
        Dim CustomsInfo, FoundCustomsDefault As Boolean
        Dim IEPSTasa, IvaTasaDtl As Double
        Dim Rsd3 As Object ', NotaCargo As Boolean        '05/21/15 added notacargo to check if MC is present.  MC is the main line created, and it should be in line 1, or it won't work.  can be added tax, ieps
        'Dim SeparateLiqTaxMsg As String  --- 2012 oct
        Dim UnitPrice, UPC, Importe As String 'Importe (Price * Qty)
        Dim TotCustomInfo As Short
        Dim EaImpIvaIeps, CsImpIvaIeps As Double
        Dim FirstPortionOfName As Short
        Dim LineIVA, ItemDesc, TmpTotImp As String ' When notas de cargo & sales tax.
        'TaxableItem As String
        Dim CSPrice, Eaprice, LiqTax2 As Double
        Dim IEPSAmt, IEPSAmt_Prt As String
        Dim CsImporte As Double 'ONLY USED WHEN CS & EA FOR LIQUOR NO DESGLOSE CALCULATION... ON EACH
        Dim DocID, PortName, SQLInsert, DocDate As String '', ItemCatalog As String, UOMCatalog As String
        Dim RSCustom As Object
        Dim SQLCustom As String '07/22/23
        'CARTA PORTE
        Dim TotLines As Short
        Dim DtlWeight As Double
        Dim LiqPresentDtl As Boolean
        Dim Peli, CvePeli As String 'Peli Material Peligroso (Dangerous Material)
        Dim IvhdrNo, LineNo As String '04/09/24 Modify for Upgrade in .NET


        '5/1/2024 upgrade for vb.net
        Dim OCM As OracleCommand ' 04/25/2024 add for vb.net upgrade
        Dim OTR As OracleTransaction '04/25/2024 add for vb.net upgrade
        Dim dt, dtc As DataTable '04/25/2024 add for vb.net upgrade
        Dim ItemCode As String
        Dim ODR As OracleDataReader
        Dim DRc As DataRow
        Dim Adpt As OracleDataAdapter '05/22/2024 add for vb.net upgrade

        'Peli = si if branch_item.liquorcode = 3; Peli = 0 if branch_item <> 3, but item catalog is liquor
        '------------
        On Error GoTo ErrHndlr
        '------------------------08/21/18----------SIR FOR CITY FRESKO DIGITAL INVOICE IN EA ONLY -------------------
        If RFC = "CCF121101KQ4" Then 'carta porte added gross and ne weight to sql 2021-DECEMBER-
            SQL = "select i.unitupccode,i.jancode,b.liqtax,nvl(b.liqtax2,0)liqtax2,b.eapercs,b.saltaxcode,b.eapercs, nvl(GROSSWEIGHT,0)  gross, round(nvl(GROSSWEIGHT,0) /2.205 ,2) grosskg ,nvl(netweight,0) net, " & vbNewLine
            SQL = SQL & "b.mx_sat_catalog_id as satItem, b.mx_unitofmeasure as satUOM, nvl(b.liquorcode,0) liquorcode, " & vbNewLine
            SQL = SQL & "D.INVHDRNUM, D.LINENUM, D.ITEMCODE, D.REMARKID, D.LINEDESC, (D.CSSHIPPED * B.EAPERCS + D.EASHIPPED) AS EASHIPPED, 0 AS CSSHIPPED," & vbNewLine
            '052524 JBS An: escape from db Error, OCI-22053: overflow error
            'SQL = SQL & "D.INVCSPRICE / B.EAPERCS AS INVEAPRICE, D.LINETAX, " & vbNewLine
            SQL = SQL & "ROUND(d.INVCSPRICE / B.EAPERCS * 100000000) / 100000000 As INVEAPRICE, D.LINETAX, " & vbNewLine
            SQL = SQL & "D.LINEPRICE,D.LINECOST,D.LINEREDEM,D.LINETOTAL,D.INVCSPRICE,D.LINELIQTAX, D.LINESALESTAX "
            SQL = SQL & "from invdtl d, jfcitem i, branch_item b where invhdrnum = '" & InvNum & "' " & vbNewLine
            SQL = SQL & "and d.itemcode = i.itemcode and i.itemcode = b.itemcode and (eashipped > 0 or csshipped > 0)order by d.itemcode "
        Else 'carta porte added gross and ne weight to sql 2021-DECEMBER- added liquorcode 12/28/23
            SQL = "select i.unitupccode,i.jancode,b.liqtax,nvl(b.liqtax2,0)liqtax2,b.eapercs,b.saltaxcode,b.eapercs, nvl(GROSSWEIGHT,0)  gross, round(nvl(GROSSWEIGHT,0) /2.205 ,2) grosskg ,nvl(netweight,0) net, " & vbNewLine
            SQL = SQL & "b.mx_sat_catalog_id as satItem, b.mx_unitofmeasure as satUOM, nvl(b.liquorcode,0) liquorcode, " & vbNewLine
            SQL = SQL & "d.* from invdtl d, jfcitem i, branch_item b where invhdrnum = '" & InvNum & "' " & vbNewLine
            SQL = SQL & "and d.itemcode = i.itemcode and i.itemcode = b.itemcode and (eashipped > 0 or csshipped > 0)order by d.itemcode "
        End If

        'Debug.Print SQL
        TotAmtToBeTaxed = 0 'Used to print next to iva msg to know total amt of taxable items. '12/28/10

        '5/1/2024 upgrade for vb.net
        'RS = Db.CreateDynaset(SQL, &H4)

        OCM = New OracleCommand(SQL, conn)
        dt = New DataTable
        'dt.Load(OCM.ExecuteReader) '052224 JBS An: different way to read data for better performance
        Adpt = New OracleDataAdapter(OCM)
        Adpt.Fill(dt)

        If dt.Rows.Count > 0 Then
            d = 0 ' If CS and EA, we separate into two lines, so we can't use i as detail line number.
            For Each DR As DataRow In dt.Rows
                ItemCode = DR("ItemCode")
                DiscAmt = 0 : DiscExists = False : DiscPer = 0 '02/09/18 Clear up discounts when free items
                IEPSTasa = Val(vbNullString & DR("liqtax")) 'added vbnullstring 12/10/14
                LiqTax2 = Val(vbNullString & DR("LiqTax2"))
                If IEPSTasa <> 0 Then IEPSTasaGlobal = IEPSTasa ' if stmt to getlast ieps tasa in the ord. When last item did not have tasa, was upadating tasa global to 0.
                IEPSAmt = vbNullString : LineIVA = vbNullString : IvaAmt = vbNullString 'use when ea and CS exists to properly calculate line iva (in DB amt is CS and EA combined)
                IEPSAmt_Prt = vbNullString 'added on 05/23/13 bug detected by JMX users... was not being cleared...
                If RFC = "NWM9709244W4" Then
                    UPC = ("" & DR("jancode")) 'added on 02/28/13 for WM use Jancode first, if not found, use UPC.
                    If Val(UPC) = 0 Then UPC = ("" & DR("unitupccode"))
                    'If LiqTax2 > 0 Then IEPSTasa = 0.0001 'Added per MasterEDI on 08/09/17 !!!!!!!!! for WM, but still it did not work, commented out for version 3.3
                Else 'Other than WM check upc first, if null, then look for Jan code
                    UPC = ("" & DR("unitupccode")) '12/06/10
                    If Val(UPC) = 0 Then UPC = ("" & DR("jancode"))
                End If
                If Val(UPC) = 0 Then UPC = vbNullString 'if both upc and jan code are 0, then don't print all zeroes, leave it blank
                If RFC = "NWM9709244W4" Then UPC = Format(Val(UPC), "0000000000000") 'WM wants 13 digits field. ---WAL-MART
                '052124, JBS An: remove space by endtrim 
                'ItemDesc = vbNullString & DR("linedesc") '2012 I need just desc for Soriana
                ItemDesc = vbNullString & DR("linedesc").ToString().TrimEnd() '2012 I need just desc for Soriana
                ItemCatalog = vbNullString & DR("satItem") '3.3 version
                UOMCatalog = vbNullString & DR("satuom") '3.3 version
                '!!!!!!!!!!!!INVALID ITEM CATALOGS !!!!!!!!!!!!!!!!
                If Trim(ItemCatalog) = vbNullString Then ItemCatalog = "50171500" 'defatult item catalog !!! !!!!!!!!!!!!!
                If Trim(ItemCatalog) = "50424800" Then ItemCatalog = "50171500" 'THIS ITEM CATALOG DOES NOT EXIST-- INVALID!!!
                If Trim(ItemCatalog) = "50347000" Then ItemCatalog = "50171500" 'THIS ITEM CATALOG DOES NOT EXIST-- INVALID!!!
                '!!!!!!!PRODUCTO PELIGROSO!!!!!!!!! HERE!!!!!!!!!!!!!!!!!
                '------MATERIAL PELIGROSO---Jan 2023 CP v 3.0 material peligroso (licor) puede ser Si o No.  If "Si" we need dangerous product info. If "No", we do not put the info
                Peli = vbNullString : CvePeli = vbNullString : LiqPresentDtl = False '01/23/24 added  LiqPresentDtl = false 02/07/24
                Select Case Trim(ItemCatalog)
                    Case "50202200", "50202201", "50202206", "50202210", "50121500", "50171708" 'Pescado, licor y vino, 50171708 added 02/07/24
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
                '--------------------------------------------
                d = d + 1
                Eaprice = DR("INVEAPRICE")
                CSPrice = DR("INVCSPRICE")
                If DR("eashipped") > 0 And DR("csshipped") > 0 Then
                    Units = "CE"
                ElseIf DR("eashipped") > 0 Then
                    Units = "EA"
                    UOMCatalog = "H87" '06/21/18 JMX NOW WANTS TO CHANGE UNIT OF MEASURE TO BE 'XBX' IF CASES, OR 'H87' IF EACHES.
                Else 'Either just cases, or no cases and no each.  When xtra charges in detail (AC, tax, etc)
                    Units = "CA"
                    UOMCatalog = "XBX" '06/21/18 JMX NOW WANTS TO CHANGE UNIT OF MEASURE TO BE 'XBX' IF CASES, OR 'H87' IF EACHES.
                End If
                If DocName = "NOTA DE CARGO" Then '10/03/18 PER MA ELENA'S EMAIL ON 10/03/18 SIR NOT YET SUBMITTED
                    UOMCatalog = "ACT"
                    ItemCatalog = "84111506"
                End If
                '---------
                If DR("linetax") - DR("lineliqtax") > 0 Then 'LineTax has combined tax (iva & ieps), to get iva, substract both taxes.
                    IvaTasaDtl = IVATasa
                    TotAmtToBeTaxed = Val(CStr(TotAmtToBeTaxed)) + Val(DR("lineprice")) + Val(DR("lineliqtax")) 'SUBTAI
                    TotalTax = CStr(Val(TotalTax) + (Val(DR("linetax")) - Val(DR("lineliqtax")))) 'calculate totaltax based on detail
                Else
                    IvaTasaDtl = 0
                    TotAmtNotTaxable = Val(CStr(TotAmtNotTaxable)) + Val(DR("lineprice")) 'SUBTSI
                End If

                IvhdrNo = DR("invhdrnum") '04/09/24 Modify for Upgrade in .NET
                LineNo = DR("linenum") '04/09/24 Modify for Upgrade in .NET
                SeparateLiqTaxFlag = IepsSeparate(IvhdrNo, LineNo) '04/09/24 Modify for Upgrade in .NET

                '07/14/23 comment out next line, already taking care of the flag NO_separateXML
                'If SeparateLiqTaxFlag = "Y" Or (SeparateLiqTaxFlag = "N" And Val(IEPSTasa) = 0) Then IEPS_NO_SeparateXml = False  '04/23/19 SIR 1794 separate and do not separate should not be togheter.  One flag is detail, the other one is global.
                '-----
                TotAmtBeforeTaxes = TotAmtBeforeTaxes + DR("lineprice") '01/23/18
                '07/14/23 comment out this line.... I think we don't need this line any more
                'If Val(IEPSTasa) > 0 And IEPS_NO_SeparateXml Then CkifInArr Val(IEPSTasa), IEPS_NO_SeparateXml   '---->>>>07/26/1'--->>>>>>>>>>>>>>
                '-----------------------------------------------------
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
                        If (SeparateLiqTaxFlag = "N" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "N" And LiqTax2 > 0) Then 'Price should add IEPS on inv to be sent on EXTRA FIELDS' 12/23/13 ----------
                            If IEPSTasa > 0 Then UnitPrice = Format(Eaprice + (Eaprice * (IEPSTasa / 100)), "0.00###") '01/17/18 added one more zero because it doesn't pass SAT checking if round up here
                            'PENDING-----------------------07/14/23----------------------------------------------
                            'NEED TO REPEAT THIS SAME MOD TO THE OTHER SECTIONS (CA, CA & EA) -------------
                            Importe = Format(Val(DR("lineprice")) + Val(DR("lineliqtax")), "##0.00") '052124 JBS An: corrected decimal format. 
                            If LiqTax2 > 0 Then UnitPrice = CStr(Val(Importe) / DR("eashipped")) '07/14/23 use this instead of bellow -- was off some cents per eaprice so total was not matching....
                            'If LiqTax2 > 0 Then UnitPrice = Format(Eaprice + (LiqTax2 / DR("eapercs")), "##0.00###") ' '01/17/18 added one more zero because it doesn't pass SAT checking if round up herE
                            IEPSAmt_Prt = vbNullString
                        End If
                        If IvaTasaDtl > 0 Then
                            LineIVA = CStr(Val(DR("linetax")) - Val(DR("lineliqtax")))
                            TotalTax_New = TotalTax_New + CDbl(LineIVA) '----04/12/19 sir 1794, when ea and cs present one cent difference and it doesn't generate cfdi
                        End If
                        TotImp_Prt = TotImp_Prt + Val(Importe)
                        '052124 JBS An: no need trim here. 
                        'DTL = DTL & "CANTID  " & DR("eashipped") & vbNewLine & "CANTID_EA  " & DR("eashipped") & vbNewLine & "DESCRI  " & ItemDesc.ToString.Trim & vbNewLine
                        DTL = DTL & "CANTID  " & DR("eashipped") & vbNewLine & "CANTID_EA  " & DR("eashipped") & vbNewLine & "DESCRI  " & ItemDesc & vbNewLine
                        DTL = DTL & "CANPAQ  " & DR("eashipped") & vbNewLine
                        DTL = DTL & "CANEMP  " & DR("eashipped") & vbNewLine & "UNIDAD  " & Units & vbNewLine & "CVESKU  " & DR("itemcode") & vbNewLine
                        DTL = DTL & "ESTILV     " & DR("itemcode") & vbNewLine '10/31/17 per Master EDI to be able to print item code on pdf.
                        DTL = DTL & "CVEPRODSERV     " & ItemCatalog & vbNewLine '10/31/17 New SAT item code
                        DTL = DTL & "CVEUNIDAD           " & UOMCatalog & vbNewLine '11/15/17 New for 3.3 version
                        DTL = DTL & "CODUPC  " & UPC & vbNewLine & "PIEPEM  " & "1" & vbNewLine & "PIEPEM2 " & DR("eapercs") & vbNewLine & "CODDUN  " & vbNewLine
                        '************************NOVEMBER  2021
                        'HERE!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                        DTLCP = DTLCP & "      COM_CPT_INIMER " & vbNewLine & vbNewLine 'Inicio de mercancia
                        DTLCP = DTLCP & "      COM_CPT_MER_BIENTRA " & ItemCatalog & vbNewLine 'BienesTransp (clave de producto)
                        DTLCP = DTLCP & "      COM_CPT_MER_DESCRI " & ItemDesc & vbNewLine
                        DTLCP = DTLCP & "      COM_CPT_MER_CANTID 1" & vbNewLine
                        DTLCP = DTLCP & "      COM_CPT_MER_CVUNID  " & UOMCatalog & vbNewLine '
                        '052424 JBS An: fixed double round up issue,
                        DtlWeight = RoundUpToDecimals((DR("eashipped") / DR("eapercs") * DR("grosskg")), 2)
                        If DtlWeight = 0 Then DtlWeight = 0.01
                        TotWeight = TotWeight + DtlWeight
                        DTLCP = DTLCP & "      COM_CPT_MER_PKG " & DtlWeight & vbNewLine 'Format((DR("eashipped / DR("eapercs) * DR("grosskg, "###0.00000") & vbNewLine & vbNewLine      'itemweight e  PesoEnKg

                        'HERE!!!!!!!!!  FOR PRODUCTO PELIGROSO WE NEED TO READ A TABLE FOR CLAVEMATERIALPELIGROSO.....!!!!!!!!!!!!!!!!
                        If LiqPresentDtl Then
                            DTLCP = DTLCP & "      COM_CPT_MER_MATPEL Sí  " & vbNewLine 'Material peligroso
                            DTLCP = DTLCP & "      COM_CPT_MER_CVEMATPEL   " & CvePeli & vbNewLine '3065  drinks 24% pero no más de 70%  alcohol , or 1011 gas butano
                            DTLCP = DTLCP & "      COM_CPT_MER_EMB 4G" & vbNewLine '01/26/24  4C1 modified to 4G per Ma Elena
                        End If
                        If Peli = "No" Then
                            DTLCP = DTLCP & "      COM_CPT_MER_MATPEL No " & vbNewLine '01/23/24
                        End If
                        DTLCP = DTLCP & "      COM_CPT_INICANTRAN " & vbNewLine 'Inicio canidad trasladada
                        DTLCP = DTLCP & "         COM_CPT_CMER_CANTID " & DR("eashipped") & vbNewLine
                        DTLCP = DTLCP & "         COM_CPT_CMER_IDORI OR000001" & vbNewLine
                        DTLCP = DTLCP & "         COM_CPT_CMER_IDDES DE000001" & vbNewLine
                        DTLCP = DTLCP & "      COM_CPT_FINCANTRAN" & vbNewLine
                        DTLCP = DTLCP & "COM_CPT_FINMER" & vbNewLine & vbNewLine
                        '----------------------CHANGES FOR FREE PRODUCT-----------------02/09/18
                        If Eaprice = 0 Then 'free
                            Eaprice = 0.01
                            DiscExists = True
                            DiscPer = 100
                            'VERSION 4 DO NOT ADD THE NOT TAXABLE TO FREE PRODUC- INVOICE WILL ERROR OUT COMMENT THESE TWO LINES
                            'TotAmtNotTaxable = Val(TotAmtNotTaxable) + (0.01 * DR("eashipped"))
                            'TotAmtBeforeTaxes = TotAmtBeforeTaxes + (0.01 * DR("eashipped"))
                            'END VERSION 4 MOD
                            DiscAmt = 0.01 * DR("eashipped") 'DO NOT Accumulate total discount DiscAmt + (0.01 * DR("EASHIPPED"))
                            IEPSTasa = 0 'In case free item has ieps do not report tasa because it won't have any ieps amount
                            LiqTax2 = 0 '01/25/24  In case free item, do not report any tasa, no ieps amt should be reported
                            TotDiscAmt = TotDiscAmt + DiscAmt
                        End If
                        '--------------------------------------------------------------------------------END OF 02/19/18
                        If RFC <> "PHI830429MG6" Then '-----------------Palacio de Hierro doesn't want 0 when iva is 0 '04/20/18
                            If Trim(LineIVA) = vbNullString Then LineIVA = "0.00" : If Trim(IEPSAmt) = vbNullString Then IEPSAmt = "0.00" '04/20/18
                        End If
                        '---------------------------------------------------------------------------------------------
                        '-----------VERSION 4.0 - NO TASIPE WHEN 0, NO MONIPE   WHEN 0,  NO TASIEP  when 0,  NO MONIEP WHEN 0
                        '------------------------------------------------------------------------------------------------------
                        If Val(CStr(IvaTasaDtl)) > 0 Then 'VERSION 4.0 ADDED IF TO JUST SEND THESE WHEN > 0  -  MAY 2022
                            DTL = DTL & "TASIPE  " & IvaTasaDtl & vbNewLine
                            DTL = DTL & "MONIPE  " & LineIVA & vbNewLine
                        End If
                        '--------------------VERSION 4.0 - OBJECTO DE IMPUESTO ------------
                        If Val(CStr(IEPSTasa)) > 0 Or LiqTax2 > 0 Then
                            If SeparateLiqTaxFlag = "N" Then
                                DTL = DTL & "OBJIMP 02" & vbNewLine 'TEMP CHANGED EVERYTHING TO 02
                            Else
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            End If
                            If Val(CStr(IvaTasaDtl)) = 0 Then DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine '08/29/23 calpico issue
                        Else
                            If Val(CStr(IvaTasaDtl)) > 0 Then
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            Else
                                If DiscExists Then
                                    DTL = DTL & "OBJIMP 01" & vbNewLine '05/27/22 WHEN FREE ITEMS NO IVA SHOULD BE REPORTED
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
                            DTL = DTL & "TIPIPETR Tasa" & vbNewLine '------------& "TASIEP 0" & vbNewLine
                            TotAmtNotTaxable = TotAmtNotTaxable + CDbl(IEPSAmt)
                        End If
                        '---------------END VERSION 4.............

                        DTL = DTL & "MONIEP_IEPS  " & IEPSAmt_Prt & vbNewLine
                        DTL = DTL & "IMPIVAIEPS  " & CDbl(DR("linetotal")) & vbNewLine
                        DTL = DTL & "PBRUDE_IEPS  " & Format(CDbl(UnitPrice), "##0.00") & vbNewLine '01/02/18 '052124 JBS An: corrected decimal format
                        DTL = DTL & "IMPBRU_PRT  " & Importe & vbNewLine '01/02/18
                        '********************************************SIR 1794********************************************************
                        If ((SeparateLiqTaxFlag = "N" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "N" And LiqTax2 > 0)) And IEPS_NO_SeparateXml Then '-->>>04/09/19 ****SIR # 1794*** ------------->>>>>>>>>>>
                            'DTL = DTL & "TASIEP  " & "0" & vbNewLine & "MONIEP  " & "0" & vbNewLine  '---VERSION 4.0 COMMENT THIS OUT NO ZERO VALUE
                            DTL = DTL & "PBRUDE  " & Format(CDbl(UnitPrice), "##0.00") & vbNewLine '052124 JBS An: Added "CE format. 
                            DTL = DTL & "VALUNI  " & Format(CDbl(UnitPrice), "##0.00") & vbNewLine '052124 JBS An: Added decimal format. 
                            DTL = DTL & "IMPBRU  " & Format(CDbl(UnitPrice) * DR("eashipped"), "##0.00") & vbNewLine '052124 JBS An: Added decimal format. 
                            DTL = DTL & "IMPORT  " & Format(CDbl(UnitPrice) * DR("eashipped"), "##0.00") & vbNewLine '052124 JBS An: Added decimal format. 
                            '--VERSION 4  ---  NOT SURE IF NEED TO USE OBJIMP = 03 ---- WILL CHECK THIS LATER------------
                        Else 'SIR 1794 bellow is same, what changes is above.
                            If Val(CStr(IEPSTasa)) > 0 Then DTL = DTL & "TASIEP  " & IEPSTasa & vbNewLine & "MONIEP  " & CDbl(IEPSAmt) & vbNewLine 'VERSION 4.0 ADDED IF ONLY IF >0
                            DTL = DTL & "PBRUDE  " & Format(Eaprice, "##0.00") & vbNewLine '052124 JBS An: Added decimal format. 
                            DTL = DTL & "VALUNI  " & Format(Eaprice, "##0.00") & vbNewLine '052124 JBS An: Added decimal format. 
                            DTL = DTL & "IMPBRU  " & Format(Eaprice * DR("eashipped"), "##0.00") & vbNewLine '052124 JBS An: Added decimal format. 
                            DTL = DTL & "IMPORT  " & Format(Eaprice * DR("eashipped"), "##0.00") & vbNewLine '052124 JBS An: Added decimal format. 
                        End If 'SIR 1794
                        '********END OF MOD*******************SIR 1794********************************************************
                        If LiqTax2 > 0 Then '01/23/18---------SUGARY DRINKS!!!!!!!!!!------------------
                            'if DiscAmt = 0 Then ' 02/09/18   don't do this bellow for free items COMMENT OUT ON 07/14/23 REPLACED BY BELLOW
                            If DiscAmt = 0 And SeparateLiqTaxFlag = "Y" Then ' don't do this bellow for free items 07/14/23 ADDED SEPARATELIQTAXFLAG
                                DTL = DTL & "FCTTASIEP   " & IEPSperLiter & vbNewLine '05/09/17  'DTL = DTL & "FCTTASIEP   1.000000" & vbNewLine '
                                DTL = DTL & "TIPIEPTR    Cuota" & vbNewLine
                                'DTL = DTL & "IMPORTIEP    " & IEPSAmt & vbNewLine 'IMPORTIEP is the base (total liters per line), or iepsamt/1.17
                                'DTL = DTL & "IMPORTIEP    " & Format(Val(IEPSAmt) / IEPSperLiter, "0.######") & vbNewLine
                                DTL = DTL & "IMPORTIEP    " & Format(Val(IEPSAmt) / IEPSperLiter, "0.##") & vbNewLine 'VERSION 4 USE ONLY TWO DIGITS, OTHERWISE ERRORS OUT
                                'VERSION 4 ADD THE FOLLOWING..... WE NEED TO SPECITY THE IVA WHEN IT IS 0% like in SOME OF the sugar products which have IEPS but 0% IVA
                                DTL = DTL & "MONIEP " & CDbl(Val(IEPSAmt)) & vbNewLine
                                ' If LineIVA = 0 Then '05/22/22 some 1.17 cuota also have iva, so do not report 0% if iva exists.
                                'If Val(LineIVA) = 0 Then '05/22/22 some 1.17 cuota also have iva, so do not report 0% if iva exists. ----01/11/23  added Val, bug detected inv 0981304
                                If Val(CStr(IvaTasaDtl)) = 0 Then '01/26/23  the modificatin above, should actually be this one  ivatasadtl not lineiva.....
                                    DTL = DTL & "TASIPE 0 " & vbNewLine & "MONIPE 0" & vbNewLine & "FCTTASIPE 0.000000" & vbNewLine
                                    DTL = DTL & "TIPIPETR Tasa " & vbNewLine & "IMPORTIPE " & ((Eaprice * DR("eashipped"))) + CDbl(IEPSAmt) & vbNewLine
                                    IEPSnoIVA = IEPSnoIVA + CDbl(IEPSAmt)
                                    'TotAmtNotTaxable = TotAmtNotTaxable + IEPSAmt 'I think I don'tneed this, ck.........01/26/23 comment this out
                                Else

                                End If
                                '--------------END OF ADDITON FOR VERSION 4--------------------------------
                            End If
                        End If
                        If DiscExists Then
                            DTL = DTL & "TDECON    " & DiscPer & vbNewLine
                            DTL = DTL & "MDECON    " & (Eaprice * DR("eashipped")) * (DiscPer / 100) & vbNewLine
                        Else
                            DTL = DTL & "TDECON  " & "0" & vbNewLine & "MDECON  " & "0.00" & vbNewLine 'NO DESCUENTO EN DETALLE HARD CODE 0
                        End If

                        DTL = DTL & "NUMLIN    " & d & vbNewLine ''-------------01/05/18  I am placing at end  ----  EA -----3
                        TotIEPSAmt = TotIEPSAmt + Val(IEPSAmt) '
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
                            Importe = Format(Val(DR("lineprice")) + Val(DR("lineliqtax")), "##0.00") 'Importe con IEPS '052124 JBS An: corrected decimal format
                            IEPSAmt_Prt = vbNullString
                        End If
                        If Val(Importe) = 0 And TotInvAmt <> 0 Then 'when eaqty and csqty = 0, but detail exists, probably XTRA CHARGE!!!
                            Importe = Format(CDbl(DR("lineprice")), "##0.00") 'ONLY WHEN EAQTY AND CSQTY = 0 WE WILL HAVE THIS SITUATION '052124 JBS An: corrected decimal format. 
                            UnitPrice = "0.00"
                        End If
                        If IvaTasaDtl > 0 Then
                            LineIVA = CStr(Val(DR("linetax")) - Val(DR("lineliqtax")))
                            TotalTax_New = TotalTax_New + CDbl(LineIVA) '----04/12/19 sir 1794, when ea and cs present one cent difference and it doesn't generate cfdi
                        End If
                        TotImp_Prt = TotImp_Prt + Val(Importe) 'Total importe a imprimir (Accumulate accordingly for printing pursposes which is diff. for SAT purposes)
                        DTL = DTL & "D" & vbNewLine
                        DTL = DTL & "CANTID  " & DR("csshipped") & vbNewLine & "CANTID_CA  " & DR("csshipped") & vbNewLine & "DESCRI  " & ItemDesc.ToString.Trim & vbNewLine
                        DTL = DTL & "CANPAQ  " & DR("csshipped") & vbNewLine
                        DTL = DTL & "CANEMP  " & DR("csshipped") & vbNewLine & "UNIDAD  " & Units & vbNewLine & "CVESKU  " & DR("itemcode") & vbNewLine
                        DTL = DTL & "ESTILV     " & DR("itemcode") & vbNewLine '10/31/17 per Master EDI to be able to print item code on pdf.
                        DTL = DTL & "CVEPRODSERV     " & ItemCatalog & vbNewLine '10/31/17 New SAT item code
                        DTL = DTL & "CVEUNIDAD     " & UOMCatalog & vbNewLine '11/15/17 New for 3.3 version
                        DTL = DTL & "CODUPC  " & UPC & vbNewLine & "PIEPEM  " & DR("eapercs") & vbNewLine & "PIEPEM2 " & DR("eapercs") & vbNewLine & "CODDUN  " & vbNewLine
                        '*****************OCTOBER 2021*****************************
                        DTLCP = DTLCP & "COM_CPT_INIMER " & vbNewLine & vbNewLine 'Inicio de mercancia
                        DTLCP = DTLCP & "   COM_CPT_MER_BIENTRA " & ItemCatalog & vbNewLine 'BienesTransp (clave de producto)
                        DTLCP = DTLCP & "   COM_CPT_MER_DESCRI " & ItemDesc & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_MER_CANTID 1" & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_MER_CVUNID  " & UOMCatalog & vbNewLine '
                        '052424 JBS An: fixed double round up issue,  
                        'DtlWeight = CDbl(Format(RoundUpToDecimals(DR("csshipped") * DR("grosskg"), 2), "###0.00"))
                        DtlWeight = RoundUpToDecimals(DR("csshipped") * DR("grosskg"), 2)
                        If DtlWeight = 0 Then DtlWeight = 0.01
                        TotWeight = TotWeight + DtlWeight
                        DTLCP = DTLCP & "   COM_CPT_MER_PKG " & DtlWeight & vbNewLine & vbNewLine 'itemweight e  PesoEnKg
                        If LiqPresentDtl Then
                            DTLCP = DTLCP & "      COM_CPT_MER_MATPEL Sí  " & vbNewLine 'Material peligroso
                            DTLCP = DTLCP & "      COM_CPT_MER_CVEMATPEL   3065  " & vbNewLine ' BEBIDAS ALCOHOLICAS, 24% pero no más de 70% de alcohol en volumen
                            DTLCP = DTLCP & "      COM_CPT_MER_EMB 4C1" & vbNewLine 'NEED TO KNOW THE EMBALAJE CORRECTO !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                        End If
                        If Peli = "No" Then
                            DTLCP = DTLCP & "      COM_CPT_MER_MATPEL No " & vbNewLine '01/23/24
                        End If
                        'If Peli = vbNullString Then DTLCP = DTLCP & "COM_CPT_MER_MATPEL " & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_INICANTRAN " & vbNewLine 'Inicio canidad trasladada
                        DTLCP = DTLCP & "   COM_CPT_CMER_CANTID " & DR("csshipped") & vbNewLine
                        DTLCP = DTLCP & "      COM_CPT_CMER_IDORI OR000001" & vbNewLine
                        DTLCP = DTLCP & "     COM_CPT_CMER_IDDES DE000001" & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_FINCANTRAN" & vbNewLine
                        DTLCP = DTLCP & "COM_CPT_FINMER" & vbNewLine & vbNewLine

                        '----------------------CHANGES FOR FREE PRODUCT-----------------02/09/18
                        If CSPrice = 0 Then 'free
                            CSPrice = 0.01
                            DiscExists = True
                            DiscPer = 100
                            'VERSION 4 DO NOT ADD INVOICE WILL ERROR OUT -- COMMENT THESE TWO LINES
                            'TotAmtNotTaxable = Val(TotAmtNotTaxable) + (0.01 * DR("csshipped"))
                            'TotAmtBeforeTaxes = TotAmtBeforeTaxes + (0.01 * DR("csshipped"))
                            'END VERSION 4 MOD
                            DiscAmt = 0.01 * DR("csshipped")
                            IEPSTasa = 0 'In case free item has ieps do not report tasa because it won't have any ieps amount
                            LiqTax2 = 0 '01/25/24  In case free item, do not report any tasa, no ieps amt should be reported
                            TotDiscAmt = TotDiscAmt + DiscAmt
                        End If
                        '--------------------------------------------------------------------------------END OF 02/19/18
                        If RFC <> "PHI830429MG6" Then '-----------------Palacio de Hierro doesn't want 0 when iva is 0 '04/20/18
                            If Trim(LineIVA) = vbNullString Then LineIVA = "0" : If Trim(IEPSAmt) = vbNullString Then IEPSAmt = "0" '04/20/18
                        End If '04/20/18
                        '---------------------------------------------------------------------------------------------
                        '-----------VERSION 4.0 - NO TASIPE WHEN 0, NO MONIPE   WHEN 0,  NO TASIEP  when 0,  NO MONIEP WHEN 0
                        '------------------------------------------------------------------------------------------------------
                        If Val(CStr(IvaTasaDtl)) > 0 Then
                            DTL = DTL & "TASIPE  " & IvaTasaDtl & vbNewLine & "MONIPE  " & LineIVA & vbNewLine 'VERSION 4 MAY 2022 ADDED IF
                        Else '08/29/23 added else calpico issue
                            DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine
                        End If
                        '--------------------VERSION 4.0 - OBJECTO DE IMPUESTO ------------
                        If Val(CStr(IEPSTasa)) > 0 Or LiqTax2 > 0 Then '5/11/22 added LiqTax2 > 0 for the sugar drink when no separate was using objimp 02
                            If SeparateLiqTaxFlag = "N" Then
                                If DiscExists Then
                                    DTL = DTL & "OBJIMP 01" & vbNewLine '10/24/22 don't know why liqtax2 > 0 even though this is free product
                                Else
                                    DTL = DTL & "OBJIMP 02" & vbNewLine 'temporary everything 02
                                End If
                            Else
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            End If
                        Else
                            If Val(CStr(IvaTasaDtl)) > 0 Then
                                If SeparateLiqTaxFlag = "N" And Val(CStr(IEPSTasa)) > 0 Then
                                    DTL = DTL & "OBJIMP 02" & vbNewLine 'Temporary everything 02
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
                            DTL = DTL & "TIPIPETR Tasa" & vbNewLine '-----------------------& "TASIEP 0" & vbNewLine
                            TotAmtNotTaxable = TotAmtNotTaxable + CDbl(IEPSAmt)
                            'Else '08/25/23  --------------DELETE SINCE I AM NOT ADDING IT HERE----------------
                            'DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine
                        End If

                        '---------------------------------------------------------------------------------

                        DTL = DTL & "MONIEP_IEPS  " & IEPSAmt_Prt & vbNewLine
                        DTL = DTL & "IMPIVAIEPS  " & CDbl(DR("linetotal")) & vbNewLine
                        DTL = DTL & "PBRUDE_IEPS  " & UnitPrice & vbNewLine '01/02/18
                        DTL = DTL & "IMPBRU_PRT  " & Importe & vbNewLine '01/02/18
                        '********************************************SIR 1794********************************************************
                        If ((SeparateLiqTaxFlag = "N" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "N" And LiqTax2 > 0)) And IEPS_NO_SeparateXml Then '-->>>04/09/19 ****SIR # 1794*** ------------->>>>>>>>>>>
                            ' DTL = DTL & "TASIEP  " & "0" & vbNewLine & "MONIEP  " & "0" & vbNewLine 'VERSION 4.0  NO MORE ZERO VALUE COMMENT
                            DTL = DTL & "PBRUDE  " & Format(CDbl(UnitPrice), "##0.00") & vbNewLine '052124 JBS An: corrected decimal format.
                            DTL = DTL & "VALUNI  " & Format(CDbl(UnitPrice), "##0.00") & vbNewLine '052124 JBS An: corrected decimal format.
                            DTL = DTL & "IMPBRU  " & Format(CDbl(UnitPrice) * DR("csshipped"), "##0.00") & vbNewLine '052124 JBS An: corrected decimal format.
                            DTL = DTL & "IMPORT  " & Format(CDbl(UnitPrice) * DR("csshipped"), "##0.00") & vbNewLine '052124 JBS An: corrected decimal format.
                        Else
                            If Val(CStr(IEPSTasa)) > 0 Then DTL = DTL & "TASIEP  " & IEPSTasa & vbNewLine & "MONIEP  " & CDbl(IEPSAmt) & vbNewLine 'VERSION 4 ADDED IF ONLY WHEN NON ZERO
                            DTL = DTL & "PBRUDE  " & Format(CSPrice, "##0.00") & vbNewLine '052124 JBS An: corrected decimal format.
                            DTL = DTL & "IMPBRU  " & Format(CSPrice * DR("csshipped"), "##0.00") & vbNewLine '052124 JBS An: corrected decimal format.
                            DTL = DTL & "VALUNI  " & Format(CSPrice, "##0.00") & vbNewLine '052124 JBS An: corrected decimal format.
                            DTL = DTL & "IMPORT  " & Format(CSPrice * DR("csshipped"), "##0.00") & vbNewLine '052124 JBS An: corrected decimal format.
                        End If
                        '---------
                        TotIEPSAmt = TotIEPSAmt + Val(IEPSAmt) '---NOT SURE IF I NEED TO ADD IS NO DESGLOSE......................................???????????????/////
                        If LiqTax2 > 0 Then '01/23/18---------SUGARY DRINKS!!!!!!!!!!  1.17 per liter ------------------
                            ' If DiscAmt = 0 Then '  02/09/18 don't do this bellow for free items  COMMENT OUT 07/14/23 REPLACED BY BELLOW
                            If DiscAmt = 0 And SeparateLiqTaxFlag = "Y" Then '  02/09/18 don't do this bellow for free items'  02/09/18 don't do this bellow for free items AND don't do this if no separate 07/14/23 added
                                DTL = DTL & "FCTTASIEP   " & IEPSperLiter & vbNewLine ''05/09/18    'DTL = DTL & "FCTTASIEP   1.000000" & vbNewLine '
                                DTL = DTL & "TIPIEPTR    Cuota" & vbNewLine
                                'DTL = DTL & "IMPORTIEP    " & Format(Val(IEPSAmt) / IEPSperLiter, "0.######") & vbNewLine    '05/09/18 ADDED /IEPSperLiter
                                DTL = DTL & "IMPORTIEP    " & Format(Val(IEPSAmt) / IEPSperLiter, "0.##") & vbNewLine '05/31/22 VERSION 4, USE ONLY TWO DECIMALS, OTHERWISE ERRORS OUT
                                '*************----------------******************
                                'VERSION 4 ---- TO ADD THE IVA 0% AT DETAIL LEVEL ---------
                                DTL = DTL & "MONIEP " & CDbl(Val(IEPSAmt)) & vbNewLine
                                'If LineIVA = 0 Then '05/22/22 some 1.17 cuota also have iva, so do not report 0% if iva exists. '01/11/23 replaced with bellow... bug detected inv 0981304
                                ''''''''''If Val(LineIVA) = 0 Then '05/22/22 some 1.17 cuota also have iva, so do not report 0% if iva exists. '---01/11/23 added val
                                If Val(CStr(IvaTasaDtl)) = 0 Then '01/26/23  the modificatin above, should actually be this one  ivatasadtl not lineiva.....
                                    DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine & "FCTTASIPE   0.000000" & vbNewLine
                                    DTL = DTL & "IMPORTIPE " & Format(CSPrice * DR("csshipped") + CDbl(IEPSAmt), "##0.00") & vbNewLine 'Base to calculate Tasa 0%
                                    DTL = DTL & "TIPIPETR Tasa" & vbNewLine & "TASIEP 0" & vbNewLine
                                    IEPSnoIVA = IEPSnoIVA + CDbl(IEPSAmt)
                                    'TotAmtNotTaxable = TotAmtNotTaxable + IEPSAmt 'I think I don'tneed this, ck...........01/26/23    comment out.........
                                Else

                                End If
                                'VERSION 4 END ADDITON
                                '**********-------------------------**************
                            End If
                        End If
                        If DiscExists Then '02/09/18
                            DTL = DTL & "TDECON    " & DiscPer & vbNewLine
                            DTL = DTL & "MDECON    " & (CSPrice * DR("csshipped")) * (DiscPer / 100) & vbNewLine '02/16/18 ---> was using EaPrice here, so I changed to CsPrice
                        Else
                            DTL = DTL & "TDECON  " & "0" & vbNewLine & "MDECON  " & "0.00" & vbNewLine 'NO DESCUENTO EN DETALLE HARD CODE 0
                        End If

                        DTL = DTL & "NUMLIN  " & d & vbNewLine '-------------01/05/18  I am placing at end   --------------CA------------------
                        'Else 'CE    CS & EA
                    Case Is = "CE" '----------- CS & EA
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''LiqTax2 = 4.69 '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!test TAKEOUT !!!!!!!!!!!!!
                        CsImpIvaIeps = 0 : EaImpIvaIeps = 0
                        'PRICEPLUSLIQTAX USED TO CALCULATE IVA BECAUSE IVA INCLUDES LIQ TAX ON TOP OF PRICE.
                        PricePlusLiqTax = Val(DR("INVCSPRICE")) + (Val(DR("INVCSPRICE")) * (IEPSTasa / 100)) 'To calculate iva based on this
                        If LiqTax2 > 0 Then PricePlusLiqTax = Val(DR("INVCSPRICE")) + LiqTax2 '---12/26/13
                        'this line is done here because we will separate cases line from each line. Because this is the 1st line, needs to have everything
                        UnitPrice = Format(CSPrice, "##0.00")
                        Importe = Format(CDbl(UnitPrice) * DR("csshipped"), "##0.00###")
                        TotImp = TotImp + Val(Importe)
                        If Val(CStr(IEPSTasa)) > 0 Then IEPSAmt = Format((Val(CStr(CSPrice)) * DR("csshipped")) * (IEPSTasa / 100), "##0.00###") 'SAT IEPSAmt
                        If LiqTax2 > 0 Then IEPSAmt = Format(DR("csshipped") * LiqTax2, "##0.00") '12/23/13 ------------------------------
                        If (SeparateLiqTaxFlag = "Y" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "Y" And LiqTax2 > 0) Then '---12/26/13
                            IEPSAmt_Prt = IEPSAmt 'Ieps printing
                            TotIEPS_Prt = CStr(Val(TotIEPS_Prt) + Val(IEPSAmt))
                        End If
                        If (SeparateLiqTaxFlag = "N" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "N" And LiqTax2 > 0) Then 'Price should add IEPS on inv '
                            If Val(CStr(IEPSTasa)) > 0 Then UnitPrice = Format(DR("INVCSPRICE") + (DR("INVCSPRICE") * (IEPSTasa / 100)), "##0.00###")
                            If LiqTax2 > 0 Then UnitPrice = Format(DR("INVCSPRICE") + LiqTax2, "##0.00###") '12/23/2013 ------
                            Importe = Format(CDbl(UnitPrice) * DR("csshipped"), "##0.00###") '(calculation * qty does not match JETS)
                            CsImporte = CDbl(Importe) 'ONLY USED WHEN CS & EA FOR LIQUOR NO DESGLOSE CALCULATION... ON EACH
                            IEPSAmt_Prt = vbNullString
                        End If
                        TotImp_Prt = TotImp_Prt + Val(Importe)
                        If Val(CStr(IvaTasaDtl)) > 0 Then
                            IvaAmt = CStr(PricePlusLiqTax * DR("csshipped") * (IVATasa / 100))
                            '052624 JBS An: this format will just return "0" instead of roundup. 
                            'IvaAmt = Format(IvaAmt, "0.00") '01/27/23 ADDED FORMAT, IF TOO MANY DECIMALS IT WON'T PROCESS....
                            IvaAmt = RoundUpToDecimals(IvaAmt, 2)
                            TotalTax_New = TotalTax_New + CDbl(IvaAmt) '----04/12/19 sir 1794, when ea and cs present one cent difference and it doesn't generate cfdi
                        End If
                        CsImpIvaIeps = (Val(CStr(CSPrice)) * DR("csshipped")) + Val(IvaAmt) + Val(IEPSAmt)
                        '' CsImpIvaIeps = (Val(UnitPrice) * DR("CSSHIPPED")) + Val(IvaAmt)
                        Units = "CA"
                        UOMCatalog = "XBX" '06/21/18 JMX NOW WANTS TO CHANGE UNIT OF MEASURE TO BE 'XBX' IF CASES, OR 'H87' IF EACHES.
                        'DTL = DTL & "NUMLIN  " & d & vbNewLine '-------------01/05/18 commented because I am placing at end
                        DTL = DTL & "CANTID  " & DR("csshipped") & vbNewLine & "CANTID_CA  " & DR("csshipped") & vbNewLine & "DESCRI  " & ItemDesc.ToString.Trim & vbNewLine
                        DTL = DTL & "CANPAQ  " & DR("csshipped") & vbNewLine
                        DTL = DTL & "CANEMP  " & DR("csshipped") & vbNewLine & "UNIDAD  " & Units & vbNewLine & "CVESKU  " & DR("itemcode") & vbNewLine
                        DTL = DTL & "ESTILV     " & DR("itemcode") & vbNewLine '10/31/17 per Master EDI to be able to print item code on pdf.
                        DTL = DTL & "CVEPRODSERV     " & ItemCatalog & vbNewLine '10/31/17 New SAT item code
                        DTL = DTL & "CVEUNIDAD     " & UOMCatalog & vbNewLine '11/15/17 New for 3.3 version
                        DTL = DTL & "CODUPC  " & UPC & vbNewLine & "PIEPEM  " & DR("eapercs") & vbNewLine & "PIEPEM2 " & DR("eapercs") & vbNewLine & "CODDUN  " & vbNewLine
                        '*****************OCTOBER 2021*****************************
                        DTLCP = DTLCP & "COM_CPT_INIMER " & vbNewLine & vbNewLine 'Inicio de mercancia
                        DTLCP = DTLCP & "   COM_CPT_MER_BIENTRA " & ItemCatalog & vbNewLine 'BienesTransp (clave de producto)
                        DTLCP = DTLCP & "   COM_CPT_MER_DESCRI " & ItemDesc & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_MER_CANTID 1" & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_MER_CVUNID  " & UOMCatalog & vbNewLine '
                        '052424 JBS An: fixed double round up issue,  
                        'DtlWeight = CDbl(Format(RoundUpToDecimals(DR("csshipped") * DR("grosskg"), 2), "###0.00"))
                        DtlWeight = RoundUpToDecimals(DR("csshipped") * DR("grosskg"), 2)
                        If DtlWeight = 0 Then DtlWeight = 0.01
                        TotWeight = TotWeight + DtlWeight
                        DTLCP = DTLCP & "   COM_CPT_MER_PKG " & DtlWeight & vbNewLine & vbNewLine 'itemweight  PesoEnKg
                        If LiqPresentDtl Then
                            DTLCP = DTLCP & "      COM_CPT_MER_MATPEL Sí  " & vbNewLine 'Material peligroso
                            DTLCP = DTLCP & "      COM_CPT_MER_CVEMATPEL   3065  " & vbNewLine ' BEBIDAS ALCOHOLICAS, 24% pero no más de 70% de alcohol en volumen
                            DTLCP = DTLCP & "      COM_CPT_MER_EMB 4C1" & vbNewLine 'NEED TO KNOW THE EMBALAJE CORRECTO !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                        End If
                        If Peli = "No" Then
                            DTLCP = DTLCP & "      COM_CPT_MER_MATPEL No " & vbNewLine '01/23/24
                        End If
                        DTLCP = DTLCP & "   COM_CPT_INICANTRAN " & vbNewLine 'Inicio canidad trasladada
                        DTLCP = DTLCP & "   COM_CPT_CMER_CANTID " & DR("csshipped") & vbNewLine
                        DTLCP = DTLCP & "      COM_CPT_CMER_IDORI OR000001" & vbNewLine
                        DTLCP = DTLCP & "     COM_CPT_CMER_IDDES DE000001" & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_FINCANTRAN" & vbNewLine
                        DTLCP = DTLCP & "COM_CPT_FINMER" & vbNewLine & vbNewLine

                        '----------------------CHANGES FOR FREE PRODUCT-----------------02/09/18
                        If CSPrice = 0 Then 'FREE CASES FIRST (THIS IS WHEN CS AND EA PRESENT)  !!!!!!!!!!!!!!!!!!
                            CSPrice = 0.01
                            DiscExists = True
                            DiscPer = 100
                            'VERSION 4.0 MODIFICATION COMMENT THESE TWO LINES
                            'TotAmtNotTaxable = Val(TotAmtNotTaxable) + (0.01 * DR("csshipped"))
                            'TotAmtBeforeTaxes = TotAmtBeforeTaxes + (0.01 * DR("csshipped"))
                            'END VERSION 4 MOD
                            DiscAmt = 0.01 * DR("csshipped")
                            IEPSTasa = 0 'In case free item has ieps do not report tasa because it won't have any ieps amount
                            LiqTax2 = 0 '01/25/24  In case free item, do not report any tasa, no ieps amt should be reported
                            TotDiscAmt = TotDiscAmt + DiscAmt
                        End If
                        If RFC <> "PHI830429MG6" Then '-----------------Palacio de Hierro doesn't want 0 when iva is 0 '04/20/18
                            If Trim(IvaAmt) = vbNullString Then IvaAmt = "0" : If Trim(IEPSAmt) = vbNullString Then IEPSAmt = "0" '04/20/18
                        End If '04/20/18
                        '--------------------------------------------------------------------------------END OF 02/19/18
                        '---VERSION 4.0 IF ZERO DO NOT USE THESE---
                        '----------JUST CHECKING TAKING OUT IF ----!!!! CHECK LATER!!!
                        'If Val(IvaTasaDtl) > 0 Then   '01/25/24 ---- CALPICO ISSUE NOT SURE IT DOESN'T COME OUT IF NO TASIPE MONIPE
                        DTL = DTL & "TASIPE  " & IvaTasaDtl & vbNewLine
                        DTL = DTL & "MONIPE  " & IvaAmt & vbNewLine 'IvaAmt is used here instead of lineiva because calculation
                        'End If
                        '--------------------VERSION 4.0 - OBJECTO DE IMPUESTO ------------
                        If Val(CStr(IEPSTasa)) > 0 Or LiqTax2 > 0 Then
                            If SeparateLiqTaxFlag = "N" Then
                                'DTL = DTL & "OBJIMP 03" & vbNewLine
                                DTL = DTL & "OBJIMP 02" & vbNewLine 'TEMP CHANGED EVERYTHING TO 02
                            Else
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            End If
                        Else
                            If Val(CStr(IvaTasaDtl)) > 0 Then
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            Else
                                If DiscExists Then '05/27/22
                                    DTL = DTL & "OBJIMP 01" & vbNewLine '05/27/22 WHEN FREE ITEMS NO IVA SHOULD BE REPORTED
                                    DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine
                                Else
                                    DTL = DTL & "OBJIMP 02" & vbNewLine 'PER MA. ELENA  ALL ITEMS SHOULD BE 02, JFC DOES NOT HAVE 01 TYPE-ASK IF OBJIMP SHOULD EXIST FOR 8% WICH IS IEPS BUT NOT IVA--------- VERSION 4.???
                                    DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine & "FCTTASIPE   0.000000" & vbNewLine
                                    DTL = DTL & "IMPORTIPE " & Format(CSPrice * DR("csshipped"), "##0.00") & vbNewLine
                                    DTL = DTL & "TIPIPETR Tasa" & vbNewLine & "TASIEP 0" & vbNewLine
                                End If
                            End If
                        End If
                        If IEPSTasa = 8 Then '8% ieps but no iva, detail should show 0% IVA
                            DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine & "FCTTASIPE   0.000000" & vbNewLine
                            DTL = DTL & "IMPORTIPE " & Format((CSPrice * DR("csshipped")) + CDbl(IEPSAmt), "##0.00") & vbNewLine
                            DTL = DTL & "TIPIPETR Tasa" & vbNewLine '----------------------------- & "TASIEP 0" & vbNewLine
                            TotAmtNotTaxable = TotAmtNotTaxable + CDbl(IEPSAmt)
                        End If
                        '------------------------------
                        DTL = DTL & "MONIEP_IEPS  " & IEPSAmt_Prt & vbNewLine 'added iepsamt_prnt
                        DTL = DTL & "IMPIVAIEPS  " & CsImpIvaIeps & vbNewLine
                        DTL = DTL & "PBRUDE_IEPS  " & UnitPrice & vbNewLine '01/02/18
                        DTL = DTL & "IMPBRU_PRT  " & CDbl(Importe) & vbNewLine '01/02/18
                        '********************************************SIR 1794********************************************************
                        If ((SeparateLiqTaxFlag = "N" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "N" And LiqTax2 > 0)) And IEPS_NO_SeparateXml Then '-->>>04/09/19 ****SIR # 1794*** ------------->>>>>>>>>>>
                            ' DTL = DTL & "TASIEP  " & "0" & vbNewLine & "MONIEP  " & "0" & vbNewLine  'VERSION 4 DO NOT USE IF ZERO
                            DTL = DTL & "PBRUDE  " & Format(CDbl(UnitPrice), "##0.00###") & vbNewLine '052124 JBS An: added decimal format
                            DTL = DTL & "VALUNI  " & Format(CDbl(UnitPrice), "##0.00###") & vbNewLine '052124 JBS An: added decimal format
                            DTL = DTL & "IMPBRU  " & Format(CDbl(UnitPrice) * DR("csshipped"), "##0.00###") & vbNewLine
                            DTL = DTL & "IMPORT  " & Format(CDbl(UnitPrice) * DR("csshipped"), "##0.00###") & vbNewLine
                        Else
                            If Val(CStr(IEPSTasa)) > 0 Then DTL = DTL & "TASIEP  " & IEPSTasa & vbNewLine & "MONIEP  " & CDbl(IEPSAmt) & vbNewLine 'VERSION 4. ADDED IF ONLY USE IF > 0
                            DTL = DTL & "PBRUDE  " & Format(CSPrice, "##0.00") & vbNewLine '052124 JBS An: added decimal format
                            DTL = DTL & "IMPBRU  " & Format(CSPrice * DR("csshipped"), "##0.00") & vbNewLine '052124 JBS An: added decimal format
                            DTL = DTL & "VALUNI  " & Format(CSPrice, "##0.00") & vbNewLine '052124 JBS An: added decimal format
                            DTL = DTL & "IMPORT  " & Format(CSPrice * DR("csshipped"), "##0.00") & vbNewLine '052124 JBS An: added decimal format
                        End If

                        If DiscExists Then
                            DTL = DTL & "TDECON    " & DiscPer & vbNewLine
                            DTL = DTL & "MDECON    " & (CSPrice * DR("csshipped")) * (DiscPer / 100) & vbNewLine
                        Else
                            DTL = DTL & "TDECON  " & "0" & vbNewLine & "MDECON  " & "0.00" & vbNewLine 'NO DESCUENTO EN DETALLE HARD CODE 0
                        End If

                        '----------------------------------------
                        'addenda does not get created because when ea and cs present, I did not have the NUMLIN for this 1stline created (CA), only for EA.
                        DTL = DTL & "NUMADU  " & DocID & vbNewLine & "FECADU  " & DocDate & vbNewLine & "ADUANA  " & PortName & vbNewLine & "EANADU  " & vbNewLine
                        '10/20/22 ADDED FOR ADUANA, SAT FORMAT IS TWO ESPACES IN BETWEEN, MASTER EDI WILL SEPARATE IF ADUANA INFO HAS NO SPACE, SO I WILL TRIM
                        DTL = DTL & "NUMPED " & Trim(DocID) & vbNewLine

                        ''''DTL = DTL & "NUMLIN  " & d & vbNewLine  '01/05/18  -- NumLin was being used wrongly ifused on top. Now I am placing at the end. 01/17/23 placed NUMLIN AFTER SUGAR TAX
                        If LiqTax2 > 0 Then '01/23/18---------SUGARY DRINKS!!!!!!!!!!------------------
                            'If DiscAmt = 0 And SeparateLiqTaxFlag = "Y" Then  ' don't do this bellow for free items COMMENT OUT 07/14/23 REPLACED BY BELLOW
                            If DiscAmt = 0 And SeparateLiqTaxFlag = "Y" Then ' don't do this bellow for free items --07/14/23  ADDED SEPARATELIQTAXFLAG
                                DTL = DTL & "FCTTASIEP   " & IEPSperLiter & vbNewLine '05/09/18  now is 1.17   'DTL = DTL & "FCTTASIEP   1.000000" & vbNewLine '
                                DTL = DTL & "TIPIEPTR    Cuota" & vbNewLine
                                'DTL = DTL & "IMPORTIEP    " & Format(Val(IEPSAmt) / IEPSperLiter, "0.######") & vbNewLine    '05/09/18 added /IEPSperLiter
                                DTL = DTL & "IMPORTIEP    " & Format(Val(IEPSAmt) / IEPSperLiter, "0.##") & vbNewLine '05/31/22 VERSION 4 USE ONLY TWO DIGITS, OTHERWISE ERRORS OUT
                                'VERSION 4 ADD THE FOLLOWING..... WE NEED TO SPECIFY THE IVA EVEN IF IT IS 0% like in the sugar products which have IEPS but 0% IVA
                                DTL = DTL & "MONIEP " & CDbl(Val(IEPSAmt)) & vbNewLine
                                'If LineIVA = 0 Then '05/22/22 some 1.17 cuota also have iva, so do not report 0% if iva exists. --01/11/23
                                ''''''''If Val(LineIVA) = 0 Then '05/22/22 some 1.17 cuota also have iva, so do not report 0% if iva exists.  --01/11/23 added val bug detected 0981304
                                If Val(CStr(IvaTasaDtl)) = 0 Then '01/26/23  the modificatin above, should actually be this one  ivatasadtl not lineiva.....
                                    DTL = DTL & "TASIPE 0 " & vbNewLine & "MONIPE 0" & vbNewLine & "FCTTASIPE 0.000000" & vbNewLine
                                    DTL = DTL & "TIPIPETR Tasa " & vbNewLine & "IMPORTIPE " & ((CSPrice * DR("csshipped")) + CDbl(IEPSAmt)) & vbNewLine
                                    IEPSnoIVA = IEPSnoIVA + CDbl(IEPSAmt)
                                    'TotAmtNotTaxable = TotAmtNotTaxable + IEPSAmt 'I think I don'tneed this, ck..........01/26/23    comment out.... this is cases and ea present.....
                                Else

                                End If
                                '--------------END OF ADDITON FOR VERSION 4--------------------------------
                            End If
                        End If
                        DTL = DTL & "NUMLIN  " & d & vbNewLine '01/27/23 MOVED NUMLIN AFER SUGAR TAX HERE !!
                        DTL = DTL & "D " & vbNewLine '01/27/23 WAS MISSING THE NEW DETAIL LINE...
                        TotIEPSAmt = TotIEPSAmt + Val(IEPSAmt)

                        '************next line,EACH info *********************************************
                        d = d + 1
                        'PricePlusLiqTax = Val(DR("inveaprice")) + (Val(DR("inveaprice")) * (IEPStasa / 100)) '--12/26/13
                        PricePlusLiqTax = Val(DR("INVEAPRICE")) '--12/26/13 'Before taxes, if any tax, add bellow
                        If IEPSTasa > 0 Then PricePlusLiqTax = Val(DR("INVEAPRICE")) + (Val(DR("INVEAPRICE")) * (IEPSTasa / 100)) '12/26/13
                        If LiqTax2 > 0 Then PricePlusLiqTax = Val(DR("INVEAPRICE")) + (LiqTax2 / DR("eapercs")) '--12/26/13
                        '--2012 Oct 'If SlsTaxFlag Then TotAmtToBeTaxed = Val(TotAmtToBeTaxed) + Val(DR("lineprice")) + Val(DR("lineliqtax")) 'Amt for ea and cs combined we do it once

                        UnitPrice = Format(Eaprice, "##0.00")
                        Importe = Format(CDbl(UnitPrice) * DR("eashipped"), "##0.00###")
                        TotImp = TotImp + Val(Importe)
                        IEPSAmt = vbNullString
                        If Val(CStr(IEPSTasa)) > 0 Then IEPSAmt = Format((Val(CStr(Eaprice)) * DR("eashipped")) * (IEPSTasa / 100), "##0.00###") 'SAT IEPSAmt
                        If LiqTax2 > 0 Then IEPSAmt = Format(DR("eashipped") * (LiqTax2 / DR("eapercs")), "##0.00###") '12/23/13 ------------------------------
                        If SeparateLiqTaxFlag = "N" Then 'Price should add IEPS on inv '12/01/10 modify Ea and Cs prices on the file requested by WFactura
                            If IEPSTasa > 0 Then UnitPrice = Format(DR("INVEAPRICE") + (DR("INVEAPRICE") * (IEPSTasa / 100)), "##0.00###")
                            If LiqTax2 > 0 Then UnitPrice = Format(Eaprice + (LiqTax2 / DR("eapercs")), "##0.00###")
                            'If IEPStasa <= 0 Then 'CHANGED ONLY TO MATCH CALCULATION ON JETS IE 0243455 WHERE 34.16 * 12 = 409.86 INSTEAD OF 409.92
                            If IEPSTasa = 0 And LiqTax2 = 0 Then 'CHANGED ONLY TO MATCH CALCULATION ON JETS IE 0243455 WHERE 34.16 * 12 = 409.86 INSTEAD OF 409.92
                                Importe = Format(CDbl(UnitPrice) * DR("eashipped"), "##0.00###") '01/26/11 (calculation * qty does not match JETS) Put it back on 2/17/11
                            Else
                                Importe = Format(Val(DR("lineprice")) + Val(DR("lineliqtax")) - CsImporte, "##0.00###")
                            End If
                        End If
                        If (SeparateLiqTaxFlag = "Y" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "Y" And LiqTax2 > 0) Then
                            IEPSAmt_Prt = IEPSAmt 'Ieps printing
                            TotIEPS_Prt = CStr(Val(TotIEPS_Prt) + Val(IEPSAmt))
                        End If
                        TotImp_Prt = TotImp_Prt + Val(Importe)
                        'ACCUMULATE IEPSAMT BECAUSE PRINTED AFER EACH, SHOULD HAVE CS AMT INCLUDED! WAS PRINTING SEPARATELY
                        If SeparateLiqTaxFlag = "Y" And LiqTax2 > 0 Then IEPSAmt = Format((LiqTax2 / DR("eapercs")) * (DR("eashipped")), "##0.00")
                        If Val(CStr(IvaTasaDtl)) > 0 Then
                            IvaAmt = CStr(PricePlusLiqTax * DR("eashipped") * (IVATasa / 100))

                            '052624 JBS An: this format will just return "0" instead of roundup. 
                            'IvaAmt = Format(IvaAmt, "0.00") '01/27/23 ADDED FORMAT, IF TOO MANY DECIMALS IT WON'T PROCESS
                            IvaAmt = RoundUpToDecimals(IvaAmt, 2)
                            TotalTax_New = TotalTax_New + CDbl(IvaAmt) '----04/12/19 sir 1794, when ea and cs present one cent difference and it doesn't generate cfdi
                        End If
                        EaImpIvaIeps = (Val(CStr(Eaprice)) * DR("eashipped")) + Val(IvaAmt) + Val(IEPSAmt)
                        Units = "EA"
                        UOMCatalog = "H87" '06/21/18 JMX NOW WANTS TO CHANGE UNIT OF MEASURE TO BE 'XBX' IF CASES, OR 'H87' IF EACHES.

                        '--------------------VERSION 4.0 - OBJECTO DE IMPUESTO WHEN EA AND CS PRESENT CS SHOULD HAVE THEIR OBJIMP ALREADY------------
                        If Val(CStr(IEPSTasa)) > 0 Or LiqTax2 > 0 Then
                            If SeparateLiqTaxFlag = "N" Then
                                'DTL = DTL & "OBJIMP 03" & vbNewLine
                                DTL = DTL & "OBJIMP 02" & vbNewLine 'TEMP CHANGED EVERYTHING TO 02
                            Else
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            End If
                        Else
                            If Val(CStr(IvaTasaDtl)) > 0 Then
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            Else
                                If DiscExists Then '05/27/22
                                    DTL = DTL & "OBJIMP 01" & vbNewLine '05/27/22 WHEN FREE ITEMS NO IVA SHOULD BE REPORTED
                                    DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine
                                Else
                                    DTL = DTL & "OBJIMP 02" & vbNewLine ' PER MA. ELENA  ALL ITEMS SHOULD BE 02, JFC DOES NOT HAVE 01 TYPE-CHECK IF 8% SHOULD BE 02 BECAUSE IT IS IESPS BUT NO  IVA OR IEPS--------- VERSION 4.???
                                    DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine & "FCTTASIPE   0.000000" & vbNewLine
                                    DTL = DTL & "IMPORTIPE " & Format(CDbl(UnitPrice) * DR("eashipped"), "##0.00") & vbNewLine
                                    DTL = DTL & "TIPIPETR Tasa" & vbNewLine & "TASIEP 0" & vbNewLine
                                End If
                            End If
                        End If
                        If IEPSTasa = 8 Then '8% ieps but no iva, detail should show 0% IVA
                            DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine & "FCTTASIPE   0.000000" & vbNewLine
                            'I think we need here to add the iesamt.............................!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                            DTL = DTL & "IMPORTIPE " & Format(CDbl(UnitPrice) * DR("eashipped"), "##0.00") & vbNewLine
                            DTL = DTL & "TIPIPETR Tasa" & vbNewLine '---------------- & "TASIEP 0" & vbNewLine
                            TotAmtNotTaxable = TotAmtNotTaxable + CDbl(IEPSAmt)
                        End If
                        '--------------------END VERSION 4.0 ADDITION-------------------

                        DTL = DTL & "CANTID  " & DR("eashipped") & vbNewLine & "CANTID_EA  " & DR("eashipped") & vbNewLine & "DESCRI  " & ItemDesc.Trim & vbNewLine
                        DTL = DTL & "CANPAQ  " & DR("eashipped") & vbNewLine
                        DTL = DTL & "CANEMP  " & DR("eashipped") & vbNewLine & "UNIDAD  " & Units & vbNewLine & "CVESKU  " & DR("itemcode") & vbNewLine
                        DTL = DTL & "ESTILV     " & DR("itemcode") & vbNewLine '10/31/17 per Master EDI to be able to print item code on pdf.
                        DTL = DTL & "CVEPRODSERV     " & ItemCatalog & vbNewLine '10/31/17 New SAT item code
                        DTL = DTL & "CVEUNIDAD     " & UOMCatalog & vbNewLine '11/15/17 New for 3.3 version
                        DTL = DTL & "CODUPC  " & UPC & vbNewLine & "PIEPEM  " & "1" & vbNewLine & "PIEPEM2 " & DR("eapercs") & vbNewLine & "CODDUN  " & vbNewLine
                        '*****************OCTOBER 2021*****************************
                        DTLCP = DTLCP & "COM_CPT_INIMER " & vbNewLine & vbNewLine 'Inicio de mercancia
                        DTLCP = DTLCP & "   COM_CPT_MER_BIENTRA " & ItemCatalog & vbNewLine 'BienesTransp (clave de producto)
                        DTLCP = DTLCP & "   COM_CPT_MER_DESCRI " & ItemDesc.Trim & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_MER_CANTID 1" & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_MER_CVUNID  " & UOMCatalog & vbNewLine '
                        '052424 JBS An: fixed double round up issue,
                        'DtlWeight = CDbl(Format(RoundUpToDecimals((DR("eashipped") / DR("eapercs")) * DR("grosskg"), 2), "###0.00"))
                        DtlWeight = RoundUpToDecimals((DR("eashipped") / DR("eapercs")) * DR("grosskg"), 2)
                        If DtlWeight = 0 Then DtlWeight = 0.01
                        TotWeight = TotWeight + DtlWeight
                        DTLCP = DTLCP & "   COM_CPT_MER_PKG " & DtlWeight & vbNewLine & vbNewLine 'itemweight e  PesoEnKg
                        If LiqPresentDtl Then
                            DTLCP = DTLCP & "      COM_CPT_MER_MATPEL Sí  " & vbNewLine 'Material peligroso
                            DTLCP = DTLCP & "      COM_CPT_MER_CVEMATPEL   3065  " & vbNewLine ' BEBIDAS ALCOHOLICAS, 24% pero no más de 70% de alcohol en volumen
                            DTLCP = DTLCP & "      COM_CPT_MER_EMB 4C1" & vbNewLine 'NEED TO KNOW THE EMBALAJE CORRECTO !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                        End If
                        If Peli = "No" Then
                            DTLCP = DTLCP & "      COM_CPT_MER_MATPEL No " & vbNewLine '01/23/24
                        End If
                        DTLCP = DTLCP & "   COM_CPT_INICANTRAN " & vbNewLine 'Inicio canidad trasladada
                        DTLCP = DTLCP & "   COM_CPT_CMER_CANTID " & DR("eashipped") & vbNewLine
                        DTLCP = DTLCP & "      COM_CPT_CMER_IDORI OR000001" & vbNewLine
                        DTLCP = DTLCP & "     COM_CPT_CMER_IDDES DE000001" & vbNewLine
                        DTLCP = DTLCP & "   COM_CPT_FINCANTRAN" & vbNewLine
                        DTLCP = DTLCP & "COM_CPT_FINMER" & vbNewLine & vbNewLine

                        '----------------------CHANGES FOR FREE PRODUCT-----------------02/09/18
                        If Eaprice = 0 Then 'FREE EA (WHEN CS AND EA PRESENT) !!!!!!!!!!
                            Eaprice = 0.01
                            DiscExists = True
                            DiscPer = 100
                            'VERSION 4.0 DO NOT ADD FREE AMOUNT TO TOTAL NOT TAXABLE-- INVOICE DOES NOT COME OUT----COMMENTING THIS...
                            'TotAmtNotTaxable = Val(TotAmtNotTaxable) + (0.01 * DR("eashipped"))
                            'TotAmtBeforeTaxes = TotAmtBeforeTaxes + (0.01 * DR("eashipped"))
                            '---END VERSION 4.0  ---
                            DiscAmt = 0.01 * DR("eashipped")
                            IEPSTasa = 0 'In case free item has ieps do not report tasa because it won't have any ieps amount
                            LiqTax2 = 0 '01/25/24  In case free item, do not report any tasa, no ieps amt should be reported
                            TotDiscAmt = TotDiscAmt + DiscAmt
                        End If
                        '--------------------------------------------------------------------------------END OF 02/19/18
                        If RFC <> "PHI830429MG6" Then '-----------------Palacio de Hierro doesn't want 0 when iva is 0 '04/20/18
                            If Trim(IvaAmt) = vbNullString Then IvaAmt = "0" : If Trim(IEPSAmt) = vbNullString Then IEPSAmt = "0" '04/20/18
                        End If
                        'VERSION 4.0  IF ZERO DO NOT USE THESE VALUES
                        '01/25/24 TEMPORARY   TAKE OUT THE IF --- CALPICO ITEMS ISSUE -----!!!!
                        'If Val(IvaTasaDtl) > 0 Then
                        DTL = DTL & "TASIPE  " & IvaTasaDtl & vbNewLine & "MONIPE  " & IvaAmt & vbNewLine
                        'End If
                        '--------------------VERSION 4.0 - OBJECTO DE IMPUESTO ------------
                        If Val(CStr(IEPSTasa)) > 0 Or LiqTax2 > 0 Then
                            If SeparateLiqTaxFlag = "N" Then
                                'DTL = DTL & "OBJIMP 03" & vbNewLine
                                DTL = DTL & "OBJIMP 02" & vbNewLine 'TEMP CHANGED EVERYTHING TO 02
                            Else
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            End If
                        Else
                            If Val(CStr(IvaTasaDtl)) > 0 Then
                                DTL = DTL & "OBJIMP 02" & vbNewLine
                            Else
                                If DiscExists Then '05/27/22
                                    DTL = DTL & "OBJIMP 01" & vbNewLine '05/27/22 WHEN FREE ITEMS NO IVA SHOULD BE REPORTED
                                    DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine
                                Else
                                    DTL = DTL & "OBJIMP 02" & vbNewLine 'PER MA. ELENA  ALL ITEMS SHOULD BE 02, JFC DOES NOT HAVE 01 TYPE-ASK IF OBJIMP SHOULD EXIST FOR 8% WICH IS IEPS BUT NOT IVA--------- VERSION 4.???
                                    DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine & "FCTTASIPE   0.000000" & vbNewLine
                                    DTL = DTL & "IMPORTIPE " & Format(Eaprice * DR("eashipped"), "##0.00") & vbNewLine
                                    DTL = DTL & "TIPIPETR Tasa" & vbNewLine & "TASIEP 0" & vbNewLine
                                End If
                            End If
                        End If
                        If IEPSTasa = 8 Then '8% ieps but no iva, detail should show 0% IVA
                            DTL = DTL & "TASIPE  0" & vbNewLine & "MONIPE 0 " & vbNewLine & "FCTTASIPE   0.000000" & vbNewLine
                            DTL = DTL & "IMPORTIPE " & Format((Eaprice * DR("eashipped")) + CDbl(IEPSAmt), "##0.00") & vbNewLine
                            DTL = DTL & "TIPIPETR Tasa" & vbNewLine '---------------------------& "TASIEP 0" & vbNewLine
                            TotAmtNotTaxable = TotAmtNotTaxable + CDbl(IEPSAmt)
                        End If
                        '-------END VERSION 4 ADDITION

                        DTL = DTL & "MONIEP_IEPS  " & IEPSAmt_Prt & vbNewLine '12/27/13 added iepsamt_prnt
                        DTL = DTL & "IMPIVAIEPS  " & EaImpIvaIeps & vbNewLine
                        DTL = DTL & "PBRUDE_IEPS  " & UnitPrice & vbNewLine '01/02/18
                        DTL = DTL & "IMPBRU_PRT  " & CDbl(Importe) & vbNewLine '01/02/18

                        '********************************************SIR 1794********************************************************
                        If ((SeparateLiqTaxFlag = "N" And Val(CStr(IEPSTasa)) > 0) Or (SeparateLiqTaxFlag = "N" And LiqTax2 > 0)) And IEPS_NO_SeparateXml Then '-->>>04/09/19 ****SIR # 1794*** ------------->>>>>>>>>>>
                            'DTL = DTL & "TASIEP  " & "0" & vbNewLine & "MONIEP  " & "0" & vbNewLine 'VERSION 4.  ONLY IF > 0, SO I COMMENTED
                            DTL = DTL & "PBRUDE  " & Format(CDbl(UnitPrice), "##0.00###") & vbNewLine '052124 JBS An: Added decimal format. 
                            DTL = DTL & "VALUNI  " & Format(CDbl(UnitPrice), "##0.00###") & vbNewLine '052124 JBS An: Added decimal format. 
                            DTL = DTL & "IMPBRU  " & Format(CDbl(UnitPrice) * DR("eashipped"), "##0.00###") & vbNewLine
                            DTL = DTL & "IMPORT  " & Format(CDbl(UnitPrice) * DR("eashipped"), "##0.00###") & vbNewLine
                        Else
                            If Val(CStr(IEPSTasa)) > 0 Then DTL = DTL & "TASIEP  " & IEPSTasa & vbNewLine & "MONIEP  " & CDbl(IEPSAmt) & vbNewLine 'VERSION 4 ADDED IF ? 0
                            DTL = DTL & "PBRUDE  " & Format(Eaprice, "##0.00") & vbNewLine '052124 JBS An: Added decimal format. 
                            DTL = DTL & "VALUNI  " & Format(Eaprice, "##0.00") & vbNewLine '052124 JBS An: Added decimal format. 
                            DTL = DTL & "IMPBRU  " & Format(Eaprice * DR("eashipped"), "##0.00") & vbNewLine '052124 JBS An: Added decimal format. 
                            DTL = DTL & "IMPORT  " & Format(Eaprice * DR("eashipped"), "##0.00") & vbNewLine '052124 JBS An: Added decimal format. 
                        End If
                        '**********END*******************************SIR 1794 ************************************
                        If DiscExists Then '02/09/18
                            DTL = DTL & "TDECON    " & DiscPer & vbNewLine
                            DTL = DTL & "MDECON    " & (Eaprice * DR("eashipped")) * (DiscPer / 100) & vbNewLine
                        Else
                            DTL = DTL & "TDECON  " & "0" & vbNewLine & "MDECON  " & "0.00" & vbNewLine 'NO DESCUENTO EN DETALLE HARD CODE 0
                        End If
                        If LiqTax2 > 0 Then '01/23/18---------SUGARY DRINKS!!!!!!!!!!------------------
                            'If DiscAmt = 0 Then ' don't do this bellow for free items COMMENT OUT 07/14/23 REPLACED BY BELLOW
                            If DiscAmt = 0 And SeparateLiqTaxFlag = "Y" Then ' don't do this bellow for free items 07/14/23 ADDED SEPARATELIQTAXFLAG
                                DTL = DTL & "FCTTASIEP   " & IEPSperLiter & vbNewLine '05/09/18   now is 1.17 'DTL = DTL & "FCTTASIEP   1.000000" & vbNewLine '
                                DTL = DTL & "TIPIEPTR    Cuota" & vbNewLine
                                'DTL = DTL & "IMPORTIEP    " & Format(Val(IEPSAmt) / IEPSperLiter, "0.######") & vbNewLine
                                DTL = DTL & "IMPORTIEP    " & Format(Val(IEPSAmt) / IEPSperLiter, "0.##") & vbNewLine 'VERSION 4  USE ONLY TWO DECIMALS, OTHERWISE ERRORS OUT
                                DTL = DTL & "MONIEP " & CDbl(Val(IEPSAmt)) & vbNewLine 'ADDED VERSION 4.0  NOT SURE IF IT WAS ON PREVIOS VERSION
                                'VERSION 4 ADD THE FOLLOWING..... WE NEED TO SPECITY THE IVA EVEN IF IT IS 0% like in the sugar products which have IEPS but 0% IVA
                                DTL = DTL & "MONIEP " & CDbl(Val(IEPSAmt)) & vbNewLine
                                'If Val(LineIVA) = 0 Then '05/22/22 some 1.17 cuota also have iva, so do not report 0% if iva exists. '-- 01/11/23 added val -- bug detected inv 0981304 on this section when ea and cs sugar drinks
                                If Val(CStr(IvaTasaDtl)) = 0 Then '01/26/23  the modificatin above, should actually be this one  ivatasadtl not lineiva.....
                                    DTL = DTL & "TASIPE 0 " & vbNewLine & "MONIPE 0" & vbNewLine & "FCTTASIPE 0.000000" & vbNewLine
                                    DTL = DTL & "TIPIPETR Tasa " & vbNewLine & "IMPORTIPE " & ((Eaprice * DR("eashipped")) + CDbl(IEPSAmt)) & vbNewLine
                                    IEPSnoIVA = IEPSnoIVA + CDbl(IEPSAmt)
                                    'TotAmtNotTaxable = TotAmtNotTaxable + IEPSAmt 'I think I don'tneed this, ck.........01/26/23 this for ea and cases.....
                                Else

                                End If

                                '--------------END OF ADDITON FOR VERSION 4--------------------------------
                            End If
                        End If
                        TotIEPSAmt = TotIEPSAmt + Val(IEPSAmt) '2012 Oct
                End Select
                SQL = "select * from custom_doc where itemcode = '" & DR("itemcode") & "'"

                OCM = New OracleCommand(SQL, conn)
                dtc = New DataTable
                dtc.Load(OCM.ExecuteReader)
                'RSc = Db.CreateDynaset(SQL, &H4)
                DocID = vbNullString : DocDate = vbNullString : PortName = vbNullString
                If dtc.Rows.Count > 0 Then
                    DRc = dtc.Rows(0)
                    PortName = vbNullString & DRc("Port")
                    FirstPortionOfName = InStr(PortName, ",")
                    If FirstPortionOfName > 0 Then PortName = Left(PortName, FirstPortionOfName - 1)
                    PortName = Left(PortName, 11) 'restrict port name to a max of 11 per e-mail
                    DocID = StripChars(vbNullString & DRc("doc_id"))
                    If Len(Trim(DocID)) <> 15 Then '12/15/22
                        DocID = vbNullString '12/15/22  to avoid invoice error, make sure docid has 15 characters
                    End If
                    DocDate = Format(DRc("doc_date"), "yyyy-MM-dd")
                    'Do not insert into customs ----- commented on 03/09/22
                    'Put back 07/19/23
                    If Not SorianaErr Then '07/23/23 Soriana some times has errors so we do not want  to insert in custom_doc_invoice table  because many entries could be inserted meantime the error is solved.
                        If Trim(DocID) <> vbNullString Then '07/21/23 if the docid < 15 in lenght, we dont report it because invoice will error out
                            '------07/22/23 --- DO  NOT INSERT INTO CUSTOM_DOC_INVOICE IF ALREADY RECORD WAS ADDED --
                            SQLInsert = "insert into custom_doc_invoice values ('" & InvNum & "','" & DR("itemcode") & "','" & DocID & "','" & Format(DRc("doc_date"), "dd-MMM-yyyy") & "','" & PortName & "')"
                            OCM = New OracleCommand(SQLInsert, conn)
                            OCM.ExecuteNonQuery()
                        End If
                    End If
                End If
                '-------------------------------------
                DTL = DTL & "NUMADU  " & DocID & vbNewLine & "FECADU  " & DocDate & vbNewLine & "ADUANA  " & PortName & vbNewLine & "EANADU  " & vbNewLine
                DTL = DTL & "NUMPED " & Trim(DocID) & vbNewLine '10/20/22 --> create exe on 10/24/22 NUMADU is not being used by MASTEREDI, THEY USE NUMPED
                DTL = DTL & "NUMLIN  " & d & vbNewLine ''-------------01/05/18 placing at end, does not work properly is placed first at when dtl starts
            Next
            DTLCP = DTLCP & "COM_CPT_FINMER" & vbNewLine & "COM_CPT_FINMERS " & vbNewLine
            'Debug.Print DTLCP
            ' ------------TOTAL de peso en lineas de detalle y total de lineas (va antes del lineas de detalle)
            DTLCPT = DTLCPT & "COM_CPT_INIMERS " & vbNewLine 'DELIMITADOR COMIEZO  MERCANCIA
            DTLCPT = DTLCPT & "COM_CPT_MER_PESBRU " & Format(TotWeight, "##0.00") & vbNewLine ''PESO BRUTO HERE!!!!!!!!!!!!!!!!!!
            DTLCPT = DTLCPT & " COM_CPT_MER_PESONET " & Format(TotWeight, "##.00") & vbNewLine ' 'PESO NETO HERE!! using peso bruto as not all items have peso neto.
            'DTLCPT = DTLCPT & "COM_CPT_MER_UNIPES X4G " & vbNewLine 'Ver 3 12/28/23 cajas es incorrecto para el peso
            DTLCPT = DTLCPT & "COM_CPT_MER_UNIPES KGM" & vbNewLine 'Ver 3 12/28/23 kilogramos es lo correcto para el peso
            DTLCPT = DTLCPT & "COM_CPT_MER_NUMTOT  " & d & vbNewLine & vbNewLine '-MOD THIS TO REAL VALUE...  'NumTotalMercancias # total  mercancías q se trasladan Este # debe ser = al # d secciones Mercancia q se registren
            'DTLCPT = DTLCPT & "" & vbNewLine & vbNewLine
        End If
        'NO ITEM CODE SO JFCITEM AND BRANCHITEM PRODUCE NOTHING, NO CS, no EA.. some times we get only remark code no item....  ie:  39-0128122,39-0210747, specially notas de cargo
        'OR IT MIGHT BE ITEMCODE PRESENT BUT NO EA NOR CS SHIPPED, MIGHT BE ADDITIONAL CHARGE (ERROR ON PRICE OR SOMETHING ELSE) ie: 39-0356084
        NoItemCdRoutine(InvNum, IEPSTasa, IEPSAmt)
        Exit Sub
ErrHndlr:
        ErrMsgLog = New String("*", 50) & vbNewLine & Format(Now, "MM/dd/yy HH:mm") & " -> ERROR:  " & Err.Number & "-->" & Err.Description & "." & vbNewLine & "Error on GetDtl Routine.  Prog:  DigInv3." & vbNewLine & "Total Invoices:  " & TotInvs & vbNewLine & vbNewLine & "Last SQL ran : " & vbNewLine & SQL & vbNewLine & "Item being processed for line: " & d & " " & ItemCode & vbNewLine & "Order being processed: " & InvNum & vbNewLine
        CkFileExists(DirToOutputError, ErrMsgLog, "ERR-Dig-Inv3_3.log")
        SendEmail(ErrMsgLog)
        End 'for error, just end and figure out what happened!!! we are using begintrans and commit, so reprocess.
    End Sub
    '========================================================
    Private Function GetRelUuid(ByRef RefInvNo As String) As String
        '========================================================
        '02/21/18 Function added for SIR # 1702
        Dim SQL As String
        Dim OCM As OracleCommand 'add for upgrade 04/26/2024 
        Dim OTR As OracleTransaction
        Dim DT As DataTable
        On Error GoTo ErrHndlr
        SQL = "select r." & Chr(34) & "rep_UUID" & Chr(34) & " UUID " & "from dbo.reporte@jmxinv r " & "Where r." & Chr(34) & "rep_folio" & Chr(34) & " = '" & RefInvNo & "' "
        ''Debug.Print SQL
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
    '===============================================================================
    Private Sub NoItemCdRoutine(ByRef InvNum As String, ByRef IEPSTasa As Double, ByRef IEPSAmt As String)
        '===============================================================================
        'Routine is in place of code commented getdtl routine for dtl with no items 05/21/15
        Dim Rsd2, Rsd3 As Object
        Dim Created1stPart, Created2ndPart As Boolean
        Dim IEPSAmt_Prt, SQL, TmpTotImp As String
        Dim j As Short
        Dim LineImpIvaIeps As Double
        Dim UnitPrice As Double
        Dim FinishedDtl As Boolean 'added 11/02/16 in case we have NC and IVA, need to complete IVA if IVA is last line and no IEPs. Prog waits to read all lines bef printing may have to add ieps to unit price before printing now.

        Dim OCM As OracleCommand 'add for upgrade 05/02/2024 
        Dim ODR As OracleDataReader 'add for upgrade 05/02/2024
        Dim dtd2, dtd3 As DataTable
        Dim DRd3 As DataRow
        Dim Adpt As OracleDataAdapter

        On Error GoTo ErrorHandler
        'SQL = "SELECT * FROM INVDTL D where invhdrnum = " & InvNum & " and itemcode is null ORDER BY LINENUM"
        '05/26/15 in case is AC, it might have item code, but nothing shipped, but has line total (maybe price error, or somehing else)
        'SQL = "SELECT linenum, remarkid, lineprice, linedesc FROM INVDTL D where invhdrnum = '" & InvNum & "' and (itemcode is null or (csshipped = 0 and eashipped = 0 and linetotal > 0) )ORDER BY LINENUM"
        SQL = "SELECT * FROM INVDTL D where invhdrnum = '" & InvNum & "' and (itemcode is null or (csshipped = 0 and eashipped = 0 and linetotal > 0) ) ORDER BY LINENUM"
        ''Debug.Print SQL
        OCM = New OracleCommand(SQL, conn)
        Adpt = New OracleDataAdapter(OCM)
        dtd2 = New DataTable
        'dtd2.Load(OCM.ExecuteReader) '052224 JBS An: different way to read data for better performance
        Adpt.Fill(dtd2)
        If dtd2.Rows.Count > 0 Then
            j = 1
            For Each DR2 As DataRow In dtd2.Rows
                SeparateLiqTaxFlag = IepsSeparate(InvNum, DR2("linenum")) '11/01/16 City Fresko does not need to separate when printing, Comercial Mexica needs separation.
                Select Case (vbNullString & DR2("remarkid"))
                    Case Is = "MC", "AC" 'Nota de Cargo, Aditional charge
                        If j > 1 Then DTL = DTL & "IMPIVAIEPS  " & LineImpIvaIeps & vbNewLine & "NUMLIN  " & TotLines & vbNewLine '--if more than one line of nota de cargo, finish the dtl line, otherwise, it will start wiht 2nd line without finishing 1st.
                        Create1stPartofDtl(DR2)
                        Created1stPart = True : Created2ndPart = True '-- if nota de cargo, and iva, ieps, NC will create the 2nd part of dtl, so we make flag = T.
                        TotImp_Prt = TotImp_Prt + DR2("lineprice") '07/11/14 if line items with lines with no itemcode, then we need to add, no overwrite TotImp_Prt is used in the total line
                        TotAmtToBeTaxed = DR2("lineprice")
                        UnitPrice = DR2("lineprice") '11/02/16  save unit price, this price will be used if no ieps found.
                        LineImpIvaIeps = DR2("lineprice")
                        '11/01//16 - only create the line when liqtax will be separated, otherwise, we will create this line at the end, added if on 11/01/16
                        DTL = DTL & "CODUPC  " & "001234567890" & vbNewLine '07/15/15 per TCM send some value on the UPC for athe addendas.
                        DTL = DTL & "VALUNI  " & DR2("lineprice") & vbNewLine & "IMPORT  " & DR2("lineprice") & vbNewLine
                        DTL = DTL & "PBRUDE  " & DR2("lineprice") & vbNewLine
                        DTL = DTL & "IMPBRU  " & DR2("lineprice") & vbNewLine '02/22/19 bug on nota de cargo was not writting IMPBRU
                        DTL = DTL & "CVEPRODSERV " & "84111506" & vbNewLine '02/26/19 New SAT item code
                        DTL = DTL & "CVEUNIDAD " & "ACT" & vbNewLine '02/26/19 New for 3.3 version
                        If SeparateLiqTaxFlag = "Y" Then '11/01/16 -->> if separate, buld dtl, if not, build dtl later  on...
                            DTL = DTL & "PBRUDE_IEPS  " & DR2("lineprice") & vbNewLine
                            DTL = DTL & "IMPBRU_PRT  " & CDbl(DR2("lineprice")) & vbNewLine
                        End If '11/01/16
                    Case Is = "FC" 'Freight Charge is a complete line by itself. Others might combine up to 3 dtl lines in printing
                        If j > 1 Then DTL = DTL & "IMPIVAIEPS  " & LineImpIvaIeps & vbNewLine & "NUMLIN  " & TotLines & vbNewLine '--if more than one line of nota de cargo, finish the dtl line, otherwise, it will start wiht 2nd line without finishing 1st.
                        Create1stPartofDtl(Rsd2)
                        DTL = DTL & "PBRUDE  " & DR2("lineprice") & vbNewLine & "PBRUDE_IEPS  " & DR2("lineprice") & vbNewLine
                        DTL = DTL & "IMPBRU  " & DR2("lineprice") & vbNewLine & "IMPBRU_PRT  " & CDbl(DR2("lineprice")) & vbNewLine
                        DTL = DTL & "VALUNI  " & DR2("lineprice") & vbNewLine & "IMPORT  " & DR2("lineprice") & vbNewLine
                        If Val(CStr(IEPSTasa > 0)) Then DTL = DTL & "TASIEP  " & IEPSTasa & vbNewLine & "MONIEP  " & CDbl(IEPSAmt) & vbNewLine 'VERSION 4 ADDED IF ONLY IF >0
                        LineImpIvaIeps = DR2("lineprice") '05/22/15
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
                        'TotalTax = Val(DR2("lineprice")) '11/19/15 commented, replaced by bellow stmt
                        TotalTax = CStr(Val(TotalTax) + Val(DR2("lineprice"))) '--11/20/15  added val() for totaltax -- 11/19/15 bug found on printout of invoice:  39-0411482 420.00 printed instead of total 10,125.91
                        TotalTax_New = TotalTax_New + CDbl(TotalTax) '---->>>05/23/2019 SIR 1794, BUG FOUND ON NOTA DE CARGO, NOW WE USE TOTALTAX_NEW INSTEAD OF TOTALTAX
                        'LineImpIvaIeps = LineImpIvaIeps + TotalTax '11/19/15 commented, replaced by bellow stmt
                        LineImpIvaIeps = LineImpIvaIeps + Val(DR2("lineprice")) ' to correct bug found on invoice 39-0411482
                        '-------VERSION 4. ONLY USE THESE VALUES IF > 0
                        If Val(CStr(IVATasa)) > 0 Then
                            DTL = DTL & "MONIPE  " & DR2("lineprice") & vbNewLine 'This line of sls info will be added to dtl on line of NC, or if just iva, the main line was created above.
                            DTL = DTL & "TASIPE  " & IVATasa & vbNewLine '02/25/19
                        End If
                        FinishedDtl = False '11/02/16 -Waiting to read ieps in case we need to add ieps to price for cust with no separate option
                    Case Else ' L1, .... L5 (diff types of ieps) -- prog will work ONLY if just one type of  ieps in in dtl....
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
                                If SeparateLiqTaxFlag = "Y" And Val(CStr(IEPSTasa)) > 0 Then '11/01/16, now we need  ck if cust separates
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
                                If Val(CStr(IEPSTasa)) > 0 Then DTL = DTL & "TASIEP  " & IEPSTasa & vbNewLine & "MONIEP  " & CDbl(IEPSAmt) & vbNewLine 'VERSION 4. ADDED IF
                                DTL = DTL & "MONIEP_IEPS  " & IEPSAmt_Prt & vbNewLine
                                If DR2("lineprice") = 0 Then 'Line price is 0, use the amount in line total
                                    TmpTotImp = DR2("linetotal") 'TmpTotImp is used in dtl for  import because value might be in lineprice or  line total, if i line price, rather read line price
                                Else
                                    TmpTotImp = DR2("lineprice")
                                End If
                                If SeparateLiqTaxFlag = "Y" Then
                                    LineImpIvaIeps = LineImpIvaIeps + CDbl(TmpTotImp) ' DR2("lineprice")
                                    'TotImp_Prt = TotImp_Prt + DR2("lineprice")  '--->>> PENDING CK THIS ONE, I THINK THIS SHOULD BE COMMENTED.!!!!!!  --->>>>
                                Else 'DO NOT SEPARATE, ADD IEPS TO UNIT PRICE.
                                    '--- Now we create the part that was not created because it it only created when we separate ieps.
                                    LineImpIvaIeps = LineImpIvaIeps + CDbl(TmpTotImp) 'DR2("linetotal")  'Add ieps to the import amount to be printed
                                    TotImp_Prt = TotImp_Prt + CDbl(TmpTotImp)
                                    DTL = DTL & "PBRUDE_IEPS  " & TotImp_Prt & vbNewLine '  LineImpIvaIeps & vbNewLine
                                    DTL = DTL & "IMPBRU_PRT  " & TotImp_Prt & vbNewLine ' LineImpIvaIeps & vbNewLine
                                End If
                                FinishedDtl = True
                                '--------------------------
                                TotAmtToBeTaxed = TotAmtToBeTaxed + DR2("lineprice")
                            End If
                        End If
                End Select
                j = j + 1
            Next
            '---------------------------------
            '02/26/19 for nota cargo sir 1702
            '--------------------------------
            If j > 1 Then
                If Val(IEPSAmt) = 0 Then
                    DTL = DTL & "TASIEP " & "0" & vbNewLine
                    DTL = DTL & "MONIEP " & "0" & vbNewLine
                End If
                If CDbl(TotalTax) = 0 Then
                    DTL = DTL & "TASIPE " & "0" & vbNewLine
                    DTL = DTL & "MONIPE " & "0" & vbNewLine
                End If
                '02/26/19 end adding for SIR 1702
                '----------------------
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
    '========================================
    Private Sub Create1stPartofDtl(ByRef DR2 As Object)
        '========================================
        DTL = DTL & "CANTID  " & "1" & vbNewLine & "CANTID_EA  " & "1" & vbNewLine & "DESCRI  " & DR2("linedesc").ToString.Trim & vbNewLine
        DTL = DTL & "CANEMP  " & vbNewLine & "UNIDAD  " & "EA" & vbNewLine
        DTL = DTL & "TDECON  " & "0" & vbNewLine & "MDECON  " & "0.00" & vbNewLine 'NO DESCUENTO EN DETALLE HARD CODE 0
        TotLines = TotLines + 1
    End Sub
    '============================
    Private Sub Create2ndPartOfDtl()
        '============================
        DTL = DTL & "PBRUDE  " & vbNewLine & "PBRUDE_IEPS  " & vbNewLine
        DTL = DTL & "IMPBRU  " & vbNewLine & "IMPBRU_PRT  " & vbNewLine
        DTL = DTL & "VALUNI  " & "0" & vbNewLine & "IMPORT  " & "0" & vbNewLine
    End Sub
    '==========================================================================
    Private Function IepsSeparate(ByRef InvNo As String, ByRef LineNo As String) As String
        '==========================================================================
        '01/10/14 added to get if the ieps for the item is to be separated or not.
        Dim SQL As String
        Dim OCM As OracleCommand 'add for upgrade 05/10/2024 
        Dim DT As DataTable 'add for upgrade 05/10/2024
        Dim DR As DataRow   'add for upgrade 05/10/2024
        On Error GoTo ErrorHandler

        SQL = "select IepsSeparate('" & InvNo & "', '" & LineNo & "') separate from dual"

        OCM = New OracleCommand(SQL, conn)
        DT = New DataTable
        DT.Load(OCM.ExecuteReader)

        If DT.Rows.Count > 0 Then
            DR = DT.Rows(0)
            If Val("" & DR("SEPARATE")) = 1 Then
                IepsSeparate = "Y"
                IEPS_NO_SeparateXml = False '07/10/23--- not sure if this flag should be used at all
            Else
                IepsSeparate = "N"
                '05/31/23 -- SIR #  2536.
                IEPS_NO_SeparateXml = True 'Before we had the option to no separate on pdf, but separate or not on xml, now it should not separate should be on both doc
            End If
        Else
            IepsSeparate = "N"
            '07/10/23 -- SIR #  2536.
            IEPS_NO_SeparateXml = True 'Before we had the option to no separate on pdf, but separate or not on xml, now it should not separate should be on both doc
        End If
        'Debug.Print SQL
        Exit Function 'If no error
ErrorHandler:  'write error file.
        ErrMsgLog = Err.Number & "-->" & Err.Description & "." & vbCrLf & "Error on call stored procedure."
        Call CkFileExists(DirToOutputError, ErrMsgLog, "ERR-DigInv3_3.log")
        SendEmail(ErrMsgLog)
    End Function


    '===========================================================
    Private Function GetJustItemDesc(ByRef ItemCd As String) As String
        '===========================================================
        Dim SQL As String
        Dim OCM As OracleCommand 'add for upgrade 05/10/2024 
        Dim DT As DataTable 'add for upgrade 05/10/2024
        Dim DR As DataRow   'add for upgrade 05/10/2024

        On Error GoTo ErrHndlr
        SQL = "select * from v_jfcitem_package where itemcode = '" & ItemCd & "' "
        OCM = New OracleCommand(SQL, conn)
        DT = New DataTable
        DT.Load(OCM.ExecuteReader)
        If DT.Rows.Count > 0 Then
            DR = DT.Rows(0)
            GetJustItemDesc = vbNullString & DR("PURE_ITEMDESC")
        Else
            GetJustItemDesc = vbNullString
        End If
        Exit Function
ErrHndlr:
        ErrMsgLog = New String("*", 50) & vbNewLine & "Today: " & Format(Now, "MM/dd/yy HH:mm") & "---> Eror # " & Err.Number & "-->" & Err.Description & "." & vbNewLine & "Error on DigInv3 program, GetJustItemDesc Function" & vbNewLine & "SQL: " & SQL & vbNewLine
        ErrMsgLog = ErrMsgLog & "PROGRAM ENDED WITHOUT PROCESSING ANYTHING!!!!" & vbNewLine
        Call CkFileExists(DirToOutputError, ErrMsgLog, "ERR-DigInv3_3.log")
        SendEmail(ErrMsgLog)
        End
    End Function
    '========================================================
    Private Function GetMetodoPago(ByRef Cust As String) As String 'NOT BEING USED ANY MORE.  METODO DE PAGO IS HARD-CODED FOR ALL CUSTOMERS AS PPD
        '========================================================
        Dim SQL, Metodo As String
        Dim OCM As OracleCommand 'add for upgrade 05/10/2024 
        Dim DT As DataTable 'add for upgrade 05/10/2024
        Dim i As Short
        On Error GoTo ErrHndlr
        SQL = "select * from liquorcust where customerid = '" & Cust & "' and liquorcode in ('1','2') order by liquorcode" '-12/17/14 added quotes around 1,2 because was bumping when 'X' in DB
        OCM = New OracleCommand(SQL, conn)
        DT = New DataTable
        DT.Load(OCM.ExecuteReader)
        If DT.Rows.Count > 0 Then
            For Each DR As DataRow In DT.Rows
                If i > 1 Then
                    Metodo = Metodo & ", " & DR("LICENSENUM")
                Else
                    Metodo = vbNullString & DR("LICENSENUM")
                End If
            Next
            GetMetodoPago = Metodo
        Else
            GetMetodoPago = "NO IDENTIFICADO"
        End If
        Exit Function
ErrHndlr:
        ErrMsgLog = New String("*", 50) & vbNewLine & "Today: " & Format(Now, "MM/dd/yy HH:mm") & "---> Eror # " & Err.Number & "-->" & Err.Description & "." & vbNewLine & "Error on DigInv3 program, GetMetodoPago sub" & vbNewLine & "SQL: " & SQL & vbNewLine
        ErrMsgLog = ErrMsgLog & "PROGRAM ENDED WITHOUT PROCESSING ANYTHING!!!!" & vbNewLine
        Call CkFileExists(DirToOutputError, ErrMsgLog, "ERR-Dig-Inv3_3.log")
        SendEmail(ErrMsgLog)
        End
    End Function
    '====================================================
    Private Function StripRFC(ByRef RFC As String) As String
        '====================================================
        Dim TmpStr As String
        TmpStr = RFC
        TmpStr = Replace(TmpStr, "R.F.C.", "")
        TmpStr = Replace(TmpStr, ".", "")
        TmpStr = Replace(TmpStr, " ", "")
        StripRFC = TmpStr
    End Function
    '======================================================
    Function MoneyPhrase(ByVal Amount As Double) As String
        '======================================================
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
    '===========================================================================
    Sub CkFileExists(ByRef StrDir As String, ByRef StrData As String, ByRef StrFile As String)
        '==========================================================================
        Dim filenum As Object
        Dim strFilePath As String
        Dim strFileName As String
        Dim strDirExist As String
        On Error GoTo ErrHndlr
        'Debug.Print StrData
        filenum = FreeFile()
        '18/Jul/2024 JBS Mex
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
        If strFilePath = "" Then
            FileOpen(filenum, strFileName, OpenMode.Output)
            Print(filenum, StrData)
            FileClose(filenum)
        Else
            FileOpen(filenum, strFileName, OpenMode.Append)
            Print(filenum, StrData)
            FileClose(filenum)
        End If
        Exit Sub
ErrHndlr:
        ErrMsgLog = New String("*", 50) & vbNewLine & "Today: " & Format(Now, "MM/dd/yy HH:mm") & "---> Eror # " & Err.Number & "-->" & Err.Description & "." & vbNewLine & "Error on DigInv3 program, CkFileExists sub" & vbNewLine ' & "SQL: " & SQL & vbNewLine
        ErrMsgLog = ErrMsgLog & "PROGRAM ENDED WITHOUT PROCESSING ANYTHING!!!!" & vbNewLine '& Serie & Folio & vbNewLine
        Call CkFileExists(DirToOutputError, ErrMsgLog, "ERR-Dig-Inv3_3.log")
        SendEmail(ErrMsgLog)
    End Sub
    '==================================================================================================
    Private Function GetTermMX(ByRef DR As DataRow) As String
        '===================================================================================================
        If "" & DR("CreditCodeId") = "A" Or "" & "" & DR("CreditCodeId") = "B" Then
            GetTermMX = "DUE: C.O.D"
            TermsDate = Format(DR("invdate"), "yyyy-MM-dd") '09/12/18 Added per Ma Elena's request because there was no date, it was defaulting 06-23-2018, don't know why.
        ElseIf "" & DR("codonliqinv") = "1" And "" & DR("sectioncode") = "5" Then
            GetTermMX = "DUE: C.O.D"
            TermsDate = Format(DR("invdate"), "yyyy-MM-dd") '09/12/18 Added per Ma Elena's request because there was no date, it was defaulting 06-23-2018, don't know why.
        Else
            GetTermMX = Term2MX(DR)
        End If
    End Function
    '================================================
    Private Function Term2MX(ByRef DR As DataRow) As String
        '=================================================
        If DR("DISCPCNT") = 0 Then
            If DR("netdays") = 0 And DR("ardays") = 0 Then
                Term2MX = "DUE:" & Format(DateAdd(DateInterval.Day, 30, CDate(DR("invdate"))), "dd-MM-yy") '04/09/24 Modify for Upgrade in .NET
                TermsDate = Format(DateAdd(DateInterval.Day, 30, CDate(DR("invdate"))), "yyyy-MM-dd") '04/09/24 Modify for Upgrade in .NET
            Else
                If DR("ardays") > DR("netdays") And DR("netdays") <> 0 Then
                    Term2MX = "NET " & DR("netdays") & " DAYS"
                    TermsDate = Format(DateAdd(DateInterval.Day, CInt(DR("netdays")), DR("invdate")), "yyyy-MM-dd") '04/09/24 Modify for Upgrade in .NET
                ElseIf DR("ardays") = 0 And DR("netdays") > 0 Then  'ARDays = 0, NetDays > 0
                    Term2MX = "NET " & DR("netdays") & " DAYS" '04/09/24 Modify for Upgrade in .NET
                    TermsDate = Format(DateAdd(DateInterval.Day, CInt(DR("netdays")), DR("invdate")), "yyyy-MM-dd") '04/09/24 Modify for Upgrade in .NET
                Else
                    If DR("ardays") < 30 Then
                        Term2MX = "DUE:" & Format(DateAdd(DateInterval.Day, CInt(DR("ardays")), DR("invdate")), "dd-MM-yy") '04/09/24 Modify for Upgrade in .NET
                        TermsDate = Format(DateAdd(DateInterval.Day, CInt(DR("ardays")), DR("invdate")), "yyyy-MM-dd") '04/09/24 Modify for Upgrade in .NET

                    Else
                        Term2MX = "DUE:" & Format(DateAdd(DateInterval.Day, 30, CDate(DR("invdate"))), "dd-MM-yy") '04/09/24 Modify for Upgrade in .NET
                        TermsDate = Format(DateAdd(DateInterval.Day, 30, CDate(DR("invdate"))), "yyyy-MM-dd") '04/09/24 Modify for Upgrade in .NET
                    End If
                End If
            End If
        Else
            If DR("netdays") <> 0 Then
                Term2MX = Format(DR("DiscPct"), "#0") & "% " & Format(DR("DISCDAYS"), "#0") & "/NET " & Format(DR("netdays"), "#0") '04/09/24 Modify For Upgrade In .NET
                TermsDate = Format(DateAdd(DateInterval.Day, CInt(DR("netdays")), CDate(DR("invdate"))), "yyyy-MM-dd") '04/09/24 Modify for Upgrade in .NET
            Else
                Term2MX = Format(DR("DiscPct"), "#0") & "% " & Format(DR("DISCDAYS"), "#0") '04/09/24 Modify for Upgrade in .NET
                TermsDate = Format(DateAdd(DateInterval.Day, 30, CDate(DR("invdate"))), "yyyy-MM-dd") '04/09/24 Modify for Upgrade in .NET
            End If
        End If
    End Function
    '============================================================================
    Private Sub UpdateInvHdr(ByRef InvNum As String, ByRef Serie As String, ByRef Folio As String)
        '=============================================================================
        Dim SQL As String
        'Dim RS As Object
        Dim OCM As OracleCommand   '04/29/24 Modify for Upgrade in .NET
        Dim ODR As OracleDataReader '04/29/24 Modify for Upgrade in .NET

        SQL = "select * from invhdr where invhdrnum = '" & InvNum & "'"
        OCM = New OracleCommand(SQL, conn)
        OCM.ExecuteNonQuery()
        On Error GoTo ErrHndlr

        If ODR.HasRows Then
            While ODR.Read() '04/29/24 Modify for Upgrade in .NET
                SQL = "update invhdr set refinvnum = '" & Serie & Folio & "' where invhdrnum = '" & InvNum & "'" '04/29/24 Modify for Upgrade in .NET
                OCM = New OracleCommand(SQL, conn)
                OCM.ExecuteNonQuery()
            End While
        End If
        Exit Sub

        'JBS memo when to commit SQL?
ErrHndlr:
        ErrMsgLog = New String("*", 50) & vbNewLine & "Today: " & Format(Now, "MM/DD/yy HH:mm") & "---> Eror # " & Err.Number & "-->" & Err.Description & "." & vbNewLine & "Error on DigInv3 program, UpdateInvHdr sub" & vbNewLine & "SQL: " & SQL & vbNewLine '04/09/24 Modify for Upgrade in .NET
        ErrMsgLog = ErrMsgLog & "PROGRAM ENDED WITHOUT PROCESSING ANYTHING!!!!" & vbNewLine & Serie & Folio & vbNewLine
        Call CkFileExists(DirToOutputError, ErrMsgLog, "ERR-Dig-Inv3_3.log")
        SendEmail(ErrMsgLog)
        End
    End Sub
    '========================================================================================
    Private Sub SendEmail(ByRef MSG As String)
        '========================================================================================
        Dim Body, EmailAddress, Subject As String
        On Error GoTo ErrHndlr
        Dim ObjMessage As Object

        '04/04/2024 JBS USA change for test start
        'EmailAddress = "ycerda@kmsnet.com"
        'EmailAddress = "hikaru.tanaka@kmsnet.com"
        '18/Jul/2024 JBS Mex
        EmailAddress = "jmx_diginv_error@kmsnet.com"
        '04/04/2024 JBS USA change for test start

        ObjMessage = CreateObject("CDO.Message")
        With ObjMessage
            .From = """JMX-SERVER"" <notice@kmsnet.com>"
            .To = EmailAddress
            .Subject = "ERROR - JMX - DigInv3 - "
            Body = Subject & vbNewLine & "Today: " & Format(Now, "MM/dd/yyyy HH:mm") & vbNewLine & "Check error folder on JMX: " & DirToOutputError & vbNewLine & vbNewLine & MSG
            .textBody = Body
            .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
            .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.kmsnet.com" ' 10/13/21   "smtp.gmail.com"  'Name of Remote SMTP Server
            .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 '10/13/21     '465
            .Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60 'seconds
            .Configuration.Fields.Update()
            '.Send()
        End With
        'UPGRADE_NOTE: Object ObjMessage may not be destroyed until it is garbage collected. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6E35BFF6-CD74-4B09-9689-3E1A43DF8969"'
        ObjMessage = Nothing
        Exit Sub
ErrHndlr:
        ErrMsgLog = "Today: " & Format(Now, "MM/dd/yy HH:mm") & "---> Eror # " & Err.Number & "-->" & Err.Description & "." & vbNewLine & "Error on DigInv2 program, SendEmail sub" & vbNewLine
        Call CkFileExists(DirToOutputError, ErrMsgLog, "ERR-DigInv3_3-SendEmail.txt")
    End Sub
    '====================================================
    Private Function StripChars(ByRef WithChars As String) As String
        '====================================================
        Dim TmpStr As String
        TmpStr = WithChars
        TmpStr = Replace(TmpStr, ",", " ") 'Take out comma
        TmpStr = Replace(TmpStr, ".", "") 'Take out Period
        TmpStr = Replace(TmpStr, "-", "") 'Take out dash from docnum
        TmpStr = Replace(TmpStr, " ", "") 'TAKE OUT SPACES '10/20/22
        StripChars = TmpStr
    End Function

    '052324 JBS An: round decimal 
    Function RoundUpToDecimals(value As Double, digitPosition As Integer) As Double
        Try
            'Dim roundedValue As Double = Math.Round(value, digitPosition, MidpointRounding.ToEven)
            'Return roundedValue

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
            ' If any exception occurs, return the original string
            Return value
        End Try
    End Function

    Function RoundUpWeight(value As Double, digitPosition As Integer) As Double
        Try
            Dim roundedValue As Double = Math.Round(value, digitPosition, MidpointRounding.AwayFromZero)
            Return roundedValue
        Catch ex As Exception
            ' If any exception occurs, return the original string
            Return value
        End Try
    End Function
    '============================================================
    Private Sub CkifInArr(ByRef IepsPer As Double, ByRef IepsNoSeparateXML As Boolean)
        '============================================================
        'SUB NOT USED ANY MORE --- NOW XML AND PDF SHOULD BE SAME IF SEPARATE IEPS, SEPARATE ON BOTH DOC, IF NOT, DO NOT SEPARATE IN EITHER
        Dim FoundIt As Boolean 'When  found turns true so we stop searching.
        Dim i As Short
        'Arr is bidimensional one row, multiple columns.  26.5 T    30  T     8  F  and on...
        'ieps per% (1,1), (1,3) and on..
        'Separate or not will be on (1,2), (1,4) and on...
        If TotPercentages > 0 Then
            For i = 1 To TotPercentages
                If IepsPer = Val(ArrIeps_Sep_Xml(1, (i * 2 - 1))) Then 'To get 1,3,5 and so on...
                    i = TotPercentages 'To stop search
                    FoundIt = True
                End If
            Next i
            If Not FoundIt Then
                TotPercentages = TotPercentages + 1
                ReDim Preserve ArrIeps_Sep_Xml(1, TotPercentages * 2)
                ArrIeps_Sep_Xml(1, TotPercentages * 2 - 1) = IepsPer
                ArrIeps_Sep_Xml(1, TotPercentages * 2) = IepsNoSeparateXML
                HdrNo_SeparateXmlFlag = True
            End If
        Else 'First
            TotPercentages = 1
            ReDim Preserve ArrIeps_Sep_Xml(1, 2)
            ArrIeps_Sep_Xml(1, 1) = IepsPer
            ArrIeps_Sep_Xml(1, 2) = IepsNoSeparateXML
            HdrNo_SeparateXmlFlag = True
        End If
    End Sub
End Module