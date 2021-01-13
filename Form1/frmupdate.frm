VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form frmupdate 
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5700
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin Threed.SSCommand SSCommand1 
      Height          =   660
      Index           =   0
      Left            =   2370
      TabIndex        =   0
      ToolTipText     =   "Blok Data..."
      Top             =   780
      Width           =   660
      _ExtentX        =   1164
      _ExtentY        =   1164
      _Version        =   196610
      Font3D          =   1
      MousePointer    =   16
      ForeColor       =   192
      BackColor       =   -2147483638
      PictureMaskColor=   255
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmupdate.frx":0000
      AutoSize        =   1
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   6
      BevelWidth      =   1
   End
End
Attribute VB_Name = "frmupdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SSCommand1_Click(Index As Integer)
Select Case Index
Case 0
    
   STRSQL = " insert into mandiri.tblnegoptp(custid,promisedate,promisepay,inputdate,type)      select custid,promisedate,promisepay,now(),type from tblreserve where (date_part('month',promisedate) <= date_part('month',now())) and (date_part('year',promisedate) = date_part('year',now())) and stsmove=0 "
   'M_OBJCONN.Execute (strsql)
   
   STRSQL = "update mandiri.tblreserve set stsmove=1 where (date_part('month',promisedate) <= date_part('month',now())) and (date_part('year',promisedate) = date_part('year',now())) and stsmove=0 "
 '  M_OBJCONN.Execute (strsql)
   
   CMDSQL = " insert into mandiri.mgm_hst(custid,agent,hst,datetime,products,tgl,f_cek,statuscall) SELECT custid,agent,'BP-BROKEN PROMISE-Auto-dari account PTP NEW' AS hst, date(now())  as datetime,'Collection' as product , now() tgl,f_cek ,laststatus"
   CMDSQL = CMDSQL + " FROM mandiri.MGM WHERE f_cek like 'PTP-NE' AND  custid  in (select custid from mandiri.vwnegoptplast where hari>7)"
   CMDSQL = CMDSQL + " and custid not in( select custid from mandiri.vwwlunas where custid<>'')"
  ' M_OBJCONN.Execute (cmdsql)
   
    STRSQL = " update mandiri.mgm set tglstatus= now() ,LASTSTATUS='BP-PTP NEW BROKEN PROMISE',KETHSLKERJA='BP-PTP NEW BROKEN PROMISE',F_CEK='BP-',REMARKS = 'BP-BROKEN PROMISE-Auto-dari account PTP New ',RECSTATUS='C',OTO='Y'  where f_cek like 'PTP-NE' AND  custid  in (select custid from mandiri.vwnegoptplast where hari>7) "
    STRSQL = STRSQL + " and custid not in( select custid from mandiri.vwwlunas where custid<>'')"
   ' M_OBJCONN.Execute (strsql)
    WaitSecs (0.2)
    
    
   CMDSQL = " insert into mandiri.mgm_hst(custid,agent,hst,datetime,products,tgl,f_cek,statuscall) SELECT custid,agent,'BP-BROKEN PROMISE-Auto-dari account PTP POP' AS hst, date(now())  as datetime,'Collection' as product , now() tgl,f_cek ,laststatus"
   CMDSQL = CMDSQL + " FROM mandiri.MGM WHERE f_cek like 'PTP-PO' AND  custid  in (select custid from mandiri.vwnegoptplast where hari>7)"
   CMDSQL = CMDSQL + " and custid not in( select custid from vwwlunas where custid<>'')"
   'M_OBJCONN.Execute (cmdsql)
   
    STRSQL = " update mandiri.mgm set tglstatus= now() ,F_CEK='BP-',LASTSTATUS='BP-PTP POP BROKEN PROMISE',KETHSLKERJA='BP-PTP POP BROKEN PROMISE',REMARKS = 'BP-POP BROKEN PROMISE-Auto-dari account POP ',RECSTATUS='C',OTO='Y' where f_cek like 'PTP-PO' AND  custid  in (select custid from mandiri.vwnegoptplast where hari>7) "
    STRSQL = STRSQL + " and custid not in( select custid from mandiri.vwwlunas where custid<>'')"
    M_OBJCONN.Execute (STRSQL)
    WaitSecs (0.2)
    DoEvents
    
    
    CMDSQL = " insert into mandiri.mgm_hst(custid,agent,hst,datetime,products,tgl,f_cek,statuscall) SELECT custid,agent,'POP' AS hst, date(now())  as datetime,'Collection' as product , now() tgl,f_cek ,laststatus"
   CMDSQL = CMDSQL + " FROM mandiri.MGM WHERE f_cek like 'PTP%' "
   CMDSQL = CMDSQL + " and custid in( select custid from mandiri.vwwlunas where custid<>'')"
   'M_OBJCONN.Execute (cmdsql)s
   
    STRSQL = " update mandiri.mgm set tglstatus= now() ,F_CEK='POP',LASTSTATUS='POP',KETHSLKERJA='POP',REMARKS = 'POP',RECSTATUS='C',OTO='Y' where f_cek like 'PTP%' "
    STRSQL = STRSQL + " and custid  in( select custid from mandiri.vwwlunas where custid<>'')"
    M_OBJCONN.Execute (STRSQL)
    WaitSecs (0.2)
    DoEvents
    
  '  strsql = " update mgm set LASTSTATUS=KETHSLKErJA,KETHSLKErJA='POP-PROGRESS OF PAYMENT',F_CEK='POP',rEMArKS = 'POP-PROGRESS OF PAYMENT-Auto',RECSTATUS='C',OTO='Y' where f_cek like 'PTP%' AND custid  in (select custid from vwnegoptplast where hari>7) "
   ' strsql = strsql + "and custid  in( select custid from vwwlunas where custid<>'')"
  '  M_OBJCONN.Execute (strsql)
   ' WaitSecs (0.2)
   ' DoEvents
    
    'strsql = " update mgm set tglstatus= now() ,LASTSTATUS='BP-BROKEN PROMISE',KETHSLKERJA='BP-BROKEN PROMISE',F_CEK='BP-',REMARKS = 'BP-BROKEN PROMISE-Auto',RECSTATUS='C',OTO='Y' where f_cek like 'POP%' AND custid  in (select custid from vwtbllunaspoop where hari>30) "
    'strsql = strsql + " and custid not in( select custid from vwwlunas where custid<>'')"
    'M_OBJCONN.Execute (strsql)
    'DoEvents
    'WaitSecs (0.2)
    MsgBox "Has Been Done"
End Select

End Sub
