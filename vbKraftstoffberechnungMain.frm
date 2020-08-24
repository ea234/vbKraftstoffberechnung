VERSION 5.00
Begin VB.Form m_frmVbKraftstoffberechnung 
   Caption         =   "Test Kraftstoffberechnung"
   ClientHeight    =   11160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18735
   Icon            =   "vbKraftstoffberechnungMain.frx":0000
   LinkTopic       =   "ui_form_name"
   ScaleHeight     =   744
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1249
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox m_txtGewichtAnzahlLiter 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   70
      Text            =   "m_txtGewichtAnzahlLi"
      ToolTipText     =   "Sichere Flugzeit Minuten"
      Top             =   10125
      Width           =   1260
   End
   Begin VB.TextBox m_txtKraftstoffgewichtKg 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5400
      MaxLength       =   20
      TabIndex        =   68
      Text            =   "m_txtKraftstoffgewic"
      ToolTipText     =   "Sichere Flugzeit Minuten"
      Top             =   10125
      Width           =   1260
   End
   Begin VB.TextBox m_txtFaktorKilogrammJeLiter 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   66
      Text            =   "m_txtFaktorKilogramm"
      ToolTipText     =   "Reisegeschw Minute Je Liter"
      Top             =   3375
      Width           =   1260
   End
   Begin VB.CommandButton m_btnStartGeschwindigkeit3 
      Caption         =   "Calc 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5925
      TabIndex        =   62
      Top             =   300
      Width           =   1260
   End
   Begin VB.CommandButton m_btnStartGeschwindigkeit2 
      Caption         =   "Calc 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4500
      TabIndex        =   61
      Top             =   300
      Width           =   1260
   End
   Begin VB.CommandButton m_btnStartGeschwindigkeit1 
      Caption         =   "Calc 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3075
      TabIndex        =   60
      Top             =   300
      Width           =   1260
   End
   Begin VB.TextBox m_tbReisegeschwMinuteJeLiter3 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5925
      MaxLength       =   20
      TabIndex        =   59
      Text            =   "m_tbReisegeschwMinut"
      ToolTipText     =   "Reisegeschw Minute Je Liter"
      Top             =   2805
      Width           =   1260
   End
   Begin VB.TextBox m_tbReisegeschwLiterJeMinute3 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5925
      MaxLength       =   20
      TabIndex        =   58
      Text            =   "m_tbReisegeschwLiter"
      ToolTipText     =   "Reisegeschw Liter Je Minute"
      Top             =   2295
      Width           =   1260
   End
   Begin VB.TextBox m_tbReisegeschwLiter3 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5925
      MaxLength       =   20
      TabIndex        =   57
      Text            =   "m_tbReisegeschwLiter"
      ToolTipText     =   "Reisegeschw Liter"
      Top             =   1785
      Width           =   1260
   End
   Begin VB.TextBox m_tbReisegeschwKmh3 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5925
      MaxLength       =   20
      TabIndex        =   56
      Text            =   "m_tbReisegeschwKmh"
      ToolTipText     =   "Reisegeschw Kmh"
      Top             =   1275
      Width           =   1260
   End
   Begin VB.TextBox m_tbReisegeschwMinuteJeLiter2 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4500
      MaxLength       =   20
      TabIndex        =   55
      Text            =   "m_tbReisegeschwMinut"
      ToolTipText     =   "Reisegeschw Minute Je Liter"
      Top             =   2805
      Width           =   1260
   End
   Begin VB.TextBox m_tbReisegeschwLiterJeMinute2 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4500
      MaxLength       =   20
      TabIndex        =   54
      Text            =   "m_tbReisegeschwLiter"
      ToolTipText     =   "Reisegeschw Liter Je Minute"
      Top             =   2295
      Width           =   1260
   End
   Begin VB.TextBox m_tbReisegeschwLiter2 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4500
      MaxLength       =   20
      TabIndex        =   53
      Text            =   "m_tbReisegeschwLiter"
      ToolTipText     =   "Reisegeschw Liter"
      Top             =   1785
      Width           =   1260
   End
   Begin VB.TextBox m_tbReisegeschwKmh2 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4500
      MaxLength       =   20
      TabIndex        =   52
      Text            =   "m_tbReisegeschwKmh"
      ToolTipText     =   "Reisegeschw Kmh"
      Top             =   1275
      Width           =   1260
   End
   Begin VB.TextBox m_txtStdMinSichereFlugzeit 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4485
      MaxLength       =   20
      TabIndex        =   50
      Text            =   "m_tbSichereFlugzeitM"
      ToolTipText     =   "Sichere Flugzeit Minuten"
      Top             =   9480
      Width           =   1260
   End
   Begin VB.TextBox m_txtStdMinKraftstoffVorrat 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4485
      MaxLength       =   20
      TabIndex        =   49
      Text            =   "m_tbKraftstoffvorrat"
      ToolTipText     =   "Kraftstoffvorrat Minuten"
      Top             =   8460
      Width           =   1260
   End
   Begin VB.TextBox m_txtStdMinKraftstoffExtra 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4485
      MaxLength       =   20
      TabIndex        =   48
      Text            =   "m_tbExtraKraftstoffM"
      ToolTipText     =   "Extra Kraftstoff Minuten"
      Top             =   7950
      Width           =   1260
   End
   Begin VB.TextBox m_txtStdMinReserve 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4485
      MaxLength       =   20
      TabIndex        =   47
      Text            =   "m_tbReserveMinuten"
      ToolTipText     =   "Reserve Minuten"
      Top             =   6930
      Width           =   1260
   End
   Begin VB.TextBox m_txtStdMinAuschweichplatz 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4485
      MaxLength       =   20
      TabIndex        =   46
      Text            =   "m_tbAusweichplatzMin"
      ToolTipText     =   "Ausweichplatz Minuten"
      Top             =   6420
      Width           =   1260
   End
   Begin VB.TextBox m_txtStdMinSteigflug 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4485
      MaxLength       =   20
      TabIndex        =   45
      Text            =   "m_tbSteigflugMinuten"
      ToolTipText     =   "Steigflug Minuten"
      Top             =   5400
      Width           =   1260
   End
   Begin VB.TextBox m_txtStdMinZuschlagAnlassen 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4485
      MaxLength       =   20
      TabIndex        =   44
      Text            =   "m_tbZuschlagAnlassen"
      ToolTipText     =   "Zuschlag Anlassen Minuten"
      Top             =   4890
      Width           =   1260
   End
   Begin VB.TextBox m_txtStdMinReiseflug 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4485
      MaxLength       =   20
      TabIndex        =   43
      Text            =   "m_tbReiseflugMinuten"
      ToolTipText     =   "Reiseflug Minuten"
      Top             =   4380
      Width           =   1260
   End
   Begin VB.TextBox m_txtStdMinAnAbflug 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4485
      MaxLength       =   20
      TabIndex        =   42
      Text            =   "m_tbAnAbflugMinuten"
      ToolTipText     =   "An Abflug Minuten"
      Top             =   5910
      Width           =   1260
   End
   Begin VB.TextBox m_txtStdMinFlugzeitGesamt 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4485
      MaxLength       =   20
      TabIndex        =   41
      Text            =   "m_tbSichereFlugzeitM"
      ToolTipText     =   "Sichere Flugzeit Minuten"
      Top             =   8970
      Width           =   1260
   End
   Begin VB.TextBox m_tbFlugzeitGesamtMinuten 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   39
      Text            =   "m_tbSichereFlugzeitM"
      ToolTipText     =   "Sichere Flugzeit Minuten"
      Top             =   8970
      Width           =   1260
   End
   Begin VB.TextBox m_txtLog 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8865
      Left            =   7245
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   38
      Text            =   "vbKraftstoffberechnungMain.frx":030A
      Top             =   180
      Width           =   5670
   End
   Begin VB.TextBox m_tbAnAbflugLiter 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5865
      MaxLength       =   20
      TabIndex        =   37
      Text            =   "m_tbAnAbflugLiter"
      ToolTipText     =   "An Abflug Liter"
      Top             =   5910
      Width           =   1260
   End
   Begin VB.TextBox m_tbAnAbflugMinuten 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   36
      Text            =   "m_tbAnAbflugMinuten"
      ToolTipText     =   "An Abflug Minuten"
      Top             =   5910
      Width           =   1260
   End
   Begin VB.TextBox m_tbReisegeschwKmh1 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   1
      Text            =   "m_tbReisegeschwKmh"
      ToolTipText     =   "Reisegeschw Kmh"
      Top             =   1275
      Width           =   1260
   End
   Begin VB.TextBox m_tbReisegeschwLiter1 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   3
      Text            =   "m_tbReisegeschwLiter"
      ToolTipText     =   "Reisegeschw Liter"
      Top             =   1785
      Width           =   1260
   End
   Begin VB.TextBox m_tbReisegeschwLiterJeMinute1 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   5
      Text            =   "m_tbReisegeschwLiter"
      ToolTipText     =   "Reisegeschw Liter Je Minute"
      Top             =   2295
      Width           =   1260
   End
   Begin VB.TextBox m_tbReisegeschwMinuteJeLiter1 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   7
      Text            =   "m_tbReisegeschwMinut"
      ToolTipText     =   "Reisegeschw Minute Je Liter"
      Top             =   2805
      Width           =   1260
   End
   Begin VB.TextBox m_tbReiseflugMinuten 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   9
      Text            =   "m_tbReiseflugMinuten"
      ToolTipText     =   "Reiseflug Minuten"
      Top             =   4380
      Width           =   1260
   End
   Begin VB.TextBox m_tbReiseflugLiter 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5865
      MaxLength       =   20
      TabIndex        =   10
      Text            =   "m_tbReiseflugLiter"
      ToolTipText     =   "Reiseflug Liter"
      Top             =   4380
      Width           =   1260
   End
   Begin VB.TextBox m_tbZuschlagAnlassenMinuten 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   12
      Text            =   "m_tbZuschlagAnlassen"
      ToolTipText     =   "Zuschlag Anlassen Minuten"
      Top             =   4890
      Width           =   1260
   End
   Begin VB.TextBox m_tbZuschlagAnlassenLiter 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5865
      MaxLength       =   20
      TabIndex        =   13
      Text            =   "m_tbZuschlagAnlassen"
      ToolTipText     =   "Zuschlag Anlassen Liter"
      Top             =   4890
      Width           =   1260
   End
   Begin VB.TextBox m_tbSteigflugMinuten 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   15
      Text            =   "m_tbSteigflugMinuten"
      ToolTipText     =   "Steigflug Minuten"
      Top             =   5400
      Width           =   1260
   End
   Begin VB.TextBox m_tbSteigflugLiter 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5865
      MaxLength       =   20
      TabIndex        =   16
      Text            =   "m_tbSteigflugLiter"
      ToolTipText     =   "Steigflug Liter"
      Top             =   5400
      Width           =   1260
   End
   Begin VB.TextBox m_tbAusweichplatzMinuten 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   18
      Text            =   "m_tbAusweichplatzMin"
      ToolTipText     =   "Ausweichplatz Minuten"
      Top             =   6420
      Width           =   1260
   End
   Begin VB.TextBox m_tbAusweichplatzLiter 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5865
      MaxLength       =   20
      TabIndex        =   19
      Text            =   "m_tbAusweichplatzLit"
      ToolTipText     =   "Ausweichplatz Liter"
      Top             =   6420
      Width           =   1260
   End
   Begin VB.TextBox m_tbReserveMinuten 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   21
      Text            =   "m_tbReserveMinuten"
      ToolTipText     =   "Reserve Minuten"
      Top             =   6930
      Width           =   1260
   End
   Begin VB.TextBox m_tbReserveLiter 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5865
      MaxLength       =   20
      TabIndex        =   22
      Text            =   "m_tbReserveLiter"
      ToolTipText     =   "Reserve Liter"
      Top             =   6930
      Width           =   1260
   End
   Begin VB.TextBox m_tbMindestKraftstoffbedarfLiter 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5865
      MaxLength       =   20
      TabIndex        =   24
      Text            =   "m_tbMindestKraftstof"
      ToolTipText     =   "Mindest Kraftstoffbedarf Liter"
      Top             =   7440
      Width           =   1260
   End
   Begin VB.TextBox m_tbExtraKraftstoffMinuten 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   26
      Text            =   "m_tbExtraKraftstoffM"
      ToolTipText     =   "Extra Kraftstoff Minuten"
      Top             =   7950
      Width           =   1260
   End
   Begin VB.TextBox m_tbExtraKraftstoffLiter 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5865
      MaxLength       =   20
      TabIndex        =   27
      Text            =   "m_tbExtraKraftstoffL"
      ToolTipText     =   "Extra Kraftstoff Liter"
      Top             =   7950
      Width           =   1260
   End
   Begin VB.TextBox m_tbKraftstoffvorratMinuten 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   29
      Text            =   "m_tbKraftstoffvorrat"
      ToolTipText     =   "Kraftstoffvorrat Minuten"
      Top             =   8460
      Width           =   1260
   End
   Begin VB.TextBox m_tbKraftstoffvorratLiter 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5865
      MaxLength       =   20
      TabIndex        =   30
      Text            =   "m_tbKraftstoffvorrat"
      ToolTipText     =   "Kraftstoffvorrat Liter"
      Top             =   8460
      Width           =   1260
   End
   Begin VB.TextBox m_tbSichereFlugzeitMinuten 
      Alignment       =   1  'Rechts
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3075
      MaxLength       =   20
      TabIndex        =   32
      Text            =   "m_tbSichereFlugzeitM"
      ToolTipText     =   "Sichere Flugzeit Minuten"
      Top             =   9480
      Width           =   1260
   End
   Begin VB.Label m_lblGewichtKG 
      Caption         =   "KG"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6750
      TabIndex        =   72
      Top             =   10125
      Width           =   765
   End
   Begin VB.Label m_lblGewichtLiter 
      Caption         =   "Liter = "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4500
      TabIndex        =   71
      Top             =   10125
      Width           =   765
   End
   Begin VB.Label Label6 
      Caption         =   "Kraftstoffgewicht in KG "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   195
      TabIndex        =   69
      Top             =   10125
      Width           =   2790
   End
   Begin VB.Label Label5 
      Caption         =   "Faktor Kilogramm je Liter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   195
      TabIndex        =   67
      Top             =   3375
      Width           =   2790
   End
   Begin VB.Label m_lblCalcAktuell3 
      Caption         =   "m_lblCalcAktuell3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5925
      TabIndex        =   65
      Top             =   825
      Width           =   1260
   End
   Begin VB.Label m_lblCalcAktuell2 
      Caption         =   "m_lblCalcAktuell2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4500
      TabIndex        =   64
      Top             =   825
      Width           =   1260
   End
   Begin VB.Label m_lblCalcAktuell1 
      Caption         =   "m_lblCalcAktuell1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3075
      TabIndex        =   63
      Top             =   825
      Width           =   1260
   End
   Begin VB.Line LineRezise 
      X1              =   776
      X2              =   1001
      Y1              =   42
      Y2              =   607
   End
   Begin VB.Label Label4 
      Caption         =   "Std:Min"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   4680
      TabIndex        =   51
      Top             =   3870
      Width           =   870
   End
   Begin VB.Label Label3 
      Caption         =   "Flugzeit Minuten gesamt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   195
      TabIndex        =   40
      Top             =   8970
      Width           =   2790
   End
   Begin VB.Label m_lblAnAbflugMinuten 
      Caption         =   "An/Abflug (min. 10 Minuten)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   195
      TabIndex        =   35
      Top             =   5910
      Width           =   2790
   End
   Begin VB.Label Label2 
      Caption         =   "Kraftstoff Liter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5775
      TabIndex        =   34
      Top             =   3870
      Width           =   1380
   End
   Begin VB.Label Label1 
      Caption         =   "Flugzeit Minuten"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2865
      TabIndex        =   33
      Top             =   3870
      Width           =   1770
   End
   Begin VB.Label m_lblReisegeschwKmh 
      Caption         =   "Reisegeschw Kmh"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   195
      TabIndex        =   0
      Top             =   1275
      Width           =   2790
   End
   Begin VB.Label m_lblReisegeschwLiter 
      Caption         =   "Reisegeschw Liter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   195
      TabIndex        =   2
      Top             =   1785
      Width           =   2790
   End
   Begin VB.Label m_lblReisegeschwLiterJeMinute 
      Caption         =   "Liter je Minute"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   195
      TabIndex        =   4
      Top             =   2295
      Width           =   2790
   End
   Begin VB.Label m_lblReisegeschwMinuteJeLiter 
      Caption         =   "Flugminuten je Liter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   195
      TabIndex        =   6
      Top             =   2805
      Width           =   2790
   End
   Begin VB.Label m_lblReiseflugMinuten 
      Caption         =   "P23 - Reiseflug"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   195
      TabIndex        =   8
      Top             =   4380
      Width           =   2790
   End
   Begin VB.Label m_lblZuschlagAnlassenMinuten 
      Caption         =   "P24 - Zuschlag Anlassen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   195
      TabIndex        =   11
      Top             =   4890
      Width           =   2790
   End
   Begin VB.Label m_lblSteigflugMinuten 
      Caption         =   "P25 - Steigflug"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   195
      TabIndex        =   14
      Top             =   5400
      Width           =   2790
   End
   Begin VB.Label m_lblAusweichplatzMinuten 
      Caption         =   "P20 - Ausweichplatz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   195
      TabIndex        =   17
      Top             =   6420
      Width           =   2790
   End
   Begin VB.Label m_lblReserveMinuten 
      Caption         =   "Reserve (min. 30 Minuten)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   195
      TabIndex        =   20
      Top             =   6930
      Width           =   2790
   End
   Begin VB.Label m_lblMindestKraftstoffbedarfLiter 
      Caption         =   "Mindest Kraftstoffbedarf"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   195
      TabIndex        =   23
      Top             =   7440
      Width           =   2790
   End
   Begin VB.Label m_lblExtraKraftstoffMinuten 
      Caption         =   "Extra Kraftstoff"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   195
      TabIndex        =   25
      Top             =   7950
      Width           =   2790
   End
   Begin VB.Label m_lblKraftstoffvorratMinuten 
      Caption         =   "P26 - Kraftstoffvorrat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   195
      TabIndex        =   28
      Top             =   8460
      Width           =   2790
   End
   Begin VB.Label m_lblSichereFlugzeitMinuten 
      Caption         =   "P27 - Sichere Flugzeit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   195
      TabIndex        =   31
      Top             =   9480
      Width           =   2790
   End
End
Attribute VB_Name = "m_frmVbKraftstoffberechnung"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'
' Konstanten fuer die Form_Resize-Funktion
'
Private Const abstand_y = 20
Private Const abstand_x = 20

Private Const START_CALC_DATEN_1 = 1
Private Const START_CALC_DATEN_2 = 2
Private Const START_CALC_DATEN_3 = 3

Private Const INI_SECTION_KRAFTSTOFFBERECHNUNG = "Kraftstoffberechnung"

Private Const INI_SCHLUESSEL_FAKTOR_GEWICHT_KRAFTSTOFF = "FaktorGewichtKraffstoffJeLiter"

Private Const INI_SCHLUESSEL_REISEGESCHW_KMH_1 = "ReisegeschwKmh1"
Private Const INI_SCHLUESSEL_REISEGESCHW_KMH_2 = "ReisegeschwKmh2"
Private Const INI_SCHLUESSEL_REISEGESCHW_KMH_3 = "ReisegeschwKmh3"

Private Const INI_SCHLUESSEL_REISEGESCHW_LITER_1 = "ReisegeschwLiter1"
Private Const INI_SCHLUESSEL_REISEGESCHW_LITER_2 = "ReisegeschwLiter2"
Private Const INI_SCHLUESSEL_REISEGESCHW_LITER_3 = "ReisegeschwLiter3"

Private Const INI_SCHLUESSEL_REISEGESCHW_LITER_1_JE_MINUTE = "ReisegeschwLiterJeMinute"
Private Const INI_SCHLUESSEL_REISEGESCHW_MINUTE_JE_LITER = "ReisegeschwMinuteJeLiter"
Private Const INI_SCHLUESSEL_REISEFLUG_MINUTEN = "ReiseflugMinuten"
Private Const INI_SCHLUESSEL_REISEFLUG_LITER = "ReiseflugLiter"
Private Const INI_SCHLUESSEL_ZUSCHLAG_ANLASSEN_MINUTEN = "ZuschlagAnlassenMinuten"
Private Const INI_SCHLUESSEL_ZUSCHLAG_ANLASSEN_LITER = "ZuschlagAnlassenLiter"
Private Const INI_SCHLUESSEL_STEIGFLUG_MINUTEN = "SteigflugMinuten"
Private Const INI_SCHLUESSEL_STEIGFLUG_LITER = "SteigflugLiter"
Private Const INI_SCHLUESSEL_AN_ABFLUG_MINUTEN = "AnAbflugMinuten"
Private Const INI_SCHLUESSEL_AN_ABFLUG_LITER = "AnAbflugLiter"
Private Const INI_SCHLUESSEL_AUSWEICHPLATZ_MINUTEN = "AusweichplatzMinuten"
Private Const INI_SCHLUESSEL_AUSWEICHPLATZ_LITER = "AusweichplatzLiter"
Private Const INI_SCHLUESSEL_RESERVE_MINUTEN = "ReserveMinuten"
Private Const INI_SCHLUESSEL_RESERVE_LITER = "ReserveLiter"
Private Const INI_SCHLUESSEL_MINDEST_KRAFTSTOFFBEDARF_LITER = "MindestKraftstoffbedarfLiter"
Private Const INI_SCHLUESSEL_EXTRA_KRAFTSTOFF_MINUTEN = "ExtraKraftstoffMinuten"
Private Const INI_SCHLUESSEL_EXTRA_KRAFTSTOFF_LITER = "ExtraKraftstoffLiter"
Private Const INI_SCHLUESSEL_KRAFTSTOFFVORRAT_MINUTEN = "KraftstoffvorratMinuten"
Private Const INI_SCHLUESSEL_KRAFTSTOFFVORRAT_LITER = "KraftstoffvorratLiter"
Private Const INI_SCHLUESSEL_SICHERE_FLUGZEIT_MINUTEN = "SichereFlugzeitMinuten"

Private knz_ui_resize_laeuft As Boolean


'################################################################################
'
Private Sub Form_Load()

On Error Resume Next

    '
    ' Initialisieren der Eingabefelder aus der INI-Datei
    '
    m_tbReisegeschwKmh1.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_REISEGESCHW_KMH_1, "125")
    m_tbReisegeschwLiter1.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_REISEGESCHW_LITER_1, "11.8")
    
    m_tbReisegeschwKmh2.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_REISEGESCHW_KMH_2, "140")
    m_tbReisegeschwLiter2.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_REISEGESCHW_LITER_2, "14.6")
    
    m_tbReisegeschwKmh3.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_REISEGESCHW_KMH_3, "165")
    m_tbReisegeschwLiter3.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_REISEGESCHW_LITER_3, "18.0")

    m_txtFaktorKilogrammJeLiter.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_FAKTOR_GEWICHT_KRAFTSTOFF, "0.73")
    
    m_tbReiseflugMinuten.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_REISEFLUG_MINUTEN, "48")
    m_tbZuschlagAnlassenMinuten.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_ZUSCHLAG_ANLASSEN_MINUTEN, "1")
    m_tbSteigflugMinuten.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_STEIGFLUG_MINUTEN, "5")
    m_tbAusweichplatzMinuten.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_AUSWEICHPLATZ_MINUTEN, "7")
    m_tbReserveMinuten.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_RESERVE_MINUTEN, "30")
    m_tbKraftstoffvorratLiter.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_KRAFTSTOFFVORRAT_LITER, "10")
    
    m_tbReiseflugLiter.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_REISEFLUG_LITER, "")
    m_tbZuschlagAnlassenLiter.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_ZUSCHLAG_ANLASSEN_LITER, "")
    m_tbSteigflugLiter.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_STEIGFLUG_LITER, "")
    m_tbAusweichplatzLiter.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_AUSWEICHPLATZ_LITER, "")
    m_tbReserveLiter.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_RESERVE_LITER, "")
    m_tbMindestKraftstoffbedarfLiter.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_MINDEST_KRAFTSTOFFBEDARF_LITER, "")
    m_tbExtraKraftstoffMinuten.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_EXTRA_KRAFTSTOFF_MINUTEN, "")
    m_tbExtraKraftstoffLiter.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_EXTRA_KRAFTSTOFF_LITER, "")
    m_tbKraftstoffvorratMinuten.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_KRAFTSTOFFVORRAT_MINUTEN, "")
    m_tbSichereFlugzeitMinuten.Text = readIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_SICHERE_FLUGZEIT_MINUTEN, "")

    m_lblCalcAktuell1.Caption = ""
    m_lblCalcAktuell2.Caption = ""
    m_lblCalcAktuell3.Caption = ""
    
    m_tbReisegeschwLiterJeMinute1.Text = "-"
    m_tbReisegeschwMinuteJeLiter1.Text = "-"
    
    m_tbReisegeschwLiterJeMinute2.Text = "-"
    m_tbReisegeschwMinuteJeLiter2.Text = "-"
    
    m_tbReisegeschwLiterJeMinute3.Text = "-"
    m_tbReisegeschwMinuteJeLiter3.Text = "-"

    '
    ' Initialisieren der Eingabefelder
    '
    m_tbAnAbflugMinuten.Text = "10"
    m_tbAnAbflugLiter.Text = ""

    m_txtStdMinFlugzeitGesamt.Text = ""
    m_txtStdMinReiseflug.Text = ""
    m_txtStdMinZuschlagAnlassen.Text = ""
    m_txtStdMinSteigflug.Text = ""
    m_txtStdMinAnAbflug.Text = ""
    m_txtStdMinAuschweichplatz.Text = ""
    m_txtStdMinReserve.Text = ""
    m_txtStdMinKraftstoffExtra.Text = ""
    m_txtStdMinKraftstoffVorrat.Text = ""
    m_txtStdMinSichereFlugzeit.Text = ""
    
    m_txtLog.Text = ""

    knz_ui_resize_laeuft = False

    Call startCalc(START_CALC_DATEN_1)

End Sub

'################################################################################
'
Private Sub Form_UnLoad(Cancel As Integer)

On Error Resume Next

    '
    ' Schreiben der Werte aus den Eingabefeldern in die INI-Datei
    '
    writeToIniKraftstoffberechnung

End Sub

'################################################################################
'
Private Sub Form_Resize()

On Error Resume Next

    '
    ' Pruefung: Laeuft Resize schon?
    ' Wird Resize-Funktion bereits ausgefuehrt, soll kein
    ' nebenlaeufiger Resize ausgefuehrt werden.
    '
    If (knz_ui_resize_laeuft = False) Then

        knz_ui_resize_laeuft = True

        If (Me.ScaleWidth > LineRezise.X2) Then

            m_txtLog.Width = Me.ScaleWidth - (m_txtLog.Left + 5)
        
        End If

        If (Me.ScaleHeight > LineRezise.Y2) Then

            m_txtLog.Height = Me.ScaleHeight - (m_txtLog.Top + 5)

        End If

        knz_ui_resize_laeuft = False

    End If

End Sub

'################################################################################
'
Private Function writeToIniKraftstoffberechnung() As Boolean

On Error GoTo errWriteToIniKraftstoffberechnung

    writeToIniKraftstoffberechnung = False

    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_REISEGESCHW_KMH_1, m_tbReisegeschwKmh1.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_REISEGESCHW_KMH_2, m_tbReisegeschwKmh2.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_REISEGESCHW_KMH_3, m_tbReisegeschwKmh3.Text)
    
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_REISEGESCHW_LITER_1, m_tbReisegeschwLiter1.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_REISEGESCHW_LITER_2, m_tbReisegeschwLiter2.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_REISEGESCHW_LITER_3, m_tbReisegeschwLiter3.Text)

    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_FAKTOR_GEWICHT_KRAFTSTOFF, m_txtFaktorKilogrammJeLiter.Text)
    
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_REISEGESCHW_LITER_1_JE_MINUTE, m_tbReisegeschwLiterJeMinute1.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_REISEGESCHW_MINUTE_JE_LITER, m_tbReisegeschwMinuteJeLiter1.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_REISEFLUG_MINUTEN, m_tbReiseflugMinuten.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_REISEFLUG_LITER, m_tbReiseflugLiter.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_ZUSCHLAG_ANLASSEN_MINUTEN, m_tbZuschlagAnlassenMinuten.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_ZUSCHLAG_ANLASSEN_LITER, m_tbZuschlagAnlassenLiter.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_STEIGFLUG_MINUTEN, m_tbSteigflugMinuten.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_STEIGFLUG_LITER, m_tbSteigflugLiter.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_AUSWEICHPLATZ_MINUTEN, m_tbAusweichplatzMinuten.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_AUSWEICHPLATZ_LITER, m_tbAusweichplatzLiter.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_RESERVE_MINUTEN, m_tbReserveMinuten.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_RESERVE_LITER, m_tbReserveLiter.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_MINDEST_KRAFTSTOFFBEDARF_LITER, m_tbMindestKraftstoffbedarfLiter.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_EXTRA_KRAFTSTOFF_MINUTEN, m_tbExtraKraftstoffMinuten.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_EXTRA_KRAFTSTOFF_LITER, m_tbExtraKraftstoffLiter.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_KRAFTSTOFFVORRAT_MINUTEN, m_tbKraftstoffvorratMinuten.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_KRAFTSTOFFVORRAT_LITER, m_tbKraftstoffvorratLiter.Text)
    Call writeIniText(INI_SECTION_KRAFTSTOFFBERECHNUNG, INI_SCHLUESSEL_SICHERE_FLUGZEIT_MINUTEN, m_tbSichereFlugzeitMinuten.Text)

    writeToIniKraftstoffberechnung = True

EndFunktion:

    On Error Resume Next

    DoEvents

    Exit Function

errWriteToIniKraftstoffberechnung:

    'Call wl("Fehler: errWriteToIniKraftstoffberechnung: " & Err & " " & Error & " " & Erl)

    Resume EndFunktion

End Function

'################################################################################
'
Private Sub startCalc(pStartCalc As Integer)

On Error GoTo errStartCalc

Dim an_abflug_liter                As Double
Dim an_abflug_minuten              As Long
Dim ausweichplatz_liter            As Double
Dim ausweichplatz_minuten          As Long
Dim breite_liter_spalte            As Integer
Dim breite_minuten_feld            As Integer
Dim dbl_60_minuten                 As Double
Dim extra_kraftstoff_liter         As Double
Dim extra_kraftstoff_minuten       As Long
Dim flugzeit_gesamt_minuten        As Long
Dim format_float_spalte            As String
Dim kraftstoffvorrat_liter         As Double
Dim kraftstoffvorrat_minuten       As Long
Dim mindest_kraftstoffbedarf_liter As Double
Dim my_chr                         As String
Dim reiseflug_liter                As Double
Dim reiseflug_minuten              As Long
Dim reisegeschw_kmh                As Double
Dim reisegeschw_liter              As Double
Dim reisegeschw_liter_je_minute    As Double
Dim reisegeschw_minute_je_liter    As Double
Dim reserve_liter                  As Double
Dim reserve_minuten                As Long
Dim sichere_flugzeit_abzug_minuten As Long
Dim sichere_flugzeit_minuten       As Long
Dim steigflug_liter                As Double
Dim steigflug_minuten              As Long
Dim zuschlag_anlassen_liter        As Double
Dim zuschlag_anlassen_minuten      As Long
Dim str_berechnung                 As String
Dim gewicht_je_liter_kraftstoff_faktor As Double
Dim gewicht_je_liter_kraftstoff_aktuell As Double

Dim str_stdmin_reiseflug           As String
Dim str_stdmin_zuschlag_anlassen   As String
Dim str_stdmin_steigflug           As String
Dim str_stdmin_an_abflug           As String
Dim str_stdmin_ausweichplatz       As String
Dim str_stdmin_reserve             As String
Dim str_stdmin_extra_kraftstoff    As String
Dim str_stdmin_kraftstoffvorrat    As String
Dim str_stdmin_sichere_flugzeit    As String
Dim str_stdmin_flugzeit_gesamt     As String
    
    my_chr = vbCrLf

    '
    ' Datenuebernahme
    ' Die Daten aus den Eingabefeldern werden in lokale Variablen uebertragen.
    '
    zuschlag_anlassen_minuten = getInteger(m_tbZuschlagAnlassenMinuten.Text, 0)
    
    steigflug_minuten = getInteger(m_tbSteigflugMinuten.Text, 0)
    
    an_abflug_minuten = getInteger(m_tbAnAbflugMinuten.Text, 0)
    
    reiseflug_minuten = getInteger(m_tbReiseflugMinuten.Text, 0)
    
    ausweichplatz_minuten = getInteger(m_tbAusweichplatzMinuten.Text, 0)
    
    gewicht_je_liter_kraftstoff_faktor = getDouble(m_txtFaktorKilogrammJeLiter.Text, 0)
    
    extra_kraftstoff_liter = getDouble(m_tbExtraKraftstoffLiter.Text, 0)
    extra_kraftstoff_minuten = getInteger(m_tbExtraKraftstoffMinuten.Text, 0)
    kraftstoffvorrat_liter = getDouble(m_tbKraftstoffvorratLiter.Text, 0)
    kraftstoffvorrat_minuten = getInteger(m_tbKraftstoffvorratMinuten.Text, 0)
    mindest_kraftstoffbedarf_liter = getDouble(m_tbMindestKraftstoffbedarfLiter.Text, 0)
    reserve_minuten = getInteger(m_tbReserveMinuten.Text, 0)
    
    
    If (pStartCalc = START_CALC_DATEN_2) Then
        
        reisegeschw_kmh = getDouble(m_tbReisegeschwKmh2.Text, 0)
        reisegeschw_liter = getDouble(m_tbReisegeschwLiter2.Text, 0)
        
    ElseIf (pStartCalc = START_CALC_DATEN_3) Then
        
        reisegeschw_kmh = getDouble(m_tbReisegeschwKmh3.Text, 0)
        reisegeschw_liter = getDouble(m_tbReisegeschwLiter3.Text, 0)
        
    Else 'If (pStartCalc = START_CALC_DATEN_1) Then
        
        reisegeschw_kmh = getDouble(m_tbReisegeschwKmh1.Text, 0)
        reisegeschw_liter = getDouble(m_tbReisegeschwLiter1.Text, 0)
        
    End If
    
    '
    ' Berechnungen
    '

    dbl_60_minuten = 60#
    
    reisegeschw_liter_je_minute = get5nk(reisegeschw_liter / dbl_60_minuten)
    
    reisegeschw_minute_je_liter = get5nk(dbl_60_minuten / reisegeschw_liter)
    
    Dim knz_berechnung_nach_a As Boolean
    
    knz_berechnung_nach_a = False
    
    '
    ' Es kann zu Rundungsdifferenzen zwischen den beiden Verfahren kommen.
    '
    ' Wird nur mit "reisegeschw_liter_je_minute" multipliziert, koennen diese Differenzen auftreten.
    '
    ' Darum ist die ausfuehrlichere Variante korrekter.
    '
    If (knz_berechnung_nach_a) Then
    
        reiseflug_liter = get5nk((reisegeschw_liter * reiseflug_minuten) / dbl_60_minuten)
    
        zuschlag_anlassen_liter = get5nk((reisegeschw_liter * zuschlag_anlassen_minuten) / dbl_60_minuten)
    
        steigflug_liter = get5nk((reisegeschw_liter * steigflug_minuten) / dbl_60_minuten)
    
        an_abflug_liter = get5nk((reisegeschw_liter * an_abflug_minuten) / dbl_60_minuten)
    
        ausweichplatz_liter = get5nk((reisegeschw_liter * ausweichplatz_minuten) / dbl_60_minuten)
    
        reserve_liter = get5nk((reisegeschw_liter * reserve_minuten) / dbl_60_minuten)

    Else
    
        reiseflug_liter = get5nk(reisegeschw_liter_je_minute * reiseflug_minuten)
        
        zuschlag_anlassen_liter = get5nk(reisegeschw_liter_je_minute * zuschlag_anlassen_minuten)
        
        steigflug_liter = get5nk(reisegeschw_liter_je_minute * steigflug_minuten)
        
        an_abflug_liter = get5nk(reisegeschw_liter_je_minute * an_abflug_minuten)
        
        ausweichplatz_liter = get5nk(reisegeschw_liter_je_minute * ausweichplatz_minuten)
        
        reserve_liter = get5nk(reisegeschw_liter_je_minute * reserve_minuten)
    
    End If


    mindest_kraftstoffbedarf_liter = get5nk(reiseflug_liter + zuschlag_anlassen_liter + steigflug_liter + an_abflug_liter + ausweichplatz_liter + reserve_liter)

    If (kraftstoffvorrat_liter >= mindest_kraftstoffbedarf_liter) Then
    
        extra_kraftstoff_liter = 0
        
        extra_kraftstoff_minuten = 0
        
    Else
    
        extra_kraftstoff_liter = get5nk(mindest_kraftstoffbedarf_liter - kraftstoffvorrat_liter)
        
        extra_kraftstoff_minuten = get5nk(reisegeschw_minute_je_liter * extra_kraftstoff_liter)
        
    End If

    kraftstoffvorrat_minuten = get5nk(reisegeschw_minute_je_liter * kraftstoffvorrat_liter)

    flugzeit_gesamt_minuten = extra_kraftstoff_minuten + kraftstoffvorrat_minuten
    
    m_tbFlugzeitGesamtMinuten.Text = flugzeit_gesamt_minuten
    
    sichere_flugzeit_abzug_minuten = 30
    
    sichere_flugzeit_minuten = flugzeit_gesamt_minuten - sichere_flugzeit_abzug_minuten

    If (sichere_flugzeit_minuten < sichere_flugzeit_abzug_minuten) Then
    
        m_tbSichereFlugzeitMinuten.BackColor = vbRed
    
    Else
    
        m_tbSichereFlugzeitMinuten.BackColor = vbWhite
    
    End If
    
    Dim minuten_x As Long
    
    sichere_flugzeit_abzug_minuten = 30
    
    minuten_x = reiseflug_minuten + zuschlag_anlassen_minuten + steigflug_minuten + an_abflug_minuten + ausweichplatz_minuten + reserve_minuten
    
    str_stdmin_flugzeit_gesamt = getStringStundenMinuten(flugzeit_gesamt_minuten)
    str_stdmin_reiseflug = getStringStundenMinuten(reiseflug_minuten)
    str_stdmin_zuschlag_anlassen = getStringStundenMinuten(zuschlag_anlassen_minuten)
    str_stdmin_steigflug = getStringStundenMinuten(steigflug_minuten)
    str_stdmin_an_abflug = getStringStundenMinuten(an_abflug_minuten)
    str_stdmin_ausweichplatz = getStringStundenMinuten(ausweichplatz_minuten)
    str_stdmin_reserve = getStringStundenMinuten(reserve_minuten)
    str_stdmin_extra_kraftstoff = getStringStundenMinuten(extra_kraftstoff_minuten)
    str_stdmin_kraftstoffvorrat = getStringStundenMinuten(kraftstoffvorrat_minuten)
    str_stdmin_sichere_flugzeit = getStringStundenMinuten(sichere_flugzeit_minuten)

    breite_minuten_feld = getAnzahlStellen(reiseflug_minuten)
    breite_minuten_feld = getMaxInteger(breite_minuten_feld, getAnzahlStellen(zuschlag_anlassen_minuten))
    breite_minuten_feld = getMaxInteger(breite_minuten_feld, getAnzahlStellen(steigflug_minuten))
    breite_minuten_feld = getMaxInteger(breite_minuten_feld, getAnzahlStellen(an_abflug_minuten))
    breite_minuten_feld = getMaxInteger(breite_minuten_feld, getAnzahlStellen(ausweichplatz_minuten))
    breite_minuten_feld = getMaxInteger(breite_minuten_feld, getAnzahlStellen(reserve_minuten))
    
    breite_liter_spalte = 10
    format_float_spalte = FORMAT_FLOAT5

    str_berechnung = str_berechnung & my_chr & "Kraftstoffberechnung"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "1. Verbrauchsangaben aus Handbuch ermitteln"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "    Reisegeschwindigkeit " & reisegeschw_kmh & " kmh = " & reisegeschw_liter & " Liter die Stunde"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "    " & reisegeschw_liter & " Liter Verbrauch je Stunde / 60 Minuten = " & reisegeschw_liter_je_minute & " Liter Verbrauch je Minute (in Reisegeschwindigkeit)"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "    60 Minuten / " & reisegeschw_liter & " Liter Verbrauch je Stunde = " & reisegeschw_minute_je_liter & " Flugminuten je Liter (in Reisegeschwindigkeit)"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "    Mit jedem Liter kann " & reisegeschw_minute_je_liter & " Minuten geflogen werden."
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "    In einer Flugminute werden " & reisegeschw_liter_je_minute & " Liter verbraucht."
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "-------------------------------------------------------------------------------"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "2. Aus den ermittelten Flugzeiten, die benötigten Liter Kraftstoff berechnen"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "     60 Minuten = " & reisegeschw_liter & " Liter"
    str_berechnung = str_berechnung & my_chr & "      x Minuten =    y Liter"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "      " & getFeldRechtsMinInteger(reiseflug_minuten, breite_minuten_feld) & " Minuten Reiseflug      = " & reisegeschw_liter & " * " & getFeldRechtsMinInteger(reiseflug_minuten, breite_minuten_feld) & " Min / 60 Min = " & getFeldRechtsMin(Format(reiseflug_liter, format_float_spalte), breite_liter_spalte) & " Liter (P23)"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "    + " & getFeldRechtsMinInteger(zuschlag_anlassen_minuten, breite_minuten_feld) & " Minuten Anlassen       = " & reisegeschw_liter & " * " & getFeldRechtsMinInteger(zuschlag_anlassen_minuten, breite_minuten_feld) & " Min / 60 Min = " & getFeldRechtsMin(Format(zuschlag_anlassen_liter, format_float_spalte), breite_liter_spalte) & " Liter (P24)"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "    + " & getFeldRechtsMinInteger(steigflug_minuten, breite_minuten_feld) & " Minuten Steigflug      = " & reisegeschw_liter & " * " & getFeldRechtsMinInteger(steigflug_minuten, breite_minuten_feld) & " Min / 60 Min = " & getFeldRechtsMin(Format(steigflug_liter, format_float_spalte), breite_liter_spalte) & " Liter (P25)"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "    + " & getFeldRechtsMinInteger(an_abflug_minuten, breite_minuten_feld) & " Minuten An- und Abflug = " & reisegeschw_liter & " * " & getFeldRechtsMinInteger(an_abflug_minuten, breite_minuten_feld) & " Min / 60 Min = " & getFeldRechtsMin(Format(an_abflug_liter, format_float_spalte), breite_liter_spalte) & " Liter"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "    + " & getFeldRechtsMinInteger(ausweichplatz_minuten, breite_minuten_feld) & " Minuten Ausweich-Platz = " & reisegeschw_liter & " * " & getFeldRechtsMinInteger(ausweichplatz_minuten, breite_minuten_feld) & " Min / 60 Min = " & getFeldRechtsMin(Format(ausweichplatz_liter, format_float_spalte), breite_liter_spalte) & " Liter (P20)"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "    + " & getFeldRechtsMinInteger(reserve_minuten, breite_minuten_feld) & " Minuten Reserve        = " & reisegeschw_liter & " * " & getFeldRechtsMinInteger(reserve_minuten, breite_minuten_feld) & " Min / 60 Min = " & getFeldRechtsMin(Format(reserve_liter, format_float_spalte), breite_liter_spalte) & " Liter"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "   Mindest Kraftstoffbedarf sind " & Format(mindest_kraftstoffbedarf_liter, format_float_spalte) & " Liter "
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "-------------------------------------------------------------------------------"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "3. Kraftstoffvorrat (P26)"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "   = Wieviele Liter sind noch im Tank und wie lange kann ich damit fliegen?"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "   " & reisegeschw_minute_je_liter & " Flugminuten je Liter * " & kraftstoffvorrat_liter & " Liter im Tank"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "   Kraftstoffvorrat " & Format(kraftstoffvorrat_liter, format_float_spalte) & " Liter = " & getFeldRechtsMinInteger(kraftstoffvorrat_minuten, breite_minuten_feld) & " Flugminuten "
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "-------------------------------------------------------------------------------"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "4. ""Extra Kraftstoff"" berechnen (Was muss mindestens getankt werden)"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "   Extra Kraftstoff = Differenz zu Mindestbedarf und Kraftstoffvorrat (was im Tank ist)"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "   " & Format(mindest_kraftstoffbedarf_liter, format_float_spalte) & " Liter Mindestbedarf - " & Format(kraftstoffvorrat_liter, format_float_spalte) & " Liter Kraftstoffvorrat = " & Format(extra_kraftstoff_liter, format_float_spalte) & " Liter"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "   " & reisegeschw_minute_je_liter & " Flugminuten je Liter * " & extra_kraftstoff_liter & " Extra Liter zu tanken = " & getFeldRechtsMinInteger(extra_kraftstoff_minuten, breite_minuten_feld) & " Flugminuten "
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "   Extra Kraftstoffbedarf sind " & Format(extra_kraftstoff_liter, format_float_spalte) & " Liter = " & getFeldRechtsMinInteger(extra_kraftstoff_minuten, breite_minuten_feld) & " Flugminuten "
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "-------------------------------------------------------------------------------"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "5. Sichere Flugzeit berechnen"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "   1. Schritt: Alle Flugminuten zusammenrechnen"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "        + " & getFeldRechtsMinInteger(reiseflug_minuten, breite_minuten_feld) & " Minuten Reiseflug      = " & str_stdmin_reiseflug
    str_berechnung = str_berechnung & my_chr & "        + " & getFeldRechtsMinInteger(zuschlag_anlassen_minuten, breite_minuten_feld) & " Minuten Anlassen       = " & str_stdmin_zuschlag_anlassen
    str_berechnung = str_berechnung & my_chr & "        + " & getFeldRechtsMinInteger(steigflug_minuten, breite_minuten_feld) & " Minuten Steigflug      = " & str_stdmin_steigflug
    str_berechnung = str_berechnung & my_chr & "        + " & getFeldRechtsMinInteger(an_abflug_minuten, breite_minuten_feld) & " Minuten An- und Abflug = " & str_stdmin_an_abflug
    str_berechnung = str_berechnung & my_chr & "        + " & getFeldRechtsMinInteger(ausweichplatz_minuten, breite_minuten_feld) & " Minuten Ausweich-Platz = " & str_stdmin_ausweichplatz
    str_berechnung = str_berechnung & my_chr & "        + " & getFeldRechtsMinInteger(reserve_minuten, breite_minuten_feld) & " Minuten Reserve        = " & str_stdmin_reserve
    str_berechnung = str_berechnung & my_chr & "          --------------------------------------"
    str_berechnung = str_berechnung & my_chr & "        = " & getFeldRechtsMinInteger(minuten_x, breite_minuten_feld) & " Gesamtminuten Flugzeit = " & getStringStundenMinuten(minuten_x)
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "          " & getFeldRechtsMinInteger(extra_kraftstoff_minuten, breite_minuten_feld) & " Minuten extra Kraftstoff  = " & str_stdmin_extra_kraftstoff & " (getankte Liter)"
    str_berechnung = str_berechnung & my_chr & "        + " & getFeldRechtsMinInteger(kraftstoffvorrat_minuten, breite_minuten_feld) & " Minuten Kraftstoffreserve = " & str_stdmin_kraftstoffvorrat & " (Tankinhalt)"
    str_berechnung = str_berechnung & my_chr & "          --------------------------------------"
    str_berechnung = str_berechnung & my_chr & "        = " & getFeldRechtsMinInteger(flugzeit_gesamt_minuten, breite_minuten_feld) & " Gesamtminuten Flugzeit     = " & str_stdmin_flugzeit_gesamt
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "   2. Schritt: " & sichere_flugzeit_abzug_minuten & " Minuten von der Gesamtflugzeit abziehen"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "          " & getFeldRechtsMinInteger(flugzeit_gesamt_minuten, breite_minuten_feld) & " Gesamtminuten Flugzeit = " & str_stdmin_flugzeit_gesamt
    str_berechnung = str_berechnung & my_chr & "        - " & getFeldRechtsMinInteger(sichere_flugzeit_abzug_minuten, breite_minuten_feld) & " Minuten Abzug          = " & getStringStundenMinuten(sichere_flugzeit_abzug_minuten)
    str_berechnung = str_berechnung & my_chr & "          --------------------------------------"
    str_berechnung = str_berechnung & my_chr & "        = " & getFeldRechtsMinInteger(sichere_flugzeit_minuten, breite_minuten_feld) & " Gesamtminuten Flugzeit = " & str_stdmin_sichere_flugzeit
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "-------------------------------------------------------------------------------"
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "6. Massenberechnung fuer Kraftstoff"
    str_berechnung = str_berechnung & my_chr & ""
    
    Dim gewicht_kraftstoff_aktuell As Double
    Dim gewicht_kraftstoff_liter   As Double
    
    gewicht_kraftstoff_liter = (extra_kraftstoff_liter + kraftstoffvorrat_liter)
    
    m_txtGewichtAnzahlLiter.Text = get5nk(gewicht_kraftstoff_liter)

    gewicht_kraftstoff_aktuell = get5nk(gewicht_kraftstoff_liter * gewicht_je_liter_kraftstoff_faktor)
    
    m_txtKraftstoffgewichtKg.Text = get5nk(gewicht_kraftstoff_aktuell)
    
    str_berechnung = str_berechnung & my_chr & "        Anrechnungsfaktor je Liter Kraftstoff nach Handbuch = " & get5nk(gewicht_je_liter_kraftstoff_faktor)
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "        = " & getFeldRechtsMinInteger(gewicht_kraftstoff_liter, breite_minuten_feld) & " Liter = " & gewicht_kraftstoff_aktuell & " KG "
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & ""
    
    
'Dim gewicht_je_liter_kraftstoff_faktor As Double
'Dim gewicht_kraftstoff_aktuell As Double
    
    str_berechnung = str_berechnung & my_chr & ""
    str_berechnung = str_berechnung & my_chr & "Flugdurchführungsplan VFR"
    str_berechnung = str_berechnung & my_chr & "http://fsvwaechtersberg.de/fpsl/download/flugplanlba-v.pdf"
    str_berechnung = str_berechnung & my_chr & ""

    m_lblCalcAktuell1.Caption = ""
    m_lblCalcAktuell2.Caption = ""
    m_lblCalcAktuell3.Caption = ""

    '
    ' Ergebniswerte und Eingabewerte in die Eingabefelder zurueckschreiben.
    ' (Aus einem leeren Minutenfeld wird eine 0)
    '
    If (pStartCalc = START_CALC_DATEN_2) Then
        
        m_tbReisegeschwLiterJeMinute2.Text = "" & reisegeschw_liter_je_minute
        m_tbReisegeschwMinuteJeLiter2.Text = "" & reisegeschw_minute_je_liter
        
        m_lblCalcAktuell2.Caption = "Aktuell"

        
    ElseIf (pStartCalc = START_CALC_DATEN_3) Then
        
        m_tbReisegeschwLiterJeMinute3.Text = "" & reisegeschw_liter_je_minute
        m_tbReisegeschwMinuteJeLiter3.Text = "" & reisegeschw_minute_je_liter
        
        m_lblCalcAktuell3.Caption = "Aktuell"

    Else 'If (pStartCalc = START_CALC_DATEN_1) Then
        
        m_tbReisegeschwLiterJeMinute1.Text = "" & reisegeschw_liter_je_minute
        m_tbReisegeschwMinuteJeLiter1.Text = "" & reisegeschw_minute_je_liter
        
        m_lblCalcAktuell1.Caption = "Aktuell"

    End If
    
    m_tbReiseflugMinuten.Text = "" & reiseflug_minuten
    m_tbReiseflugLiter.Text = "" & reiseflug_liter
    m_tbZuschlagAnlassenMinuten.Text = "" & zuschlag_anlassen_minuten
    m_tbZuschlagAnlassenLiter.Text = "" & zuschlag_anlassen_liter
    m_tbSteigflugMinuten.Text = "" & steigflug_minuten
    m_tbSteigflugLiter.Text = "" & steigflug_liter
    m_tbAnAbflugMinuten.Text = "" & an_abflug_minuten
    m_tbAnAbflugLiter.Text = "" & an_abflug_liter
    m_tbAusweichplatzMinuten.Text = "" & ausweichplatz_minuten
    m_tbAusweichplatzLiter.Text = "" & ausweichplatz_liter
    m_tbReserveMinuten.Text = "" & reserve_minuten
    m_tbReserveLiter.Text = "" & reserve_liter
    m_tbMindestKraftstoffbedarfLiter.Text = "" & mindest_kraftstoffbedarf_liter
    m_tbExtraKraftstoffMinuten.Text = "" & extra_kraftstoff_minuten
    m_tbExtraKraftstoffLiter.Text = "" & extra_kraftstoff_liter
    m_tbKraftstoffvorratMinuten.Text = "" & kraftstoffvorrat_minuten
    m_tbKraftstoffvorratLiter.Text = "" & kraftstoffvorrat_liter
    m_tbSichereFlugzeitMinuten.Text = "" & sichere_flugzeit_minuten
    
    '
    ' Stunden/Minutenfelder schreiben
    '
    m_txtStdMinFlugzeitGesamt.Text = str_stdmin_flugzeit_gesamt
    m_txtStdMinReiseflug.Text = str_stdmin_reiseflug
    m_txtStdMinZuschlagAnlassen.Text = str_stdmin_zuschlag_anlassen
    m_txtStdMinSteigflug.Text = str_stdmin_steigflug
    m_txtStdMinAnAbflug.Text = str_stdmin_an_abflug
    m_txtStdMinAuschweichplatz.Text = str_stdmin_ausweichplatz
    m_txtStdMinReserve.Text = str_stdmin_reserve
    m_txtStdMinKraftstoffExtra.Text = str_stdmin_extra_kraftstoff
    m_txtStdMinKraftstoffVorrat.Text = str_stdmin_kraftstoffvorrat
    m_txtStdMinSichereFlugzeit.Text = str_stdmin_sichere_flugzeit

    m_txtLog.Text = str_berechnung

EndFunktion:

    On Error Resume Next

    DoEvents

    Exit Sub

errStartCalc:

     m_txtLog.Text = "Fehler: errStartCalc: " & Err & " " & Error & " " & Erl & ""

    Resume EndFunktion

End Sub

'################################################################################
'
Private Sub m_btnStartGeschwindigkeit1_Click()

    Call startCalc(START_CALC_DATEN_1)
    
End Sub

'################################################################################
'
Private Sub m_btnStartGeschwindigkeit2_Click()

    Call startCalc(START_CALC_DATEN_2)
    
End Sub

'################################################################################
'
Private Sub m_btnStartGeschwindigkeit3_Click()

    Call startCalc(START_CALC_DATEN_3)
    
End Sub

'################################################################################
'
Private Sub m_tbReisegeschwKmh2_GotFocus()

    Call uiTextBoxSelectAll(m_tbReisegeschwKmh2)

End Sub

'################################################################################
'
Private Sub m_tbReisegeschwLiter2_GotFocus()

    Call uiTextBoxSelectAll(m_tbReisegeschwLiter2)

End Sub

'################################################################################
'
Private Sub m_tbReisegeschwKmh3_GotFocus()

    Call uiTextBoxSelectAll(m_tbReisegeschwKmh3)

End Sub

'################################################################################
'
Private Sub m_tbReisegeschwLiter3_GotFocus()

    Call uiTextBoxSelectAll(m_tbReisegeschwLiter3)

End Sub

'################################################################################
'
Private Sub m_tbReisegeschwKmh1_GotFocus()

    Call uiTextBoxSelectAll(m_tbReisegeschwKmh1)

End Sub

'################################################################################
'
Private Sub m_tbReisegeschwLiter1_GotFocus()

    Call uiTextBoxSelectAll(m_tbReisegeschwLiter1)

End Sub

'################################################################################
'
Private Sub m_tbReisegeschwLiterJeMinute_GotFocus()

    Call uiTextBoxSelectAll(m_tbReisegeschwLiterJeMinute1)

End Sub

'################################################################################
'
Private Sub m_tbReisegeschwMinuteJeLiter_GotFocus()

    Call uiTextBoxSelectAll(m_tbReisegeschwMinuteJeLiter1)

End Sub

'################################################################################
'
Private Sub m_tbReiseflugMinuten_GotFocus()

    Call uiTextBoxSelectAll(m_tbReiseflugMinuten)

End Sub

'################################################################################
'
Private Sub m_tbReiseflugLiter_GotFocus()

    Call uiTextBoxSelectAll(m_tbReiseflugLiter)

End Sub

'################################################################################
'
Private Sub m_tbZuschlagAnlassenMinuten_GotFocus()

    Call uiTextBoxSelectAll(m_tbZuschlagAnlassenMinuten)

End Sub

'################################################################################
'
Private Sub m_tbZuschlagAnlassenLiter_GotFocus()

    Call uiTextBoxSelectAll(m_tbZuschlagAnlassenLiter)

End Sub

'################################################################################
'
Private Sub m_tbSteigflugMinuten_GotFocus()

    Call uiTextBoxSelectAll(m_tbSteigflugMinuten)

End Sub

'################################################################################
'
Private Sub m_tbAnAbflugMinuten_GotFocus()

    Call uiTextBoxSelectAll(m_tbAnAbflugMinuten)

End Sub

'################################################################################
'
Private Sub m_tbAnAbflugLiter_GotFocus()

    Call uiTextBoxSelectAll(m_tbAnAbflugLiter)

End Sub

'################################################################################
'
Private Sub m_tbSteigflugLiter_GotFocus()

    Call uiTextBoxSelectAll(m_tbSteigflugLiter)

End Sub

'################################################################################
'
Private Sub m_tbAusweichplatzMinuten_GotFocus()

    Call uiTextBoxSelectAll(m_tbAusweichplatzMinuten)

End Sub

'################################################################################
'
Private Sub m_tbAusweichplatzLiter_GotFocus()

    Call uiTextBoxSelectAll(m_tbAusweichplatzLiter)

End Sub

'################################################################################
'
Private Sub m_tbReserveMinuten_GotFocus()

    Call uiTextBoxSelectAll(m_tbReserveMinuten)

End Sub

'################################################################################
'
Private Sub m_tbReserveLiter_GotFocus()

    Call uiTextBoxSelectAll(m_tbReserveLiter)

End Sub

'################################################################################
'
Private Sub m_tbMindestKraftstoffbedarfLiter_GotFocus()

    Call uiTextBoxSelectAll(m_tbMindestKraftstoffbedarfLiter)

End Sub

'################################################################################
'
Private Sub m_tbExtraKraftstoffMinuten_GotFocus()

    Call uiTextBoxSelectAll(m_tbExtraKraftstoffMinuten)

End Sub

'################################################################################
'
Private Sub m_tbExtraKraftstoffLiter_GotFocus()

    Call uiTextBoxSelectAll(m_tbExtraKraftstoffLiter)

End Sub

'################################################################################
'
Private Sub m_tbKraftstoffvorratMinuten_GotFocus()

    Call uiTextBoxSelectAll(m_tbKraftstoffvorratMinuten)

End Sub

'################################################################################
'
Private Sub m_tbKraftstoffvorratLiter_GotFocus()

    Call uiTextBoxSelectAll(m_tbKraftstoffvorratLiter)

End Sub

'################################################################################
'
Private Sub m_tbSichereFlugzeitMinuten_GotFocus()

    Call uiTextBoxSelectAll(m_tbSichereFlugzeitMinuten)

End Sub

'################################################################################
'
Private Sub uiTextBoxSelectAll(pControl As Control)

    pControl.SelStart = 0
    pControl.SelLength = Len(pControl.Text)

End Sub


