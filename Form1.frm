VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memeriksa Format Tanggal"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   3120
      Top             =   3480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   2880
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Timer1.Interval = 500  'Set property intervalnya
  Timer1.Enabled = True  'Aktifkan jika belum...
End Sub

Sub PeriksaTanggal()
Dim CekTanggal As String
Ulangi:
  CekTanggal = Date  'Tampung tanggal dalam bentuk
                     'string
  'Lakukan pemeriksaan format tanggal tersebut...
  If CekTanggal <> Format(Date, "dd/mm/yyyy") Then
     'Jika formatnya tidak sama dengan 'dd/mm/yyyy',
     'tampilkan pesan berikut...
     If MsgBox("Format tanggal di komputer Anda tidak sama dengan" & vbCrLf & "'dd/mm/yyyy'. Klik OK untuk mengganti melalui menu" & vbCrLf & _
           "Regional Settings pada tab Date di kotak isian" & vbCrLf & "'Short Date Style'. Ganti menjadi format:" & vbCrLf & "dd/mm/yyyy. Jika Anda tidak melakukannya, maka" & vbCrLf & "program tidak dapat dijalankan!", vbCritical + vbOKCancel, "Format Tanggal Tidak Sama Dengan 'dd/mm/yyyy'") = vbOK And CekTanggal <> Format(Date, "dd/mm/yyyy") Then
        'Tampilkan Regional Settings dari program, dan
        'langsung ke tab Date (Tab indeks ke-4)...
        Call Shell("rundll32.exe shell32.dll," & "Control_RunDLL INTL.CPL,,4", 1)
     Else
        End  'Jika tdk mau mengganti, langsung keluar
             'program...
     End If
     If MsgBox("Apakah Anda sudah selesai menggantinya?" & vbCrLf & "Klik Yes jika format sudah dd/mm/yyyy" & vbCrLf & "atau klik No jika belum.", bQuestion + vbYesNo, "Ubah Tanggal") = vbYes Then
      'Periksa lagi, apakah sudah diganti oleh User?
        If CekTanggal <> Format(Date, "dd/mm/yyyy") Then GoTo Ulangi
     Else 'Jika belum juga, kembali lagi dari awal di
          'atas
        GoTo Ulangi
     End If
  End If
End Sub

'Jika sebelumnya format tanggal sudah 'dd/mm/yyyy', 'Anda dapat mengubahnya dengan mengklik tombol 'Command1.
'Perhatikan reaksi apa yang terjadi dari program 'setelah Anda mengubah format tanggal menjadi format yg 'tidak sesuai dengan 'dd/mm/yyyy' atau Anda juga dapat 'mengubah formatnya dari Control Panel, dan perhatikan 'juga bagaimana reaksi program!!!

Private Sub Command1_Click()
  Call Shell("rundll32.exe shell32.dll," & _
             "Control_RunDLL INTL.CPL,,4", 1)
End Sub

'Anda mungkin bertanya, mengapa kita memeriksa format
'tanggal di prosedur Timer1_Timer selain di prosedur
'Form_Load di atas?
'Jawabnya tidak lain adalah untuk mengantisipasi jika 'pada saat program dijalankan, dilakukan perubahan 'format tanggal melalui menu Regional Settings di 'Control Panel oleh user atau melalui Command1 yang ada 'di program,maka ketika program diaktifkan kembali 'format tanggal menjadi sudah tidak sama lagi dengan 'dd/mm/yyyy' sehingga harus dilakukan pemeriksaan 'kembali setiap saat program diaktifkan melalui bantuan 'Timer1 yang dapat refresh setiap saat...

Private Sub Timer1_Timer()
  If CekTanggal <> Format(Date, "dd/mm/yyyy") Then
     PeriksaTanggal
  Else
     Exit Sub  'Timer1 harus tetap aktif...
  End If
End Sub


