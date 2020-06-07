VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Deskripsi: Meng-copy Array ke Array Lainnya. Dalam coding berikut,
'           akan dibuktikan bahwa kita dapat meng-copy isi suatu array
'           ke array lainnya. Tentu saja elemen array sumber harus
'           sudah diketahui, sementara array tujuan harus dideklarasikan
'           sebagai array dinamik, agar dapat dicopy dari array sumber.
'Pembuat  : Masino Sinaga (masino_sinaga@yahoo.com)
'Diupload : Kamis, 16 Mei 2002
'Persiapan: 1. Buat 1 Project baru dengan 1 Form dan 2 Commandbutton
'           2. Copy-kan coding berikut ke dalam editor form yang bertalian.
'--------------------------------------------------------------------------

'Deklarasi arr1 statik sebanyak 3 elemen
Dim arr1(2) As Byte

'Deklarasi arr2 sebagai array dinamik, agar nantinya
'bisa dideklarasi ulang...
Dim arr2() As Byte

'Deklarasi arr3 statik sebanyak 3 elemen
Dim arr3(2) As Byte

'Deklarasi arr4 sebagai array dinamik, agar nantinya
'bisa dideklarasi ulang...
Dim arr4() As Byte

'Sama seperti Anda mengassign suatu variabel ke variabel lainnya,
'sebagai contoh: strA = strB, Anda juga dapat mengassign isi dari
'array ke array lainnya. Bayangkan, dalam hal ini, yang Anda
'inginkan adalah meng-copy sebuah array bytes dari yang satu ke
'yang lainnya. Anda dapat melakukannya dengan meng-copy satu per
'satu elemen, seperti prosedur berikut:

Sub ByteCopy1(oldCopy() As Byte, newCopy() As Byte)
   Dim i As Integer
   ReDim newCopy(LBound(oldCopy) To UBound(oldCopy))
   For i = LBound(oldCopy) To UBound(oldCopy)
      newCopy(i) = oldCopy(i)
   Next
End Sub

'Cara lainnya yang lebih efisien daripada cara di atas adalah
'menggunakan prosedur ByteCopy2 berikut ini, yaitu dengan
'cara mengassign array lama ke array baru.
Sub ByteCopy2(oldCopy() As Byte, newCopy() As Byte)
   newCopy = oldCopy
End Sub

'Command1 menggunakan prosedur ByteCopy1()
Private Sub Command1_Click()
    'Isi setiap elemen arr1
    arr1(0) = 0
    arr1(1) = 1
    arr1(2) = 2
    'Deklarasi ulang arr2 sebanyak 7 elemen
    ReDim arr2(6)
    'Isi elemen arr2 indeks ke-1 dengan 2
    arr2(1) = 2
    'Tampilkan setelah diisi...
    MsgBox "Elemen arr2 sebelum dicopy = " _
           & arr2(1), vbInformation  '--> Menghasilkan 2
    
    'Copy arr1 ke arr2. Perhatikan, pada mulanya,
    'arr2 memiliki 7 elemen (ReDim arr2(6))
    'lalu saat menjalankan ByteCopy1() di bawah ini,
    'elemen arr2() menjadi 3, karena dicopy dari arr1
    'sehingga seluruh elemen arr2() menjadi sama dengan
    'jumlah elemen arr1()
    Call ByteCopy1(arr1(), arr2())
    MsgBox "Elemen arr2 setelah dicopy " & vbCrLf & _
           "dari arr1 = " & arr2(1), _
           vbInformation '--> Menghasilkan 1 karena
                         '    sekarang arr2(1) = arr1(1)
End Sub

'Command2 menggunakan prosedur ByteCopy2()
Private Sub Command2_Click()
    'Isi setiap elemen arr3
    arr3(0) = 1
    arr3(1) = 2
    arr3(2) = 3
    'Deklarasi ulang arr4 sebanyak 11 elemen
    ReDim arr4(10)
    'Isi elemen arr2 indeks ke-2 dengan 5
    arr4(2) = 5
    'Tampilkan setelah diisi...
    MsgBox "Elemen arr4 sebelum dicopy = " _
           & arr4(2), vbInformation  '--> Menghasilkan 5
    
    'Copy arr3 ke arr4. Perhatikan, pada mulanya,
    'arr4 memiliki 11 elemen (ReDim arr4(10))
    'lalu setelah prosedur ByteCopy2() di bawah ini,
    'elemen arr4() menjadi 3, karena dicopy dari arr3
    'sehingga seluruh elemen arr4() menjadi sama dengan
    'jumlah elemen arr3()
    Call ByteCopy2(arr3(), arr4())
    'Setelah di-copy, tampilkan lagi...
    MsgBox "Elemen arr4 setelah dicopy " & vbCrLf & _
           "dari arr3 = " & arr4(2), _
           vbInformation  '--> Menghasilkan 3 karena
                          '    sekarang arr4(2) = arr3(2)
End Sub

