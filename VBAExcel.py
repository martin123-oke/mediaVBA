import streamlit as st

st.set_page_config(layout="wide")

if 'kondisi' not in st.session_state:
    st.session_state['kondisi']={'kondisi1':True,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False,
                                 'kondisi7':False, 'kondisi8':False,
                                 'kondisi9':False, 'kondisi10': False,
                                 'kondisi11':False, 'kondisi12':False,
                                 'kondisi13':False,'kondisi14':False,
                                 'kondisi15':False, 'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False,
                                 'kondisi19':False, 'kondisi20':False,
                                 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}

def kover():
    st.markdown('''
                <iframe src="https://martin123-oke.github.io/media/kover.html" style="width:100%; height:1000px">
                <iframe>
                ''',unsafe_allow_html=True)

def materi1():
    kolom1 = st.tabs(['Test Diagnosis Koding VBA for Excel','Pengantar','Projek untuk Pengembangan'])
    with kolom1[0]:
        st.markdown('''
                <iframe src="https://martin123-oke.github.io/testDiagnosis1/test_daignosis.html" style="width:100%; height:2600px">
                <iframe>
                ''',unsafe_allow_html=True)
    with kolom1[1]:
        st.markdown('''
                <iframe src="https://martin123-oke.github.io/PengVBAExcel/pelajaran%20pertama.html" style="width:100%; height:2600px">
                <iframe>
                ''',unsafe_allow_html=True)
        st.code("""
         ' Fungsi untuk memeriksa jawaban siswa
         Sub CheckAnswer()
         Dim userAnswer As Integer
         Dim correctAnswer As Integer
         ' Ambil jawaban dari cell
         userAnswer = Range("B2").Value
         correctAnswer = 56
         ' Periksa jawaban
         If userAnswer = correctAnswer Then
             MsgBox "Benar! Hebat sekali!", vbInformation
             Range("C2").Value = "✓ Benar"
             Range("C2").Interior.Color = RGB(144, 238, 144)
             ' Excel "bicara"
             Application.Speech.Speak "Selamat, jawaban Anda benar!"
         Else
             MsgBox "Coba lagi! Jawaban belum tepat.", vbExclamation
             Range("C2").Value = "✗ Salah"
             Range("C2").Interior.Color = RGB(255, 182, 193)
         End If
         End Sub
         '------------------------------------------
         ' Fungsi untuk membuat soal random
         Sub GenerateQuestion()
         Dim num1, num2 As Integer
         num1 = Int(Rnd() * 10) + 1
         num2 = Int(Rnd() * 10) + 1
         Range("A2").Value = "Berapa " & num1 & " × " & num2 & "?"
         End Sub
         #----------------Tombol Acak-------------
         Sub acak()
            Range("C2").Value = Int(Rnd() * 20)
            Range("E2").Value = Int(Rnd() * 20)
            Range("G2").Value = ""
            Range("j2").Interior.Color = vbWhite
            Range("j2").Value = ""
        End Sub
        """)
    with kolom1[2]:
        st.markdown('<div>Pecahan</div>',unsafe_allow_html=True)
        st.markdown('''
                    <iframe src="https://martin123-oke.github.io/media/pecahan.html" style="width:100%; height:1000px"></iframe>
                    ''',unsafe_allow_html=True)
        st.markdown('<div>Konverter Bilangan</div>',unsafe_allow_html=True)
        st.markdown('''
                    <iframe src="https://martin123-oke.github.io/media/konverter.html" style="width:100%; height:1000px"></iframe>
                    ''',unsafe_allow_html=True)
        st.markdown('<div>Tabel Perkalian</div>',unsafe_allow_html=True)
        st.markdown('''
                    <iframe src="https://martin123-oke.github.io/media/perkalian.html" style="width:100%; height:1000px"></iframe>
                    ''',unsafe_allow_html=True)
        st.markdown('<div>Kalkulator Bidang Datar</div>',unsafe_allow_html=True)
        st.markdown('''
                    <iframe src="https://martin123-oke.github.io/media/KalkulatorBidang.html" style="width:100%; height:1000px"></iframe>
                    ''',unsafe_allow_html=True)

def materi2():
    st.header("Kalkulator Sederhana")
    st.markdown("<iframe src='https://res.cloudinary.com/dfkw4ux0e/image/upload/v1761436970/Kalkulator1_oenshy.png' style='width:100%;height:400px'></iframe>",unsafe_allow_html=True)
    st.code('''
'-----untuk tombol tambah------
Sub tambah()
bilangan1 = Range("C2").Value
bilangan2 = Range("E2").Value
Range("G2").Value = bilangan1 + bilangan2
Range("D2").Value = "+"
If Range("J2").Value = Range("G2") Then
    Range("J2").Interior.Color = vbGreen
Else
    Range("J2").Interior.Color = vbRed
End If
End Sub
'-----untuk tombol kurang------
Sub kurang()
bilangan1 = Range("C2").Value
bilangan2 = Range("E2").Value
Range("G2").Value = bilangan1 - bilangan2
Range("D2").Value = "-"
If Range("J2").Value = Range("G2") Then
    Range("J2").Interior.Color = vbGreen
Else
    Range("J2").Interior.Color = vbRed
End If
End Sub
'-----untuk tombol kali------
Sub kali()
bilangan1 = Range("C2").Value
bilangan2 = Range("E2").Value
Range("G2").Value = bilangan1 * bilangan2
Range("D2").Value = "x"
If Range("J2").Value = Range("G2") Then
    Range("J2").Interior.Color = vbGreen
Else
    Range("J2").Interior.Color = vbRed
End If
End Sub
''-----untuk tombol bagi------
Sub bagi()
bilangan1 = Range("C2").Value
bilangan2 = Range("E2").Value
Range("G2").Value = bilangan1 / bilangan2
Range("D2").Value = ":"
If Range("J2").Value = Range("G2") Then
    Range("J2").Interior.Color = vbGreen
Else
    Range("J2").Interior.Color = vbRed
End If
End Sub
''-----untuk tombol pangkat------
Sub pangkat()
bilangan1 = Range("C2").Value
bilangan2 = Range("E2").Value
Range("G2").Value = bilangan1 ^ bilangan2
Range("D2").Value = "^"
If Range("J2").Value = Range("G2") Then
    Range("J2").Interior.Color = vbGreen
Else
    Range("J2").Interior.Color = vbRed
End If
End Sub
Sub acak()
Range("C2").Value = Int(Rnd() * 20)
Range("E2").Value = Int(Rnd() * 20)
Range("G2").Value = ""
Range("j2").Interior.Color = vbWhite
Range("j2").Value = ""
End Sub
    ''')
     
def materi3():
    st.header("Hasil Diskusi")
    st.markdown("<iframe src='https://martin123-oke.github.io/PengVBAExcel/hasilDiskusi.html' style='width:100%;height:800px'></iframe>",unsafe_allow_html=True)
    

def materi4():
    st.header("Pecahan Sederhana")
    st.markdown("<iframe src='https://res.cloudinary.com/dfkw4ux0e/image/upload/v1762577355/Pecahan_p1zm2o.png' style='width:100%;height:300px'></iframe>",unsafe_allow_html=True)
    with st.expander("Konsep Dasar For"):
        st.write("Konsep Dasar")
        st.markdown('''
            For digunakan untuk melakukan perulangan (loop) dengan jumlah iterasi yang sudah ditentukan.
👉 Struktur dasarnya di VBA:
        ''')
        st.code('''
            For i = nilai_awal To nilai_akhir [Step nilai_langkah]
    'blok perintah yang diulang
Next i
        ''')
        st.markdown('''
- i → variabel penghitung (counter)
- To → batas akhir perulangan
- Step → besar kenaikan/penurunan tiap iterasi (default = 1)
''')
    with st.expander("Contoh Dasar"):
        st.write("Contoh Dasar")
        st.code('''
            Sub LoopAngka()
    Dim i As Integer
    For i = 1 To 5
        MsgBox "Angka ke-" & i
    Next i
End Sub
        ''')
        st.markdown('''
📘 Penjelasan:
- Program menampilkan pesan 5 kali: "Angka ke-1", "Angka ke-2", dst.
- For i = 1 To 5 → dimulai dari 1 hingga 5.
''')
    st.code('''
Sub pecahan()
On Error Resume Next
Dim lembar As Worksheet
Dim gambar As Shape

'mengaktifkan lembar
Set lembar = Worksheets("pecahan")

'memastikan posisi dan ukuran
posisix = lembar.Shapes("kotak").Left
posisiy = lembar.Shapes("kotak").Top
panjang = lembar.Shapes("kotak").Width
lebar = lembar.Shapes("kotak").Height

'hapus objek yang masih ada
For i = 1 To 200
    lembar.Shapes("kotak" & i).Delete
Next i

'menggambakan penyebut
For i = 1 To Range("B3").Value
    Set gambar = lembar.Shapes.AddShape(msoShapeRectangle, posisix, posisiy + (i - 1) * lebar / Range("B3").Value, panjang, lebar / Range("B3").Value)
    gambar.Name = "kotak" & i
    lembar.Shapes("kotak" & i).Fill.ForeColor.RGB = vbWhite
    lembar.Shapes("kotak" & i).Line.ForeColor.RGB = vbBlack
    lembar.Shapes("kotah" & i).Line.Weight = 2
Next i

'Asiran untuk pembilang
For i = 1 To Range("B2")
    lembar.Shapes("kotak" & Range("B3") - (i - 1)).Fill.Patterned msoPatternDarkUpwardDiagonal
    lembar.Shapes("kotak" & Range("B3") - (i - 1)).Fill.ForeColor.RGB = vbBlack
Next i
End Sub
    ''')

def materi5():
    st.markdown('<div>Class VBA for Excel</div>',unsafe_allow_html=True)
    st.markdown('''
                <iframe src="https://martin123-oke.github.io/PengVBAExcel/oopvba.html" style="width:100%; height:1000px"></iframe>
                ''',unsafe_allow_html=True)

def materi6():
    st.markdown('<div style="font-family:Arial; font-size:30px; font-weight:bold">Dimensi 2 ke Dimensi 3</div>',unsafe_allow_html=True)
    st.markdown('''
                    <iframe src="https://res.cloudinary.com/ikip-siliwangi/image/upload/v1763183847/dimensi3_jof8fy.png" style="width:100%; height:500px"></iframe>
                    ''',unsafe_allow_html=True)
    st.code('''
Sub sumbu()
On Error Resume Next
Dim lembar As Worksheet
Dim garis1 As Shape
Dim garis2 As Shape
Dim garis3 As Shape
Dim garis4 As Shape
Dim kot1 As Shape
Dim himpunan1(0 To 4, 0 To 1) As Single
Dim kot2 As Shape
Dim himpunan2(0 To 4, 0 To 1) As Single
Dim kot3 As Shape
Dim himpunan3(0 To 4, 0 To 1) As Single
Dim kot4 As Shape
Dim himpunan4(0 To 4, 0 To 1) As Single
Dim ling As Shape
Dim himpunan5(0 To 360, 0 To 1) As Single
Set lembar = Worksheets(1)
lembar.Shapes("koordinat1").Delete
lembar.Shapes("koordinat2").Delete
lembar.Shapes("koordinat3").Delete
lembar.Shapes("koordinat4").Delete
lembar.Shapes("kotak1").Delete
lembar.Shapes("kotak2").Delete
lembar.Shapes("kotak3").Delete
lembar.Shapes("kotak4").Delete
For p = 1 To 1000
    lembar.Shapes("lingkaran" & p).Delete
Next p
a = Range("B2")
b = Range("B3")
c = Range("B4")
Pi = 4 * Atn(1)
d = Range("B6")
e = lembar.Shapes("titik").Left + lembar.Shapes("titik").Width / 2
f = lembar.Shapes("titik").Top + lembar.Shapes("titik").Height / 2
g = Range("B9")
h = Range("B10")
i = Range("B11")
x1 = d * (-Sin(Pi * a / 180) * Sin(Pi * b / 180) * Cos(Pi * c / 180) + Sin(c * Pi / 180) * Cos(a * Pi / 180))
y1 = -d * (Sin(b * Pi / 180) * Cos(a * Pi / 180) * Cos(c * Pi / 180) + Sin(a * Pi / 180) * Sin(c * Pi / 180))
x2 = d * (Sin(a * Pi / 180) * Sin(b * Pi / 180) + Cos(a * Pi / 180) * Cos(c * Pi / 180))
y2 = -d * (-Sin(b * Pi / 180) * Sin(c * Pi / 180) * Cos(a * Pi / 180) + Sin(a * Pi / 180) * Cos(c * Pi / 180))
x3 = d * (-Sin(a * Pi / 180) * Cos(b * Pi / 180))
y3 = -d * Cos(a * Pi / 180) * Cos(b * Pi / 180)
Set garis1 = lembar.Shapes.AddLine(e, f, e + x1, f + y1)
garis1.Name = "koordinat1"
lembar.Shapes("koordinat1").Line.ForeColor.RGB = vbRed
lembar.Shapes("koordinat1").Line.BeginArrowheadStyle = msoArrowheadOval
lembar.Shapes("koordinat1").Line.EndArrowheadStyle = msoArrowheadTriangle
lembar.Shapes("koordinat1").Line.Weight = 2
Set garis2 = lembar.Shapes.AddLine(e, f, e + x2, f + y2)
garis2.Name = "koordinat2"
lembar.Shapes("koordinat2").Line.ForeColor.RGB = vbRed
lembar.Shapes("koordinat2").Line.BeginArrowheadStyle = msoArrowheadOval
lembar.Shapes("koordinat2").Line.EndArrowheadStyle = msoArrowheadTriangle
lembar.Shapes("koordinat2").Line.Weight = 2
Set garis3 = lembar.Shapes.AddLine(e, f, e + x3, f + y3)
garis3.Name = "koordinat3"
lembar.Shapes("koordinat3").Line.ForeColor.RGB = vbRed
lembar.Shapes("koordinat3").Line.BeginArrowheadStyle = msoArrowheadOval
lembar.Shapes("koordinat3").Line.EndArrowheadStyle = msoArrowheadTriangle
lembar.Shapes("koordinat3").Line.Weight = 2
Set garis4 = lembar.Shapes.AddLine(e + g * x1 / d + h * x2 / d, f + g * y1 / d + h * y2 / d, e + g * x1 / d + h * x2 / d + i * x3 / d, f + g * y1 / d + h * y2 / d + i * y3 / d)
garis4.Name = "koordinat4"
lembar.Shapes("koordinat4").Line.ForeColor.RGB = vbBlue
lembar.Shapes("koordinat4").Line.Weight = 2
himpunan1(0, 0) = e
himpunan1(0, 1) = f
himpunan1(1, 0) = e + g * x1 / d
himpunan1(1, 1) = f + g * y1 / d
himpunan1(2, 0) = e + g * x1 / d + h * x2 / d
himpunan1(2, 1) = f + g * y1 / d + h * y2 / d
himpunan1(3, 0) = e + h * x2 / d
himpunan1(3, 1) = f + h * y2 / d
himpunan1(4, 0) = e
himpunan1(4, 1) = f
Set kot1 = lembar.Shapes.AddPolyline(himpunan1)
kot1.Name = "kotak1"
lembar.Shapes("kotak1").Fill.Transparency = 1
lembar.Shapes("kotak1").Line.ForeColor.RGB = vbBlue
lembar.Shapes("kotak1").Line.Weight = 2
himpunan2(0, 0) = e
himpunan2(0, 1) = f
himpunan2(1, 0) = e + g * x1 / d
himpunan2(1, 1) = f + g * y1 / d
himpunan2(2, 0) = e + g * x1 / d + i * x3 / d
himpunan2(2, 1) = f + g * y1 / d + i * y3 / d
himpunan2(3, 0) = e + i * x3 / d
himpunan2(3, 1) = f + i * y3 / d
himpunan2(4, 0) = e
himpunan2(4, 1) = f
Set kot2 = lembar.Shapes.AddPolyline(himpunan2)
kot2.Name = "kotak2"
lembar.Shapes("kotak2").Fill.Transparency = 1
lembar.Shapes("kotak2").Line.ForeColor.RGB = vbBlue
lembar.Shapes("kotak2").Line.Weight = 2
himpunan3(0, 0) = e
himpunan3(0, 1) = f
himpunan3(1, 0) = e + h * x2 / d
himpunan3(1, 1) = f + h * y2 / d
himpunan3(2, 0) = e + h * x2 / d + i * x3 / d
himpunan3(2, 1) = f + h * y2 / d + i * y3 / d
himpunan3(3, 0) = e + i * x3 / d
himpunan3(3, 1) = f + i * y3 / d
himpunan3(4, 0) = e
himpunan3(4, 1) = f
Set kot3 = lembar.Shapes.AddPolyline(himpunan3)
kot3.Name = "kotak3"
lembar.Shapes("kotak3").Fill.Transparency = 1
lembar.Shapes("kotak3").Line.ForeColor.RGB = vbBlue
lembar.Shapes("kotak3").Line.Weight = 2
himpunan4(0, 0) = e + i * x3 / d
himpunan4(0, 1) = f + i * y3 / d
himpunan4(1, 0) = e + h * x2 / d + i * x3 / d
himpunan4(1, 1) = f + h * y2 / d + i * y3 / d
himpunan4(2, 0) = e + h * x2 / d + i * x3 / d + g * x1 / d
himpunan4(2, 1) = f + h * y2 / d + i * y3 / d + g * y1 / d
himpunan4(3, 0) = e + i * x3 / d + g * x1 / d
himpunan4(3, 1) = f + i * y3 / d + g * y1 / d
himpunan4(4, 0) = e + i * x3 / d
himpunan4(4, 1) = f + i * y3 / d
Set kot4 = lembar.Shapes.AddPolyline(himpunan4)
kot4.Name = "kotak4"
lembar.Shapes("kotak4").Fill.Transparency = 1
lembar.Shapes("kotak4").Line.ForeColor.RGB = vbBlue
lembar.Shapes("kotak4").Line.Weight = 2
Radius = lembar.Range("B14")
tinggi = lembar.Range("B15")
n = 0
For k = 0 To 2 * Radius
    m = Sqr((Radius) ^ 2 - (Radius - tinggi * k) ^ 2)
    tx = e + g * x1 / (2 * d) + h * x2 / (2 * d)
    ty = f + g * y1 / (2 * d) + h * y2 / (2 * d)
    For l = 0 To 360
        n = n + 1
        himpunan5(l, 0) = tx + m * ((x1 + x2) / d) * Cos(l * Pi / 180) + k * tinggi * x3 / d
        himpunan5(l, 1) = ty + m * ((y1 + y2) / d) * Sin(l * Pi / 180) + k * tinggi * y3 / d
    Next l
    Set ling = lembar.Shapes.AddPolyline(himpunan5)
    ling.Name = "lingkaran" & k
    lembar.Shapes("lingkaran" & k).Line.ForeColor.RGB = RGB(0, 125, 0)
    lembar.Shapes("lingkaran" & k).Fill.Transparency = 1
    lembar.Shapes("lingkaran" & k).Line.Weight = 2
Next k
End Sub

Sub koordinat_x1()
Range("B2") = Range("B2") + 10
Call sumbu
End Sub

Sub koordinat_x2()
Range("B2") = Range("B2") - 10
Call sumbu
End Sub

Sub koordinat_y1()
Range("B3") = Range("B3") + 10
Call sumbu
End Sub

Sub koordinat_y2()
Range("B3") = Range("B3") - 10
Call sumbu
End Sub

Sub koordinat_z1()
Range("B4") = Range("B4") + 10
Call sumbu
End Sub

Sub koordinat_z2()
Range("B4") = Range("B4") - 10
Call sumbu
End Sub

Sub panjang1()
Range("B6") = Range("B6") + 10
Call sumbu
End Sub
Sub panjang2()
Range("B6") = Range("B6") - 10
Call sumbu
End Sub
    ''')

def materi7():
    st.markdown('<div style="font-family:Arial; font-size:30px; font-weight:bold">Luas Persegi Panjang</div>',unsafe_allow_html=True)
    st.code('''
' Module: mLuasPersegiPanjang
Sub GambarPersegiPanjang_InputBox()
Dim p As Double, l As Double
Dim luas As Double
Dim shp As Shape
Dim ws As Worksheet
Dim widthPixels As Double, heightPixels As Double
Dim scaleFactor As Double
Set ws = ThisWorkbook.ActiveSheet
' 1. Input dari user
p = Val(InputBox("Masukkan panjang (contoh: 8):", "Input Panjang"))
l = Val(InputBox("Masukkan lebar (contoh: 5):", "Input Lebar"))
If p <= 0 Or l <= 0 Then
MsgBox "Panjang dan lebar harus positif.", vbExclamation
Exit Sub
End If
' 2. Hitung luas
luas = p * l
' 3. Hapus shape lama (jika ada dengan nama tertentu)
On Error Resume Next
ws.Shapes("RectLuas").Delete
On Error GoTo 0
' 4. Skala: map ukuran satuan (misal 1 unit = 20 pixel) — sesuaikan agar muat layar
scaleFactor = 20 ' ubah sesuai kebutuhan
widthPixels = p * scaleFactor
heightPixels = l * scaleFactor
' 5. Pastikan tidak terlalu besar untuk sheet: batasi maksimal
If widthPixels > 800 Then widthPixels = 800
If heightPixels > 500 Then heightPixels = 500
' 6. Tambah shape => left, top, width, height (pixel)
Set shp = ws.Shapes.AddShape(msoShapeRectangle, 50, 50, widthPixels, heightPixels)
shp.Name = "RectLuas"
shp.Fill.ForeColor.RGB = RGB(198, 239, 206) ' warna latar (opsional)
shp.Line.Weight = 2
shp.Line.ForeColor.RGB = RGB(0, 97, 0)
' 7. Tambah teks di dalam shape: p × l = luas
shp.TextFrame2.TextRange.Text = "p=" & p & " l=" & l & vbCrLf & "L=" & luas
shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
shp.TextFrame2.VerticalAnchor = msoAnchorMiddle
shp.TextFrame2.TextRange.Font.Size = 12
' 8. Tulis juga ke sheet (opsional)
ws.Range("B2").Value = "Panjang"
ws.Range("C2").Value = p
ws.Range("B3").Value = "Lebar"
ws.Range("C3").Value = l
ws.Range("B4").Value = "Luas"
ws.Range("C4").Value = luas
MsgBox "Gambar dan perhitungan selesai. Luas = " & luas, vbInformation
End Sub
    ''')
    
def materi8():
    st.markdown('<div style="font-family:Arial; font-size:30px; font-weight:bold">Deret Aritmatika</div>',unsafe_allow_html=True)
    st.code('''
Sub VisualisasiDeretAritmatika()
Dim a As Double, d As Double, n As Integer
Dim i As Integer
Dim nilai As Double
Dim tinggi As Double
Dim kiri As Double
Dim shp As Shape
'Input nilai
a = Val(InputBox("Masukkan suku pertama (a):", "Input A"))
d = Val(InputBox("Masukkan beda (d):", "Input d"))
n = Val(InputBox("Masukkan banyak suku (n):", "Input n"))
If n <= 0 Then
MsgBox "Jumlah suku harus lebih dari 0", vbExclamation
Exit Sub
End If
'Hapus shape lama
Dim s As Shape
For Each s In ActiveSheet.Shapes
If Left(s.Name, 5) = "Batang" Or Left(s.Name, 5) = "Label" Then
s.Delete
End If
Next s
kiri = 50 'posisi awal kiri
'Menggambar batang suku deret
For i = 1 To n
nilai = a + (i - 1) * d
tinggi = nilai * 10 'skala visual
'Gambar batang
Set shp = ActiveSheet.Shapes.AddShape(msoShapeRectangle, kiri, 300 - tinggi, 30, tinggi)
shp.Name = "Batang" & i
shp.Fill.ForeColor.RGB = RGB(135, 206, 250)
shp.Line.ForeColor.RGB = RGB(0, 0, 139)
'Label nilai
Set shp = ActiveSheet.Shapes.AddLabel(msoTextOrientationHorizontal, kiri, 305, 30, 20)
shp.TextFrame2.TextRange.Text = CStr(nilai)
shp.Name = "Label" & i
shp.TextFrame2.TextRange.Font.Size = 10
shp.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
kiri = kiri + 40 'geser posisi ke kanan
Next i
MsgBox "Visualisasi deret selesai!", vbInformation
End Sub
    ''')

def materi9():
    st.markdown('<div style="font-family:Arial; font-size:30px; font-weight:bold">Perbandingan Senilai dan Terbalik</div>',unsafe_allow_html=True)
    st.code('''
Sub BuatMediaPembelajaran()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim x As Integer, y As Integer
    
    ' Buat sheet baru
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Perbandingan_" & Format(Now, "hhmmss")
    
    ' Atur lebar kolom
    ws.Columns("A:K").ColumnWidth = 12
    ws.Rows("1:50").RowHeight = 20
    
    ' ===== JUDUL UTAMA =====
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 50, 10, 600, 50)
    With shp
        .Fill.ForeColor.RGB = RGB(0, 102, 204)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "MEDIA PEMBELAJARAN PERBANDINGAN"
        .TextFrame2.TextRange.Font.Size = 18
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    
    ' ===== BAGIAN 1: PERBANDINGAN SENILAI =====
    y = 80
    
    ' Judul Perbandingan Senilai
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 50, y, 300, 35)
    With shp
        .Fill.ForeColor.RGB = RGB(46, 204, 113)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "1. PERBANDINGAN SENILAI"
        .TextFrame2.TextRange.Font.Size = 14
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
    
    ' Penjelasan Perbandingan Senilai
    y = y + 45
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 50, y, 600, 80)
    With shp
        .Fill.ForeColor.RGB = RGB(232, 245, 233)
        .Line.Weight = 2
        .Line.ForeColor.RGB = RGB(46, 204, 113)
        .TextFrame2.TextRange.Text = "Definisi: Dua besaran dikatakan senilai jika perbandingannya tetap." & vbLf & vbLf & _
                                     "Rumus: A₁/B₁ = A₂/B₂  atau  A₁ × B₂ = A₂ × B₁" & vbLf & vbLf & _
                                     "Contoh: Semakin banyak barang, semakin mahal harganya"
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.ParagraphFormat.LeftIndent = 10
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
    
    ' Contoh Soal Perbandingan Senilai
    y = y + 90
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 50, y, 600, 100)
    With shp
        .Fill.ForeColor.RGB = RGB(255, 249, 196)
        .Line.Weight = 2
        .Line.ForeColor.RGB = RGB(255, 193, 7)
        .TextFrame2.TextRange.Text = "CONTOH SOAL:" & vbLf & _
                                     "Jika 3 kg apel harganya Rp 45.000, berapa harga 7 kg apel?" & vbLf & vbLf & _
                                     "Penyelesaian:" & vbLf & _
                                     "3 kg → Rp 45.000" & vbLf & _
                                     "7 kg → (7 × 45.000) / 3 = Rp 105.000"
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.ParagraphFormat.LeftIndent = 10
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
    
    ' Tombol Latihan Senilai
    y = y + 110
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 50, y, 200, 40)
    With shp
        .Name = "BtnLatihanSenilai"
        .Fill.ForeColor.RGB = RGB(46, 204, 113)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "LATIHAN SENILAI"
        .TextFrame2.TextRange.Font.Size = 12
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .OnAction = "LatihanSenilai"
    End With
    
    ' ===== BAGIAN 2: PERBANDINGAN TERBALIK =====
    y = y + 60
    
    ' Judul Perbandingan Terbalik
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 50, y, 300, 35)
    With shp
        .Fill.ForeColor.RGB = RGB(231, 76, 60)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "2. PERBANDINGAN TERBALIK"
        .TextFrame2.TextRange.Font.Size = 14
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
    
    ' Penjelasan Perbandingan Terbalik
    y = y + 45
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 50, y, 600, 80)
    With shp
        .Fill.ForeColor.RGB = RGB(255, 235, 238)
        .Line.Weight = 2
        .Line.ForeColor.RGB = RGB(231, 76, 60)
        .TextFrame2.TextRange.Text = "Definisi: Dua besaran dikatakan terbalik jika hasil kalinya tetap." & vbLf & vbLf & _
                                     "Rumus: A₁ × B₁ = A₂ × B₂" & vbLf & vbLf & _
                                     "Contoh: Semakin banyak pekerja, semakin cepat selesai"
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.ParagraphFormat.LeftIndent = 10
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
    
    ' Contoh Soal Perbandingan Terbalik
    y = y + 90
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 50, y, 600, 100)
    With shp
        .Fill.ForeColor.RGB = RGB(255, 249, 196)
        .Line.Weight = 2
        .Line.ForeColor.RGB = RGB(255, 193, 7)
        .TextFrame2.TextRange.Text = "CONTOH SOAL:" & vbLf & _
                                     "Jika 4 orang dapat menyelesaikan pekerjaan dalam 6 hari," & vbLf & _
                                     "berapa hari waktu yang dibutuhkan jika dikerjakan 8 orang?" & vbLf & vbLf & _
                                     "Penyelesaian: (4 × 6) / 8 = 3 hari"
        .TextFrame2.TextRange.Font.Size = 11
        .TextFrame2.TextRange.ParagraphFormat.LeftIndent = 10
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
    
    ' Tombol Latihan Terbalik
    y = y + 110
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, 50, y, 200, 40)
    With shp
        .Name = "BtnLatihanTerbalik"
        .Fill.ForeColor.RGB = RGB(231, 76, 60)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "LATIHAN TERBALIK"
        .TextFrame2.TextRange.Font.Size = 12
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .OnAction = "LatihanTerbalik"
    End With
    
    MsgBox "Media pembelajaran berhasil dibuat!" & vbLf & vbLf & _
           "Klik tombol LATIHAN untuk mengerjakan soal interaktif.", vbInformation, "Berhasil"
End Sub

' ===== LATIHAN PERBANDINGAN SENILAI =====
Sub LatihanSenilai()
    Dim nilai1 As Integer, nilai2 As Integer, nilai3 As Integer
    Dim jawaban As Variant, hasil As Double
    
    ' Generate soal random
    nilai1 = Application.WorksheetFunction.RandBetween(2, 5)
    nilai2 = Application.WorksheetFunction.RandBetween(20, 50) * 1000
    nilai3 = Application.WorksheetFunction.RandBetween(6, 10)
    
    hasil = (nilai3 * nilai2) / nilai1
    
    ' Tampilkan soal
    jawaban = InputBox("SOAL PERBANDINGAN SENILAI" & vbLf & vbLf & _
                      "Jika " & nilai1 & " kg beras harganya Rp " & Format(nilai2, "#,##0") & "," & vbLf & _
                      "berapa harga " & nilai3 & " kg beras?" & vbLf & vbLf & _
                      "Masukkan jawaban (dalam rupiah, angka saja):", _
                      "Latihan Perbandingan Senilai")
    
    ' Cek jawaban
    If jawaban = "" Then Exit Sub
    
    If Val(jawaban) = hasil Then
        MsgBox "BENAR! " & Chr(10004) & vbLf & vbLf & _
               "Jawaban Anda: Rp " & Format(Val(jawaban), "#,##0") & vbLf & _
               "Penjelasan: " & nilai3 & " × " & Format(nilai2, "#,##0") & " ÷ " & nilai1 & _
               " = Rp " & Format(hasil, "#,##0"), vbInformation, "Hasil"
    Else
        MsgBox "SALAH! " & Chr(10006) & vbLf & vbLf & _
               "Jawaban Anda: Rp " & Format(Val(jawaban), "#,##0") & vbLf & _
               "Jawaban yang benar: Rp " & Format(hasil, "#,##0") & vbLf & vbLf & _
               "Penjelasan: " & nilai3 & " × " & Format(nilai2, "#,##0") & " ÷ " & nilai1 & _
               " = Rp " & Format(hasil, "#,##0"), vbExclamation, "Hasil"
    End If
    
    ' Tanya mau latihan lagi
    If MsgBox("Mau latihan soal lagi?", vbYesNo + vbQuestion, "Latihan Lagi?") = vbYes Then
        LatihanSenilai
    End If
End Sub

' ===== LATIHAN PERBANDINGAN TERBALIK =====
Sub LatihanTerbalik()
    Dim pekerja1 As Integer, hari1 As Integer, pekerja2 As Integer
    Dim jawaban As Variant, hasil As Double
    
    ' Generate soal random
    pekerja1 = Application.WorksheetFunction.RandBetween(3, 6)
    hari1 = Application.WorksheetFunction.RandBetween(8, 15)
    pekerja2 = Application.WorksheetFunction.RandBetween(8, 12)
    
    hasil = (pekerja1 * hari1) / pekerja2
    
    ' Tampilkan soal
    jawaban = InputBox("SOAL PERBANDINGAN TERBALIK" & vbLf & vbLf & _
                      "Jika " & pekerja1 & " orang dapat menyelesaikan pekerjaan dalam " & hari1 & " hari," & vbLf & _
                      "berapa hari waktu yang dibutuhkan jika dikerjakan " & pekerja2 & " orang?" & vbLf & vbLf & _
                      "Masukkan jawaban (dalam hari):", _
                      "Latihan Perbandingan Terbalik")
    
    ' Cek jawaban
    If jawaban = "" Then Exit Sub
    
    If Val(jawaban) = hasil Then
        MsgBox "BENAR! " & Chr(10004) & vbLf & vbLf & _
               "Jawaban Anda: " & Val(jawaban) & " hari" & vbLf & _
               "Penjelasan: (" & pekerja1 & " × " & hari1 & ") ÷ " & pekerja2 & _
               " = " & hasil & " hari", vbInformation, "Hasil"
    Else
        MsgBox "SALAH! " & Chr(10006) & vbLf & vbLf & _
               "Jawaban Anda: " & Val(jawaban) & " hari" & vbLf & _
               "Jawaban yang benar: " & hasil & " hari" & vbLf & vbLf & _
               "Penjelasan: (" & pekerja1 & " × " & hari1 & ") ÷ " & pekerja2 & _
               " = " & hasil & " hari", vbExclamation, "Hasil"
    End If
    
    ' Tanya mau latihan lagi
    If MsgBox("Mau latihan soal lagi?", vbYesNo + vbQuestion, "Latihan Lagi?") = vbYes Then
        LatihanTerbalik
    End If
End Sub
    ''')
def materi10():
    st.markdown('''<div style="font-size:30px;font-family:'bauhaus 93';font-weight:bold;text-align:center">Suara ketik tombol di tekan</div>''',unsafe_allow_html=True)
    st.markdown('''
                <iframe src="https://res.cloudinary.com/ikip-siliwangi/image/upload/v1763375805/tombol_suara_ysju4x.png" style="width:100%; height:400px"></iframe>
                ''',unsafe_allow_html=True)
    st.code("""
' Memanggil fungsi PlaySound dari Windows API
Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" ( _
    ByVal lpszName As String, _
    ByVal hModule As LongPtr, _
    ByVal dwFlags As Long) As Long
Public Sub Tombol_Bunyi()
    Dim fileSuara As String
    
    ' Lokasi file suara
    fileSuara = "C:\\Users\\IKIP SILIWANGI\\Desktop\\Media\\Bananas\\sounds\\Bite.wav"
    
    ' Play suara
    PlaySound fileSuara, 0, &H1
End Sub
    
    """)
    st.markdown('''<div style="font-size:20px;font-family:'broadway';font-weight:bold">Hentikan Suara</div>''',unsafe_allow_html=True)
    st.code('''
PlaySound vbNullString, 0, 0
PlaySound fileSuara, 0, &H1
    ''')
    st.markdown('''<div style="font-size:20px;font-family:'broadway';font-weight:bold">Membuat Banyak Tombol</div>''',unsafe_allow_html=True)
    st.code('''
Public Sub Bunyi_A()
    PlaySound "C:\suara\a.wav", 0, &H1
End Sub

Public Sub Bunyi_B()
    PlaySound "C:\suara\b.wav", 0, &H1
End Sub

''')

def materi11():
    kolom = st.tabs(['Koding Animasi','Bola Pantul','Bola Lompat','Solusi yang paling stabil','Teknik Animasi', 'Teknik Animasi Class'])
    with kolom[0]:
        st.markdown('''<div style="font-size:30px;font-family:'bauhaus 93';font-weight:bold;text-align:center">Konsep Animasi</div>''',unsafe_allow_html=True)
        st.markdown('''
                <iframe src="https://res.cloudinary.com/ikip-siliwangi/image/upload/v1763426601/animasi_atxnfw.png" style="width:100%; height:400px"></iframe>
                ''',unsafe_allow_html=True)
        st.code("""
Dim GerakAktif As Boolean
Dim Arah As Integer

Sub MulaiAnimasi()
    GerakAktif = True
    Arah = 1      ' 1 = kanan, -1 = kiri
    Call Animasi
End Sub

Sub StopAnimasi()
    GerakAktif = False
End Sub

Sub Animasi()
    On Error Resume Next
    Dim shp As Shape
    Set shp = Sheets("Sheet1").Shapes("Bola")

    Do While GerakAktif
        shp.Left = shp.Left + (10 * Arah)

        ' Batas kanan
        If shp.Left > 400 Then
            Arah = -1
        End If

        ' Batas kiri
        If shp.Left < 20 Then
            Arah = 1
        End If

        DoEvents
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop
End Sub
""")
    with kolom[1]:
        st.markdown('''<div style="font-size:30px;font-family:'bauhaus 93';font-weight:bold;text-align:center">Bola Pantul</div>''',unsafe_allow_html=True)
        st.markdown('''
                <iframe src="https://res.cloudinary.com/ikip-siliwangi/image/upload/v1763428400/pantulan_v9m7q7.png" style="width:100%; height:400px"></iframe>
                ''',unsafe_allow_html=True)
        st.code("""
Dim aktifDiagonal As Boolean
Dim dx As Single, dy As Single

Sub MulaiDiagonal()
    aktifDiagonal = True
    dx = 0.2   ' Gerak horizontal
    dy = 0.2   ' Gerak vertikal
    Call AnimasiDiagonal
End Sub

Sub StopDiagonal()
    aktifDiagonal = False
End Sub

Sub AnimasiDiagonal()
    On Error Resume Next
    Dim shp As Shape
    Set shp = Sheets("Sheet1").Shapes("Bola")

    Do While aktifDiagonal
        shp.Left = shp.Left + dx
        shp.Top = shp.Top + dy

        ' Batas kanan/kiri
        If shp.Left < 10 Or shp.Left > 400 Then dx = -dx

        ' Batas atas/bawah
        If shp.Top < 10 Or shp.Top > 250 Then dy = -dy

        DoEvents
        Application.Wait Now + TimeValue("0:00:00.05")
    Loop
End Sub
""")
    with kolom[2]:
        st.markdown('''<div style="font-size:30px;font-family:'bauhaus 93';font-weight:bold;text-align:center">Bola Lompat</div>''',unsafe_allow_html=True)
        st.markdown('''
                <iframe src="https://res.cloudinary.com/ikip-siliwangi/image/upload/v1763430281/lompatan_cfpqza.png" style="width:100%; height:400px"></iframe>
                ''',unsafe_allow_html=True)
        st.code("""
Dim aktifLoncat As Boolean
Dim kecepatan As Double
Dim lastFrame As Double

Sub MulaiLoncat()
    aktifLoncat = True
    kecepatan = -12          ' awal lompatan ke atas
    lastFrame = Timer
    Call AnimasiLoncat
End Sub

Sub StopLoncat()
    aktifLoncat = False
End Sub

Sub AnimasiLoncat()
    Dim shp As Shape
    Dim gravitasi As Double: gravitasi = 0.6     ' gravitasi lebih kecil = halus
    Set shp = Sheets("Sheet1").Shapes("Bola")

    Do While aktifLoncat
        
        ' FRAME RATE ~10 ms (0.01 detik)
        If Timer - lastFrame >= 0.01 Then
        
            ' update posisi vertikal
            shp.Top = shp.Top + kecepatan
            kecepatan = kecepatan + gravitasi

            ' Lantai
            If shp.Top > 250 Then
                shp.Top = 250
                kecepatan = -Abs(kecepatan) * 0.85   ' mantul sedikit
            End If

            lastFrame = Timer
        End If
        
        DoEvents
    Loop
End Sub
""")
    with kolom[3]:
        st.markdown('''<div style="font-size:30px;font-family:'bauhaus 93';font-weight:bold;text-align:center">Catatan</div>''',unsafe_allow_html=True)
        st.subheader("Gunakan DoEvents + Loop tanpa Wait")
        st.write("Ini cara yang paling sering digunakan di animasi VBA.")
        st.write("Contoh pengganti TimeValue:")
        st.code('''
Dim t As Double
t = Timer

Do While aktif
    If Timer - t > 0.01 Then          ' 0.01 detik = 10 milidetik
        shp.Left = shp.Left + 5
        t = Timer
    End If

    DoEvents
Loop

''')
        st.markdown('''
Keunggulan:
- lebih halus,
- lebih cepat,
- tergantung CPU, bukan Excel Wait,
- delay bisa diatur hingga 1 milidetik.
''')
        st.subheader("Contoh Animasi Gerak Cepat Tanpa Wait")
        st.write("Ini versi yang geraknya halus 100%:")
        st.code("""
Sub AnimasiHalus()
    aktif = True

    Dim shp As Shape
    Set shp = Sheets("Sheet1").Shapes("Bola")

    Do While aktif
        shp.Left = shp.Left + 2
        
        If shp.Left > 400 Then shp.Left = 10
        
        DoEvents
    Loop
End Sub

""")
    with kolom[4]:
        st.markdown('''<div style="font-size:30px;font-family:'bauhaus 93';font-weight:bold;text-align:center">Teknik Animasi</div>''',unsafe_allow_html=True)
        st.subheader("Deklarasi API (WAJIB)")
        st.write("Masukkan ini di bagian paling atas Module:")
        st.code('''
' --- API Windows untuk Timer Presisi ---
Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
''')
        st.write("Currency digunakan karena aman 64 bit dan akurat sampai mikrodetik.")
        st.subheader("Variabel Global")
        st.code('''
Dim aktifGame As Boolean
Dim frameRate As Double
Dim freq As Currency
Dim lastTime As Currency

''')
        st.subheader("Setup Animasi (Mengambil Frekuensi Mikrodetik)")
        st.code('''
Sub MulaiGame()
    aktifGame = True

    ' ambil frekuensi hardware timer
    QueryPerformanceFrequency freq
    QueryPerformanceCounter lastTime

    frameRate = 1 / 120     ' target 120 FPS (super halus)

    Call AnimasiGame
End Sub

Sub StopGame()
    aktifGame = False
End Sub
''')
        st.markdown('''
Anda bisa mengubah target FPS:
- 1/60 → 60 FPS,
- 1/120 → 120 FPS (sangat halus),
- 1/240 → 240 FPS (super smooth)
''')
        st.subheader("Animasi Per Frame (Teknik Game Engine)")
        st.write("Contoh ini: shape bergerak diagonal super halus seperti bola melayang.")
        st.code("""
Sub AnimasiGame()
    Dim shp As Shape
    Set shp = Sheets("Sheet1").Shapes("Bola")

    Dim current As Currency
    Dim delta As Double
    Dim xSpeed As Double: xSpeed = 200     ' pixel per detik
    Dim ySpeed As Double: ySpeed = 140     ' pixel per detik

    QueryPerformanceCounter lastTime

    Do While aktifGame

        ' waktu frame sekarang
        QueryPerformanceCounter current

        ' konversi ke detik yang presisi
        delta = (current - lastTime) / freq
       
        If delta >= frameRate Then
            
            ' --- Update physics ---
            shp.Left = shp.Left + (xSpeed * delta)
            shp.Top = shp.Top + (ySpeed * delta)
            
            ' --- Pantulan batas ---
            If shp.Left < 10 Or shp.Left > 400 Then xSpeed = -xSpeed
            If shp.Top < 10 Or shp.Top > 250 Then ySpeed = -ySpeed

            ' simpan waktu frame terakhir
            lastTime = current
        End If
        
        DoEvents
    Loop
End Sub
""")
        st.subheader("Kenapa Versi Ini Paling Halus?")
        st.markdown("""
<div>
<pre>
✔ Menggunakan timer hardware CPU

Akurasi sampai 0.000001 detik (microsecond).

✔ Bisa menentukan target FPS secara presisi

Anda bisa buat:

30 FPS

60 FPS

120 FPS

500 FPS

✔ Animasi berbasis delta time seperti Unity / Unreal Engine

Kecil-besar FPS tidak mempengaruhi kecepatan animasi.

✔ Sangat stabil (tidak patah-patah)

Tidak seperti:

Wait

Sleep

Timer
</pre>
</div>
""", unsafe_allow_html=True)
    with kolom[5]:
        st.markdown('''<div style="font-size:30px;font-family:'bauhaus 93';font-weight:bold;text-align:center">Teknik Animasi Multitimer</div>''',unsafe_allow_html=True)
        st.subheader("1. Deklarasi API (WAJIB, paling atas Module)")
        st.code('''
' ==== API TIMER PRESISI TERTINGGI ====
Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
''')
        st.subheader("2. Variabel Global Multi-Timer")
        st.code('''
Dim aktifGame As Boolean
Dim freq As Currency

Dim timers As New Collection ' berisi semua animasi paralel
''')
        st.subheader("3. Template Class Module: cTimer")
        st.code('''
' ======== CLASS: cTimer ========
Public shp As Shape
Public velX As Double
Public velY As Double
Public lastTime As Currency

Public Sub Update(dt As Double)
    shp.Left = shp.Left + velX * dt
    shp.Top = shp.Top + velY * dt
    
    ' pantulan
    If shp.Left < 10 Or shp.Left > 400 Then velX = -velX
    If shp.Top < 10 Or shp.Top > 260 Then velY = -velY
End Sub
''')
        st.write("Class ini berfungsi sebagai mini-game object dengan kecepatan sendiri.")
        st.subheader("4. Membuat Banyak Shape Dengan Timer Sendiri")
        st.code('''
Sub TambahShape(x As Double, y As Double, vx As Double, vy As Double)
    Dim t As cTimer
    Set t = New cTimer

    Set t.shp = Sheets("Sheet1").Shapes.AddShape(msoShapeOval, x, y, 30, 30)
    t.shp.Fill.ForeColor.RGB = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))

    t.velX = vx
    t.velY = vy

    QueryPerformanceCounter t.lastTime

    timers.Add t
End Sub
''')
        st.subheader("5. Setup Multi-Object")
        st.code('''
Sub MulaiGame()
    Dim i As Long
    aktifGame = True
    
    QueryPerformanceFrequency freq
    
    Randomize
    
    timers = New Collection
    
    ' buat 8 bola animasi paralel
    For i = 1 To 8
        Call TambahShape(50 * i, 50, (Rnd * 240) - 120, (Rnd * 240) - 120)
    Next i
    
    Call EngineLoop
End Sub
''')
        st.subheader("Engine Utama (Multi Timer Loop) → PALING HALUS")
        st.code('''
Sub EngineLoop()
    Dim current As Currency
    Dim dt As Double
    Dim t As cTimer

    Do While aktifGame
        
        For Each t In timers
            QueryPerformanceCounter current
            
            dt = (current - t.lastTime) / freq   ' waktu per objek
            
            ' update objek
            Call t.Update(dt)
            
            t.lastTime = current
        Next t
        
        DoEvents
    Loop
End Sub
''')
        st.subheader("7. Stop Game")
        st.code('''
Sub StopGame()
    aktifGame = False
End Sub
''')
        data={'Fitur':['Per-shape timer', 'DeltaTime realtime',
                       'QueryPerformanceCounter','Tidak ada jeda manual (Wait)',
                       'Bisa ratusan objek'],
              'Keterangan':['Setiap objek punya waktu sendiri → super smooth',
                            'Tidak tergantung FPS → gerak selalu stabil',
                            'Presisi mikrodetik, lebih baik dari Application.OnTime dan Wait',
                            'Engine terus mengalir seperti game engine',
                            'Karena tidak delay blocking']}
        st.table(data)
        
def materi12():
    st.subheader("Pertemuan 2 mobil")
    st.markdown("<iframe src='https://res.cloudinary.com/ikip-siliwangi/image/upload/v1763949131/pertemuanMobil_moitbn.png' style='width:100%;height:500px'></iframe>",unsafe_allow_html=True)
    st.code('''
Sub masukan()
On Error Resume Next
Dim lembar As Worksheet

Set lembar = Worksheets(1)
a = InputBox("Masukan Jarak")
lembar.Shapes("jarak").TextFrame2.TextRange.Text = "Jarak: " & a
End Sub
Sub masukan1()
On Error Resume Next
Dim lembar As Worksheet

Set lembar = Worksheets(1)
a = InputBox("Masukan Kecepatan Mobil 1")
lembar.Shapes("kecepatan1").TextFrame2.TextRange.Text = "Masukan Kecepatan Mobil 1: " & a
End Sub
Sub masukan2()
On Error Resume Next
Dim lembar As Worksheet

Set lembar = Worksheets(1)
a = InputBox("Masukan Kecepatan Mobil 1")
lembar.Shapes("kecepatan2").TextFrame2.TextRange.Text = "Masukan Kecepatan Mobil 2: " & a
End Sub
Sub jalankan()
On Error Resume Next
Dim lembar As Worksheet
Dim waktu As Double

Set lembar = Worksheets(1)
posisix1 = lembar.Shapes("mobil1").Left + lembar.Shapes("mobil1").Width
posisiy1 = lembar.Shapes("mobil1").Top
posisix2 = lembar.Shapes("mobil2").Left
posisiy2 = lembar.Shapes("mobil2").Top

jarak1 = Int(lembar.Shapes("pemandangan").Width) - Int(lembar.Shapes("mobil1").Width)
waktu = Timer
aktif = True
wak = 0
 MsgBox jarak1
jarak = Split(lembar.Shapes("jarak").TextFrame2.TextRange.Text, ":")(1)
v1 = Split(lembar.Shapes("kecepatan1").TextFrame2.TextRange.Text, ":")(1)
v2 = Split(lembar.Shapes("kecepatan2").TextFrame2.TextRange.Text, ":")(1)
waktu1 = Int(jarak) / (Int(v1) + Int(v2))

Do While aktif
    If Timer - t > 0.01 Then         ' 0.01 detik = 10 milidetik
        wak = wak + 0.1
        lembar.Shapes("mobil1").Left = Int(v1) * wak * jarak1 / jarak
        lembar.Shapes("mobil2").Left = jarak1 - Int(v2) * wak * jarak1 / jarak
        If wak > waktu1 Then
            aktif = False
        End If
        lembar.Shapes("waktu").TextFrame2.TextRange.Text = "waktu: " & wak
        t = Timer
    End If
    DoEvents
Loop
End Sub
Sub awal_lagi()
On Error Resume Next
Dim lembar As Worksheet

Set lembar = Worksheets(1)
lembar.Shapes("mobil1").Left = 0
lembar.Shapes("mobil2").Left = lembar.Shapes("pemandangan").Left + lembar.Shapes("pemandangan").Width - lembar.Shapes("mobil2").Width
lembar.Shapes("waktu").TextFrame2.TextRange.Text = "waktu: 0"
End Sub
    ''')
def materi13():
    st.markdown('''
                <iframe src="https://martin-bernard26.github.io/projek1/kursus.html" style="width:100%; height:4000px"></iframe>
            ''',unsafe_allow_html=True)
def materi14():
    st.markdown('''
                <iframe src="https://martin-bernard26.github.io/projek1/soalUTS.html" style="width:100%; height:4000px"></iframe>
            ''',unsafe_allow_html=True)
def materi15():
    st.markdown('''
                <iframe src="https://martin123-oke.github.io/PengVBAExcel/filsafatVBA.html" style="width:100%; height:2500px"></iframe>
            ''',unsafe_allow_html=True)
def materi16():
    st.markdown('''
                <iframe src="https://res.cloudinary.com/dfkw4ux0e/image/upload/v1764551177/subtitusi_w8usm7.png" style="width:100%; height:700px"></iframe>
            ''',unsafe_allow_html=True)
def materi17():
    st.markdown('''
                <iframe src="https://martin123-oke.github.io/PengVBAExcel/videoUTS.html" style="width:100%; height:700px"></iframe>
            ''',unsafe_allow_html=True)
def materi18():
    st.markdown('''
                <iframe src="https://martin123-oke.github.io/PengVBAExcel/ReGexVBA.html" style="width:100%; height:700px"></iframe>
            ''',unsafe_allow_html=True)
def materi19():
    kolom = st.tabs(['Konsep List','konsep Dict'])
    with kolom[0]:
        st.markdown('''
                <iframe src="https://martin123-oke.github.io/PengVBAExcel/listVBA.html" style="width:100%; height:700px"></iframe>
            ''',unsafe_allow_html=True)
    with kolom[1]:
        st.markdown('''
                <iframe src="https://martin123-oke.github.io/PengVBAExcel/dictVBA.html" style="width:100%; height:700px"></iframe>
            ''',unsafe_allow_html=True)
def materi20():
    st.markdown('''
                <iframe src="https://drive.google.com/file/d/1SykVEtNOOvcxsuPtX_zxWwQxTTW2fnPg/preview" style="width:100%; height:700px"></iframe>
            ''',unsafe_allow_html=True)
def materi21():
    st.subheader("Masukan Hasil UAS")
    st.markdown("""<ol>
    <li>Tekan Bendera Hijau </li>
    <li>Tekan Ikan Buntel</li>
    <li>Isi Nama Anda </li>
    <li>Isi NIM Anda </li>
    <li>Isikan Alamat Link File Excel Anda </li>
    <li>Isikan Alamat Link Video Anda</li>
    <li>Muncul keterangan bahwa Anda sudah mengirimkan Tugas UAS </li>
    </ol>""",unsafe_allow_html=True)
    st.markdown('''
                <iframe src="https://martin-bernard26.github.io/uasVBAExcel/MasukanUAS.html" style="width:100%; height:700px"></iframe>
            ''',unsafe_allow_html=True)
def materi22():
    st.markdown('<h1 style="text-align:center">Soal UAS<h1>',unsafe_allow_html=True)
    st.markdown('''
                <iframe src="https://drive.google.com/file/d/1xq_GugUxGHxoB99JaTn-BIAKYP37WWn2/preview" style="width:100%; height:700px"></iframe>
            ''',unsafe_allow_html=True)
#==================================================

if st.session_state.kondisi['kondisi1']:
    kover()
if st.session_state.kondisi['kondisi2']:
    materi1()
if st.session_state.kondisi['kondisi3']:
    materi2()
if st.session_state.kondisi['kondisi4']:
    materi3()
if st.session_state.kondisi['kondisi5']:
    materi4()
if st.session_state.kondisi['kondisi6']:
    materi5()
if st.session_state.kondisi['kondisi7']:
    materi6()
if st.session_state.kondisi['kondisi8']:
    materi7()
if st.session_state.kondisi['kondisi9']:
    materi8()
if st.session_state.kondisi['kondisi10']:
    materi9()
if st.session_state.kondisi['kondisi11']:
    materi10()
if st.session_state.kondisi['kondisi12']:
    materi11()
if st.session_state.kondisi['kondisi13']:
    materi12()
if st.session_state.kondisi['kondisi14']:
    materi13()
if st.session_state.kondisi['kondisi15']:
    materi14()
if st.session_state.kondisi['kondisi16']:
    materi15()
if st.session_state.kondisi['kondisi17']:
    materi16()
if st.session_state.kondisi['kondisi18']:
    materi17()
if st.session_state.kondisi['kondisi19']:
    materi18()
if st.session_state.kondisi['kondisi20']:
    materi19()
if st.session_state.kondisi['kondisi21']:
    materi20()
if st.session_state.kondisi['kondisi22']:
    materi21()
if st.session_state.kondisi['kondisi23']:
    materi22()

#==================================================

st.sidebar.markdown("Teori dan Aplikasi")
if st.sidebar.button('Beranda'):
    st.session_state['kondisi']={'kondisi1':True,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':False, 'kondisi15':False, 'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False, 'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
    
if st.sidebar.button('pengantar'):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':True,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False, 'kondisi13':False,
                                 'kondisi14':False, 'kondisi15':False,'kondisi16':False,
                                 'kondisi17':False,'kondisi18':False, 'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
if st.sidebar.button('Filsafat Pembelajaran Koding'):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False, 'kondisi13':False,
                                 'kondisi14':False, 'kondisi15':False,'kondisi16':True,
                                 'kondisi17':False, 'kondisi18':False, 'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()

if st.sidebar.button("Class VBA for Excel"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':True, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':False,'kondisi15':False,'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False, 'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
if st.sidebar.button("Suara Tombol VBA for Excel"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':True, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':False,'kondisi15':False,'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False, 'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
if st.sidebar.button("Konsep Animasi VBA for Excel"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':True,'kondisi13':False,
                                 'kondisi14':False,'kondisi15':False,'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False, 'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
if st.sidebar.button("Konsep Subtitusi"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':False,'kondisi15':False,'kondisi16':False,
                                 'kondisi17':True, 'kondisi18':False, 'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
if st.sidebar.button("ReGex VBA for Excel"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':False,'kondisi15':False,'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False, 'kondisi19':True,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
if st.sidebar.button("List dan Dict VBA for Excel"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':False,'kondisi15':False,'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False, 'kondisi19':False,
                                 'kondisi20':True, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
if st.sidebar.button("Konsep Chart"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':False,'kondisi15':False,'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False, 'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':True, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
st.sidebar.markdown("---")
st.sidebar.markdown("Kumpulan media dari generatif AI")
if st.sidebar.button("Luas Persegi Panjang AI"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':True, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':False,'kondisi15':False,'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False, 'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
if st.sidebar.button("Deret Aritmatika"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':True, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':False,'kondisi15':False,'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False,'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
if st.sidebar.button("Perbandingan Senilai dan Terbalik"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':True,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':False,'kondisi15':False,'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False,'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
st.sidebar.markdown("---")
if st.sidebar.button("Kalkulator Sederhana"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':True,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':False,'kondisi15':False,'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False,'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
if st.sidebar.button("Pecahan Sederhana"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':True, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':False,'kondisi15':False,'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False,'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
if st.sidebar.button("Dimensi 3"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':True,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':False,'kondisi15':False,'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False,'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
if st.sidebar.button("Pertemuan 2 Mobil"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':True,
                                 'kondisi14':False,'kondisi15':False,'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False,'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
if st.sidebar.button("Lihat Media Hasil Diskusi"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':True,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':False,'kondisi15':False, 'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False,'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
st.sidebar.markdown("---")
if st.sidebar.button("Kursus Media Pembelajaran berbasis Digital"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':True,'kondisi15':False,'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False,'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
if st.sidebar.button("Soal UTS"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':False,'kondisi15':True,'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False,'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
if st.sidebar.button("Pengumpulan UTS"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':False,'kondisi15':False,'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':True,'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':False}
    st.rerun()
if st.sidebar.button("Soal UAS"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':False,'kondisi15':False,'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False,'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':False,
                                 'kondisi23':True}
    st.rerun()
if st.sidebar.button("Pengumpulan UAS"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False, 'kondisi10':False,
                                 'kondisi11':False, 'kondisi12':False,'kondisi13':False,
                                 'kondisi14':False,'kondisi15':False,'kondisi16':False,
                                 'kondisi17':False, 'kondisi18':False,'kondisi19':False,
                                 'kondisi20':False, 'kondisi21':False, 'kondisi22':True,
                                 'kondisi23':False}
    st.rerun()
st.sidebar.markdown("---")
st.subheader("Ruang Diskusi")
st.markdown('''
                <iframe src="https://martin123-oke.github.io/media/diskusi1.html" style="width:100%; height:4500px"></iframe>
            ''',unsafe_allow_html=True)
