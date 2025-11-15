import streamlit as st

st.set_page_config(layout="wide")

if 'kondisi' not in st.session_state:
    st.session_state['kondisi']={'kondisi1':True,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False,
                                 'kondisi7':False, 'kondisi8':False,
                                 'kondisi9':False}

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
#==================================================

if st.sidebar.button('Beranda'):
    st.session_state['kondisi']={'kondisi1':True,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False}
    st.rerun()
    
if st.sidebar.button('pengantar'):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':True,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False}
    st.rerun()

if st.sidebar.button("Class VBA for Excel"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':True, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False}
    st.rerun()
st.sidebar.markdown("---")
if st.sidebar.button("Luas Persegi Panjang AI"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':True, 'kondisi9':False}
    st.rerun()
if st.sidebar.button("Deret Aritmatika"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':True}
    st.rerun()
st.sidebar.markdown("---")
if st.sidebar.button("Kalkulator Sederhana"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':True,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False}
    st.rerun()
if st.sidebar.button("Pecahan Sederhana"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':True, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False}
    st.rerun()
if st.sidebar.button("Dimensi 3"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':True,
                                 'kondisi8':False, 'kondisi9':False}
    st.rerun()
if st.sidebar.button("Lihat Media Hasil Diskusi"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':True,
                                 'kondisi5':False, 'kondisi6':False, 'kondisi7':False,
                                 'kondisi8':False, 'kondisi9':False}
    st.rerun()

st.subheader("Ruang Diskusi")
st.markdown('''
                <iframe src="https://martin123-oke.github.io/media/diskusi1.html" style="width:100%; height:3000px"></iframe>
            ''',unsafe_allow_html=True)
