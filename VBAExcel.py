import streamlit as st

st.set_page_config(layout="wide")

if 'kondisi' not in st.session_state:
    st.session_state['kondisi']={'kondisi1':True,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False}

def kover():
    st.markdown('''
                <iframe src="https://martin123-oke.github.io/media/kover.html" style="width:100%; height:1000px">
                <iframe>
                ''',unsafe_allow_html=True)import streamlit as st

st.set_page_config(layout="wide")

if 'kondisi' not in st.session_state:
    st.session_state['kondisi']={'kondisi1':True,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False}

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
    st.subheader("Ruang Diskusi")
    st.markdown('''
                    <iframe src="https://martin123-oke.github.io/media/diskusi1.html" style="width:100%; height:1000px"></iframe>
                    ''',unsafe_allow_html=True)
    

if st.session_state.kondisi['kondisi1']:
    kover()
if st.session_state.kondisi['kondisi2']:
    materi1()
if st.session_state.kondisi['kondisi3']:
    materi2()

if st.sidebar.button('Beranda'):
    st.session_state['kondisi']={'kondisi1':True,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False}
    st.rerun()
    
if st.sidebar.button('pengantar'):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':True,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False}
    st.rerun()
if st.sidebar.button("Kalkulator Sederhana"):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':False,
                                 'kondisi3':True,'kondisi4':False,
                                 'kondisi5':False}
    st.rerun()


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


if st.session_state.kondisi['kondisi1']:
    kover()
if st.session_state.kondisi['kondisi2']:
    materi1()

if st.sidebar.button('Beranda'):
    st.session_state['kondisi']={'kondisi1':True,'kondisi2':False,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False}
    st.rerun()
    
if st.sidebar.button('pengantar'):
    st.session_state['kondisi']={'kondisi1':False,'kondisi2':True,
                                 'kondisi3':False,'kondisi4':False,
                                 'kondisi5':False}
    st.rerun()

