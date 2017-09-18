# 
# * Utworzone przez SharpDevelop.
# * Użytkownik: m_witowski
# * Data: 2014-04-15
# * Godzina: 09:47
# *
# * Do zmiany tego szablonu użyj Narzędzia | Opcje | Kodowanie | Edycja Nagłówków Standardowych.
# 
from System import *
from System.Collections.Generic import *
from System.Drawing import *
from System.Windows.Forms import *
from Excel import *

class MainForm(Form):
	""" <summary>
	 Description of MainForm.
	 </summary>
	"""
	def __init__(self):
		#
		# The InitializeComponent() call is required for Windows Forms designer support.
		#
		self.InitializeComponent()

	#
	# TODO: Add constructor code after the InitializeComponent() call.
	#
	def OtwórzToolStripMenuItemClick(self, sender, e):
		result = openFileDialog1.ShowDialog()
		if result == DialogResult.OK:
			self._nazwa_pliku = openFileDialog1.FileName
			self._sciezka = openFileDialog1.InitialDirectory
			self.Read()
		else:
			return 

	def Read(self):
		path = self._sciezka + self._nazwa_pliku
		self._linnia = System.IO.File.ReadAllLines(path)
		textBox1.Lines = self._linnia
		# foreach (string  line in linnia ) {
		# 
		# textBox1.Text += Environment.NewLine  + line;
		# }
		#textBox1.Text = string.Join("_", linnia);
		self._dane_wczytane = True

	def Button1Click(self, sender, e):
		if self._dane_wczytane == True:
			backgroundWorker1.RunWorkerAsync()
			label1.Text = ""
			progressBar1.Value = 0
		else:
			label1.Text = "Plik txt nie wczytany"

	def BackgroundWorker1DoWork(self, sender, e):
		postep = 0
		mojeexcel = Excel.Application()
		mojeexcel.Visible = False
		pe = mojeexcel.Workbooks.Open("D:\\TRX_formularze.xlsx")
		ae = pe.Sheets[1]
		ae1 = pe.Sheets[2]
		ae2 = pe.Sheets[3]
		#zapis danych o probce
		dl = self._linnia[0].Length
		max_dl = dl - 15
		temp = self._linnia[0].ToString()
		temp = temp.Substring(15, max_dl)
		ae.Cells[10][2] = temp
		ae1.Cells[10][2] = temp
		ae2.Cells[9][2] = temp
		ae2.Cells[62][2] = temp
		postep = 1
		backgroundWorker1.ReportProgress(postep)
		dl = self._linnia[1].Length
		max_dl = dl - 15
		temp = self._linnia[1].ToString()
		temp = temp.Substring(15, max_dl)
		ae.Cells[10][5] = temp
		ae1.Cells[10][6] = temp
		ae2.Cells[9][7] = temp
		ae2.Cells[62][7] = temp
		postep = 2
		backgroundWorker1.ReportProgress(postep)
		dl = self._linnia[2].Length
		max_dl = dl - 24
		temp = self._linnia[2].ToString()
		temp = temp.Substring(24, max_dl)
		ae.Cells[10][10] = temp
		ae1.Cells[10][10] = temp
		ae2.Cells[9][12] = temp
		ae2.Cells[62][12] = temp
		dl = self._linnia[3].Length
		max_dl = dl - 14
		temp = self._linnia[3].ToString()
		temp = temp.Substring(14, max_dl)
		ae.Cells[11][2] = temp
		ae1.Cells[11][2] = temp
		ae2.Cells[10][2] = temp
		ae2.Cells[63][2] = temp
		dl = self._linnia[4].Length
		max_dl = dl - 15
		temp = self._linnia[4].ToString()
		temp = temp.Substring(15, max_dl)
		ae.Cells[11][5] = temp
		ae1.Cells[11][6] = temp
		ae2.Cells[10][6] = temp
		ae2.Cells[63][6] = temp
		dl = self._linnia[5].Length
		max_dl = dl - 15
		temp = self._linnia[5].ToString()
		temp = temp.Substring(15, max_dl)
		ae.Cells[11][10] = temp
		ae1.Cells[11][10] = temp
		ae2.Cells[10][11] = temp
		ae2.Cells[63][11] = temp
		dl = self._linnia[6].Length
		max_dl = dl - 5
		temp = self._linnia[6].ToString()
		temp = temp.Substring(5, max_dl)
		temp_digit = Single.Parse(temp)
		ae.Cells[24][6] = temp_digit
		dl = self._linnia[7].Length
		max_dl = dl - 5
		temp = self._linnia[7].ToString()
		temp = temp.Substring(5, max_dl)
		temp_digit = Single.Parse(temp)
		ae.Cells[25][6] = temp_digit
		dl = self._linnia[8].Length
		max_dl = dl - 5
		temp = self._linnia[8].ToString()
		temp = temp.Substring(5, max_dl)
		temp_digit = Single.Parse(temp)
		ae.Cells[26][6] = temp_digit
		dl = self._linnia[9].Length
		max_dl = dl - 5
		temp = self._linnia[9].ToString()
		temp = temp.Substring(5, max_dl)
		temp_digit = Single.Parse(temp)
		ae.Cells[27][6] = temp_digit
		postep = 10
		backgroundWorker1.ReportProgress(postep)
		dl = self._linnia[10].Length
		max_dl = dl - 5
		temp = self._linnia[10].ToString()
		temp = temp.Substring(5, max_dl)
		temp_digit = Single.Parse(temp)
		ae.Cells[28][6] = temp_digit
		dl = self._linnia[11].Length
		max_dl = dl - 5
		temp = self._linnia[11].ToString()
		temp = temp.Substring(5, max_dl)
		temp_digit = Single.Parse(temp)
		ae.Cells[29][6] = temp_digit
		dl = self._linnia[12].Length
		max_dl = dl - 26
		temp = self._linnia[12].ToString()
		temp = temp.Substring(26, max_dl)
		temp_digit = Single.Parse(temp)
		ae.Cells[30][6] = temp_digit
		dl = self._linnia[13].Length
		max_dl = dl - 5
		temp = self._linnia[13].ToString()
		temp = temp.Substring(5, max_dl)
		temp_digit = Single.Parse(temp)
		ae.Cells[24][9] = temp_digit
		dl = self._linnia[14].Length
		max_dl = dl - 5
		temp = self._linnia[14].ToString()
		temp = temp.Substring(5, max_dl)
		temp_digit = Single.Parse(temp)
		ae.Cells[25][9] = temp_digit
		dl = self._linnia[15].Length
		max_dl = dl - 5
		temp = self._linnia[15].ToString()
		temp = temp.Substring(5, max_dl)
		temp_digit = Single.Parse(temp)
		ae.Cells[26][9] = temp_digit
		dl = self._linnia[16].Length
		max_dl = dl - 26
		temp = self._linnia[16].ToString()
		temp = temp.Substring(26, max_dl)
		temp_digit = Single.Parse(temp)
		ae.Cells[27][9] = temp_digit
		ae1.Cells[15][1] = temp_digit
		dl = self._linnia[17].Length
		max_dl = dl - 29
		temp = self._linnia[17].ToString()
		temp = temp.Substring(29, max_dl)
		temp_digit = Single.Parse(temp)
		ae.Cells[30][9] = temp_digit
		ae.Cells[35][4] = temp_digit
		dl = self._linnia[18].Length
		max_dl = dl - 26
		temp = self._linnia[18].ToString()
		temp = temp.Substring(26, max_dl)
		temp_digit = Single.Parse(temp)
		ae.Cells[32][2] = temp_digit
		dl = self._linnia[19].Length
		max_dl = dl - 18
		temp = self._linnia[19].ToString()
		temp = temp.Substring(18, max_dl)
		temp_digit = Single.Parse(temp)
		ae.Cells[32][5] = temp_digit
		ae1.Cells[15][6] = temp_digit
		postep = 15
		backgroundWorker1.ReportProgress(postep)
		dl = self._linnia[20].Length
		max_dl = dl - 30
		temp = self._linnia[20].ToString()
		temp = temp.Substring(30, max_dl)
		temp_digit = Single.Parse(temp)
		ae.Cells[32][8] = temp_digit
		postep = 20
		backgroundWorker1.ReportProgress(postep)
		#zapis danych nasycaniu
		poczatek_nas = Array.IndexOf(self._linnia, "Nasycanie")
		poczatek_kons = Array.IndexOf(self._linnia, "Konsolidacja izotropowa")
		if poczatek_nas + 2 < poczatek_kons:
			temp = self._linnia[poczatek_nas + 2].ToString()
			temp1 = temp.Split(';')
			dane = temp1[1].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[22][1] = dane_digit
			dane = temp1[2].ToString()
			dane_digit = Single.Parse(dane)
			dane_digit = Math.Round(dane_digit, 2)
			ae1.Cells[22][3] = dane_digit
			dane = temp1[3].ToString()
			dane_digit = Single.Parse(dane)
			dane_digit = Math.Round(dane_digit, 2)
			ae1.Cells[22][5] = dane_digit
			ae1.Cells[22][7] = "-"
			ae1.Cells[22][8] = "-"
			ae1.Cells[22][10] = "-"
		if poczatek_nas + 3 < poczatek_kons:
			temp = self._linnia[poczatek_nas + 3].ToString()
			temp1 = temp.Split(';')
			dane = temp1[1].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[23][1] = dane_digit
			dane = temp1[2].ToString()
			dane_digit = Single.Parse(dane)
			dane_digit = Math.Round(dane_digit, 2)
			ae1.Cells[23][3] = dane_digit
			dane = temp1[3].ToString()
			dane_digit = Single.Parse(dane)
			dane_digit = Math.Round(dane_digit, 2)
			ae1.Cells[23][5] = dane_digit
			dane = temp1[4].ToString()
			dane_digit = Single.Parse(dane)
			dane_digit = Math.Round(dane_digit, 2)
			ae1.Cells[23][7] = dane_digit
			dane = temp1[5].ToString()
			dane_digit = Single.Parse(dane)
			dane_digit = Math.Round(dane_digit, 2)
			ae1.Cells[23][8] = "-"
			dane = temp1[6].ToString()
			dane_digit = Single.Parse(dane)
			dane_digit = Math.Round(dane_digit, 2)
			ae1.Cells[23][10] = "-"
		if poczatek_nas + 4 < poczatek_kons:
			temp = self._linnia[poczatek_nas + 4].ToString()
			temp1 = temp.Split(';')
			dane = temp1[1].ToString()
			dane_digit = Single.Parse(dane)
			dane_digit = Math.Round(dane_digit, 2)
			ae1.Cells[24][1] = dane_digit
			dane = temp1[2].ToString()
			dane_digit = Single.Parse(dane)
			dane_digit = Math.Round(dane_digit, 2)
			ae1.Cells[24][3] = dane_digit
			dane = temp1[3].ToString()
			dane_digit = Single.Parse(dane)
			dane_digit = Math.Round(dane_digit, 2)
			ae1.Cells[24][5] = dane_digit
			dane = temp1[4].ToString()
			dane_digit = Single.Parse(dane)
			dane_digit = Math.Round(dane_digit, 2)
			ae1.Cells[24][7] = "-"
			dane = temp1[5].ToString()
			dane_digit = Single.Parse(dane)
			dane_digit = Math.Round(dane_digit, 2)
			ae1.Cells[24][8] = dane_digit
			dane = temp1[6].ToString()
			dane_digit = Single.Parse(dane)
			dane_digit = Math.Round(dane_digit, 2)
			ae1.Cells[24][10] = dane_digit
		if poczatek_nas + 5 < poczatek_kons:
			temp = self._linnia[poczatek_nas + 5].ToString()
			temp1 = temp.Split(';')
			dane = temp1[1].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[25][1] = dane_digit
			dane = temp1[2].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[25][3] = dane_digit
			dane = temp1[3].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[25][5] = dane_digit
			dane = temp1[4].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[25][7] = dane_digit
			dane = temp1[5].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[25][8] = "-"
			dane = temp1[6].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[25][10] = "-"
		if poczatek_nas + 6 < poczatek_kons:
			temp = self._linnia[poczatek_nas + 6].ToString()
			temp1 = temp.Split(';')
			dane = temp1[1].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[26][1] = dane_digit
			dane = temp1[2].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[26][3] = dane_digit
			dane = temp1[3].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[26][5] = dane_digit
			dane = temp1[4].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[26][7] = "-"
			dane = temp1[5].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[26][8] = dane_digit
			dane = temp1[6].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[26][10] = dane_digit
		if poczatek_nas + 7 < poczatek_kons:
			temp = self._linnia[poczatek_nas + 7].ToString()
			temp1 = temp.Split(';')
			dane = temp1[1].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[27][1] = dane_digit
			dane = temp1[2].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[27][3] = dane_digit
			dane = temp1[3].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[27][5] = dane_digit
			dane = temp1[4].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[27][7] = dane_digit
			dane = temp1[5].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[27][8] = "-"
			dane = temp1[6].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[27][10] = "-"
		if poczatek_nas + 8 < poczatek_kons:
			temp = self._linnia[poczatek_nas + 8].ToString()
			temp1 = temp.Split(';')
			dane = temp1[1].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[28][1] = dane_digit
			dane = temp1[2].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[28][3] = dane_digit
			dane = temp1[3].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[28][5] = dane_digit
			dane = temp1[4].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[28][7] = "-"
			dane = temp1[5].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[28][8] = dane_digit
			dane = temp1[6].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[28][10] = dane_digit
		if poczatek_nas + 9 < poczatek_kons:
			temp = self._linnia[poczatek_nas + 9].ToString()
			temp1 = temp.Split(';')
			dane = temp1[1].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[29][1] = dane_digit
			dane = temp1[2].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[29][3] = dane_digit
			dane = temp1[3].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[29][5] = dane_digit
			dane = temp1[4].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[29][7] = dane_digit
			dane = temp1[5].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[29][8] = "-"
			dane = temp1[6].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[29][10] = "-"
		if poczatek_nas + 10 < poczatek_kons:
			temp = self._linnia[poczatek_nas + 10].ToString()
			temp1 = temp.Split(';')
			dane = temp1[1].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[30][1] = dane_digit
			dane = temp1[2].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[30][3] = dane_digit
			dane = temp1[3].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[30][5] = dane_digit
			dane = temp1[4].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[30][7] = "-"
			dane = temp1[5].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[30][8] = dane_digit
			dane = temp1[6].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[30][10] = dane_digit
		if poczatek_nas + 11 < poczatek_kons:
			temp = self._linnia[poczatek_nas + 11].ToString()
			temp1 = temp.Split(';')
			dane = temp1[1].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[31][1] = DataBindings
			dane = temp1[2].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[31][3] = dane_digit
			dane = temp1[3].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[31][5] = dane_digit
			dane = temp1[4].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[31][7] = dane_digit
			dane = temp1[5].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[31][8] = "-"
			dane = temp1[6].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[31][10] = "-"
		if poczatek_nas + 12 < poczatek_kons:
			temp = self._linnia[poczatek_nas + 12].ToString()
			temp1 = temp.Split(';')
			dane = temp1[1].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[32][1] = dane_digit
			dane = temp1[2].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[32][3] = dane_digit
			dane = temp1[3].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[32][5] = dane_digit
			dane = temp1[4].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[32][7] = "-"
			dane = temp1[5].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[32][8] = dane_digit
			dane = temp1[6].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[32][10] = dane_digit
		if poczatek_nas + 13 < poczatek_kons:
			temp = self._linnia[poczatek_nas + 13].ToString()
			temp1 = temp.Split(';')
			dane = temp1[1].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[33][1] = dane_digit
			dane = temp1[2].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[33][3] = dane_digit
			dane = temp1[3].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[33][5] = dane_digit
			dane = temp1[4].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[33][7] = dane_digit
			dane = temp1[5].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[33][8] = "-"
			dane = temp1[6].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[33][10] = "-"
		if poczatek_nas + 14 < poczatek_kons:
			temp = self._linnia[poczatek_nas + 14].ToString()
			temp1 = temp.Split(';')
			dane = temp1[1].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[34][1] = dane_digit
			dane = temp1[2].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[34][3] = dane_digit
			dane = temp1[3].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[34][5] = dane_digit
			dane = temp1[4].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[34][7] = "-"
			dane = temp1[5].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[34][8] = dane_digit
			dane = temp1[6].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[34][10] = dane_digit
		if poczatek_nas + 15 < poczatek_kons:
			temp = self._linnia[poczatek_nas + 15].ToString()
			temp1 = temp.Split(';')
			dane = temp1[1].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[35][1] = dane_digit
			dane = temp1[2].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[35][3] = dane_digit
			dane = temp1[3].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[35][5] = dane_digit
			dane = temp1[4].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[35][7] = dane_digit
			dane = temp1[5].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[35][8] = "-"
			dane = temp1[6].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[35][10] = "-"
		if poczatek_nas + 16 < poczatek_kons:
			temp = self._linnia[poczatek_nas + 16].ToString()
			temp1 = temp.Split(';')
			dane = temp1[1].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[36][1] = dane_digit
			dane = temp1[2].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[36][3] = dane_digit
			dane = temp1[3].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[36][5] = dane_digit
			dane = temp1[4].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[36][7] = "-"
			dane = temp1[5].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[36][8] = dane_digit
			dane = temp1[6].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[36][10] = dane_digit
		if poczatek_nas + 17 < poczatek_kons:
			temp = self._linnia[poczatek_nas + 17].ToString()
			temp1 = temp.Split(';')
			dane = temp1[1].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[37][1] = dane_digit
			dane = temp1[2].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[37][3] = dane_digit
			dane = temp1[3].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[37][5] = dane_digit
			dane = temp1[4].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[37][3] = dane_digit
			dane = temp1[5].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[37][8] = "-"
			dane = temp1[6].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[37][10] = "-"
		if poczatek_nas + 18 < poczatek_kons:
			temp = self._linnia[poczatek_nas + 18].ToString()
			temp1 = temp.Split(';')
			dane = temp1[1].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[38][1] = dane_digit
			dane = temp1[2].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[38][3] = dane_digit
			dane = temp1[3].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[38][5] = dane_digit
			dane = temp1[4].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[38][7] = "-"
			dane = temp1[5].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[38][8] = dane_digit
			dane = temp1[6].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[38][10] = dane_digit
		if poczatek_nas + 19 < poczatek_kons:
			temp = self._linnia[poczatek_nas + 19].ToString()
			temp1 = temp.Split(';')
			dane = temp1[1].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[39][1] = dane_digit
			dane = temp1[2].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[39][3] = dane_digit
			dane = temp1[3].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[39][5] = dane_digit
			dane = temp1[4].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[39][7] = dane_digit
			dane = temp1[5].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[39][8] = "-"
			dane = temp1[6].ToString()
			dane_digit = Single.Parse(dane)
			ae1.Cells[39][10] = "-"
		postep = 50
		backgroundWorker1.ReportProgress(postep)
		#zapis koncowy do pliku
		nazwa_temp = self._sciezka
		dl = self._nazwa_pliku.Length
		max_dl = dl - 4
		temp = self._nazwa_pliku
		temp = temp.Substring(0, max_dl)
		nazwa_temp = nazwa_temp + temp
		nazwa_temp = nazwa_temp + ".xlsx"
		pe.SaveAs(nazwa_temp)
		pe.Close(True)
		mojeexcel.Quit()
		backgroundWorker1.CancelAsync()

	def BackgroundWorker1ProgressChanged(self, sender, e):
		progressBar1.Value = e.ProgressPercentage

	def BackgroundWorker1RunWorkerCompleted(self, sender, e):
		label1.Text = "Formularz wygenerowany"

	def InfoToolStripMenuItemClick(self, sender, e):
		MessageBox.Show("Aplikacja tworzy automatycznie formularze z programu MW_LAB")