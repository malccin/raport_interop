/*
 * Utworzone przez SharpDevelop.
 * Użytkownik: m_witowski
 * Data: 2014-04-15
 * Godzina: 09:47
 * 
 * Do zmiany tego szablonu użyj Narzędzia | Opcje | Kodowanie | Edycja Nagłówków Standardowych.
 */
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
using System.IO;
using System.Globalization;
using System.Resources;
using System.Reflection;
using System.ComponentModel;


namespace raport_interoop
{
	/// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
		public string nazwa_pliku;
		string sciezka;
		string[] dane_z_pliku;
       	bool dane_wczytane;
		int poczatek_nas;			
		int poczatek_kons;
        int koniec_kons;
        string rozsz;
		
		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			

			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		
		void MainFormLoad(object sender, EventArgs e)
		{
			
			
   
   			Thread watek = new Thread(new ThreadStart(splashscreen));
   			watek.Start();
   			Thread.Sleep(2000);
   			watek.Abort();
   			
   			ResourceManager zasob = new ResourceManager("Raport_interoop.Resource1",Assembly.GetExecutingAssembly());
   			   			
   			//pictureBox2.Image = (Bitmap)zasob.GetObject("ajax-loader(1)");
   			toolStripButton1.Image = (Bitmap)zasob.GetObject("folder_open");
   			
   			pictureBox2.Image = (Bitmap)zasob.GetObject("pologne");
   			pictureBox3.Image = (Bitmap)zasob.GetObject("flagaAnglii");
   			
   			this.Activate();
   			
		}
		void splashscreen()
		{
			
			Application.Run(new Loading());
			
		}
		void OtwórzToolStripMenuItemClick(object sender, EventArgs e)
		{
			DialogResult result = openFileDialog1.ShowDialog();
			
			
			if (result == DialogResult.OK)
				
			{
				
			nazwa_pliku = openFileDialog1.FileName;
			sciezka = openFileDialog1.InitialDirectory;
			Read();
			}
			
			else
			{
				return;
			}
		}
		void Read()
		
		{
			string path;
				
				
			path = sciezka + nazwa_pliku;
			
			dane_z_pliku  = System.IO.File.ReadAllLines(@path);
			
			ReplaceAll(dane_z_pliku,".",",");
			
			textBox1.Lines = dane_z_pliku;
			
			rozsz = Path.GetExtension(path);
			
			dane_wczytane = true;
			
			
	}
		void ReplaceAll(string[] items, string oldValue, string newValue)
{		
			
			string temp;
			
			for (int index = 0; index < items.Length; index++){
			
				temp = items[index].ToString();
				
				temp = temp.Replace(".",",");
				
				items[index] = temp;
      		  
      		  
			}
}
		void Button1Click(object sender, EventArgs e)
		{	
			
			
			if (dane_wczytane == true)
			{
			
				if (rozsz.ToLower() == ".txt" ) {
				
				
					
			button1.Enabled = false;
			
			pictureBox1.Visible = true;
			
			//backgroundWorker1.DoWork += new DoWorkEventHandler(BackgroundWorker1DoWork);
			//backgroundWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BackgroundWorker1RunWorkerCompleted);
			//backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(BackgroundWorker1ProgressChanged);
			backgroundWorker1.RunWorkerAsync();
			label1.Text = "";
			progressBar1.Value = 0;
			textBox1.SelectionStart = 0;
			textBox1.SelectionLength = 0;
				}
				
				if (rozsz.ToLower() == ".tab") {
				
				button1.Enabled = false;
				pictureBox1.Visible = true;
				backgroundWorker2.RunWorkerAsync();
				label1.Text = "";
				progressBar1.Value = 0;
				textBox1.SelectionStart = 0;
				textBox1.SelectionLength = 0;
			}
			}
			
			else
				
			{
				
				label1.Text = "Plik nie wczytany";
				
				
			}
		}
		void BackgroundWorker1DoWork(object sender, DoWorkEventArgs e)
		{
			
			
		
			
			
			int postep = 0;
			
			string nazwa_temp;
            
            string temp;
            string[] temp1;
            float dane_digit;
            float temp_digit;
			
            int dl;
            int max_dl;
            string sciezka_form;
            
			
            
            
            Excel.Application  mojeexcel = new Excel.Application();

            mojeexcel.Visible = false;
            
            sciezka_form  = Path.GetDirectoryName(Application.ExecutablePath);
            
        	sciezka_form = sciezka_form + "\\";
        	
        	if (radioButton1.Checked == true) {
        		        		
        	
        	sciezka_form = sciezka_form + "satcon PL.xlsx";
        	}
        	else{
        		
        	sciezka_form = sciezka_form + "satcon EN.xlsx";	
        		
        		
        	}
			
            var pe = mojeexcel.Workbooks.Open(sciezka_form);

            var ae = (Excel.Worksheet)pe.Sheets[1];

            var ae1 = (Excel.Worksheet)pe.Sheets[2];

            var ae2 = (Excel.Worksheet)pe.Sheets[3];
            
            

            //zapis danych o probce
			
            		
            dl = dane_z_pliku[0].Length;
            max_dl = dl - 15;

            temp = dane_z_pliku[0].ToString();

            temp = temp.Substring(15, max_dl);
            
            if (temp != ""){
            	
            ae.Cells[10,2] = temp;
            ae1.Cells[10,2] = temp;
            ae2.Cells[9,2] = temp;
            ae2.Cells[62,2] = temp;	
            
            	
            	
            }
                        
            
            postep = 1;
            backgroundWorker1.ReportProgress(postep);
            
            dl = dane_z_pliku[1].Length;
            max_dl = dl - 15;

            temp = dane_z_pliku[1].ToString();

            temp = temp.Substring(15, max_dl);
			
            if (temp != "") {
            
            ae.Cells[10,5]= temp;
            ae1.Cells[10,6] = temp;
            ae2.Cells[9,7]= temp;
            ae2.Cells[62,7] = temp;
            
            }
            
			
            
            
            	 postep = 2;
            backgroundWorker1.ReportProgress(postep);

            
            
            dl = dane_z_pliku[2].Length;
            max_dl = dl - 24;

            temp = dane_z_pliku[2].ToString();

            temp = temp.Substring(24, max_dl);
			
             if (temp != "") {
            	
            
            ae.Cells[10,10] = temp;
            ae1.Cells[10,10] = temp;
            ae2.Cells[9,12] = temp;
			ae2.Cells[62,12] = temp;
            }
            dl = dane_z_pliku[3].Length;

            max_dl = dl - 14;

            temp = dane_z_pliku[3].ToString();

            temp = temp.Substring(14, max_dl);
			
            
             if (temp != "") {
            	
            	
            
            ae.Cells[11,2]= temp;
            ae1.Cells[11,2] = temp;
            ae2.Cells[10,2] = temp;
			ae2.Cells[63,2] = temp;
            }
            
            
            dl = dane_z_pliku[4].Length;

            max_dl = dl - 15;

            temp = dane_z_pliku[4].ToString();

            temp = temp.Substring(15, max_dl);
			
             if (temp != "") {
            	
            	
            
            ae.Cells[11,5]= temp;
            ae1.Cells[11,6]= temp;
            ae2.Cells[10,6]= temp;
			ae2.Cells[63,6] = temp;
            }
            
            
            dl = dane_z_pliku[5].Length;

            max_dl = dl - 15;

            temp = dane_z_pliku[5].ToString();

            temp = temp.Substring(15, max_dl);
			
             if (temp != "") {
            ae.Cells[11,10]= temp;
            ae1.Cells[11,10] = temp;
            ae2.Cells[10,11]= temp;
            ae2.Cells[63,11] = temp;
            }
            
            
            dl = dane_z_pliku[6].Length;

            max_dl = dl - 5;

            temp = dane_z_pliku[6].ToString();

            temp = temp.Substring(5, max_dl);
			
             if (temp != "") {
            temp_digit = float.Parse(temp);
			
            ae.Cells[24,6] = temp_digit;
            }


            dl = dane_z_pliku[7].Length;

            max_dl = dl - 5;

            temp = dane_z_pliku[7].ToString();

            temp = temp.Substring(5, max_dl);
			
             if (temp != "") {
            	
            
            temp_digit = float.Parse(temp);

            ae.Cells[25,6]= temp_digit;

            }

            dl = dane_z_pliku[8].Length;

            max_dl = dl - 5;

            temp = dane_z_pliku[8].ToString();
			
             
            temp = temp.Substring(5, max_dl);
			if (temp != "") {

            temp_digit = float.Parse(temp);

            ae.Cells[26,6] = temp_digit;
            }	



            dl = dane_z_pliku[9].Length;

            max_dl = dl - 5;

            temp = dane_z_pliku[9].ToString();

			
            temp = temp.Substring(5, max_dl);
			
            
            if (temp != "") {
            temp_digit = float.Parse(temp);

            ae.Cells[27,6]= temp_digit;
            }
            
            
				 postep = 10;
            	backgroundWorker1.ReportProgress(postep);

            dl = dane_z_pliku[10].Length;

            max_dl = dl - 5;

            temp = dane_z_pliku[10].ToString();

            temp = temp.Substring(5, max_dl);
			
            if (temp != "") {
            temp_digit = float.Parse(temp);

            ae.Cells[28,6] = temp_digit;
            }

            dl = dane_z_pliku[11].Length;

            max_dl = dl - 5;

            temp = dane_z_pliku[11].ToString();

            temp = temp.Substring(5, max_dl);
			
            if (temp != "") {
            temp_digit = float.Parse(temp);


            ae.Cells[29,6] = temp_digit;
            }


            dl = dane_z_pliku[12].Length;

            max_dl = dl - 26;

            temp = dane_z_pliku[12].ToString();

            temp = temp.Substring(26, max_dl);
			
            
            if (temp != "") {
            temp_digit = float.Parse(temp);

            ae.Cells[30,6] = temp_digit;

            }
            dl = dane_z_pliku[13].Length;

            max_dl = dl - 5;

            temp = dane_z_pliku[13].ToString();

            temp = temp.Substring(5, max_dl);
			
            if (temp != "") {
            temp_digit = float.Parse(temp);

            ae.Cells[24,9] = temp_digit;

            }
            dl = dane_z_pliku[14].Length;

            max_dl = dl - 5;

            temp = dane_z_pliku[14].ToString();

            temp = temp.Substring(5, max_dl);
            
			if (temp != "") {
            temp_digit = float.Parse(temp);


            ae.Cells[25,9] = temp_digit;
            }


            dl = dane_z_pliku[15].Length;

            max_dl = dl - 5;

            temp = dane_z_pliku[15].ToString();

            temp = temp.Substring(5, max_dl);
			
            if (temp != "") {
            temp_digit = float.Parse(temp);

            ae.Cells[26,9]= temp_digit;

            }

            dl = dane_z_pliku[16].Length;

            max_dl = dl - 26;

            temp = dane_z_pliku[16].ToString();

            temp = temp.Substring(26, max_dl);
			if (temp != "") {
            temp_digit = float.Parse(temp);

            ae.Cells[27,9] = temp_digit;
            ae1.Cells[15,1]= temp_digit;
            }


            dl = dane_z_pliku[17].Length;

            max_dl = dl - 29;

            temp = dane_z_pliku[17].ToString();

            temp = temp.Substring(29, max_dl);

            temp_digit = float.Parse(temp);
            
			if (temp != "") {
            ae.Cells[30,9] = temp_digit;
            ae.Cells[35,4] = temp_digit;
            }
            	
            	
            	
            
            dl = dane_z_pliku[18].Length;

            max_dl = dl - 26;

            temp = dane_z_pliku[18].ToString();

            temp = temp.Substring(26, max_dl);
			if (temp != "") {
            temp_digit = float.Parse(temp);

            ae.Cells[32,2] = temp_digit;
            }

            dl = dane_z_pliku[19].Length;

            max_dl = dl - 18;

            temp = dane_z_pliku[19].ToString();

            temp = temp.Substring(18, max_dl);
			
            if (temp != "") {
            	
            temp_digit = float.Parse(temp);

            ae.Cells[32,5]= temp_digit;
            ae1.Cells[15,6]= temp_digit;
            }
            
			postep = 15;
			
            backgroundWorker1.ReportProgress(postep);

            dl = dane_z_pliku[20].Length;

            max_dl = dl - 30;

            temp = dane_z_pliku[20].ToString();

            temp = temp.Substring(30, max_dl);
			
            if (temp != "") {
            temp_digit = float.Parse(temp);

            ae.Cells[32,8]= temp_digit;
            
            }
			 postep = 20;
            backgroundWorker1.ReportProgress(postep);
            
            
            
            
            					
			//zapis danych nasycaniu
			
			
			
			poczatek_nas = Array.IndexOf(dane_z_pliku,"Nasycanie");
			
			poczatek_kons = Array.IndexOf(dane_z_pliku, "Konsolidacja izotropowa");
			
			koniec_kons = Array.IndexOf(dane_z_pliku, "konsolidacja_koniec");
			
						
					
						if (poczatek_nas + 2 < poczatek_kons  ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_nas + 2].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
										
						ae1.Cells[22,1] = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[22,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[22,5]  = dane_digit;
						
												
						ae1.Cells[22,7]  = "-";
						
						
						
						ae1.Cells[22,8] = "-";
						
						
						
						ae1.Cells[22,10]  = "-";
				
					}
			
					if (poczatek_nas + 3 < poczatek_kons  ) {
						
						string dane;
						
						
						
						temp = dane_z_pliku[poczatek_nas + 3].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						
						
						ae1.Cells[23,1] = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[23,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[23,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[23,7]  = dane_digit;
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[23,8] = "-";
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[23,10]  = "-";
				
					}
					
						
			
					if (poczatek_nas + 4 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 4].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[24,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[24,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[24,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[24,7]  = "-";
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[24,8]= dane_digit;
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[24,10]  = dane_digit;
				
					}
			
					if (poczatek_nas + 5 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 5].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[25,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[25,3] = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[25,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[25,7]  = dane_digit;
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[25,8]  = "-";
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[25,10]  = "-";
				
					}
			
					if (poczatek_nas + 6 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas +6].ToString();
						
						temp1 = temp.Split(';');
						
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[26,1] = dane_digit;
						
						dane = temp1[2].ToString();
						
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[26,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[26,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[26,7]  = "-";
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[26,8]  = dane_digit;
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[26,10]  = dane_digit;
				
					}
			
					if (poczatek_nas + 7 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 7].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[27,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[27,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[27,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[27,7]  = dane_digit;
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[27,8]  = "-";
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[27,10]  = "-";
				
					}
			
					if (poczatek_nas + 8 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 8].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[28,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[28,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[28,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[28,7]  = "-";
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[28,8]  = dane_digit;
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[28,10]  = dane_digit;
				
					}
					
					if (poczatek_nas + 9 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 9].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[29,1] = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[29,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[29,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[29,7]  = dane_digit;
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[29,8]  = "-";
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[29,10] = "-";
				
					}
			
					if (poczatek_nas + 10 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 10].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[30,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[30,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[30,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[30,7] = "-";				
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[30,8]  = dane_digit;
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[30,10]  = dane_digit;
				
					}
			
					if (poczatek_nas + 11 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 11].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[31,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[31,3] = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[31,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[31,7]  = dane_digit;
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[31,8]  = "-";
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[31,10]  = "-";
				
					}
			
					if (poczatek_nas + 12 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 12].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[32,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane); 
						
						
						ae1.Cells[32,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[32,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[32,7]  = "-";
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[32,8]  = dane_digit;
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[32,10]  = dane_digit;
				
					}
			
			
					if (poczatek_nas + 13 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 13].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[33,1]= dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[33,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[33,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[33,7]  = dane_digit;
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[33,8]  = "-";
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[33,10]  = "-";
				
					}
			
					if (poczatek_nas + 14 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 14].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[34,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[34,3] = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[34,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[34,7]  = "-";
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[34,8]  = dane_digit;
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[34,10]  = dane_digit;
				
					}
					if (poczatek_nas + 15 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 15].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[35,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[35,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[35,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[35,7]  = dane_digit;
						
						dane = temp1[5].ToString();
						
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[35,8]  = "-";
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[35,10]  = "-";
				
					}
			
					if (poczatek_nas + 16 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 16].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[36,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[36,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[36,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[36,7]  = "-";
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[36,8] = dane_digit;
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[36,10]  = dane_digit;
				
					}
			
			if (poczatek_nas + 17 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 17].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[37,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[37,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[37,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[37,7]  = dane_digit;
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[37,8]  = "-";
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[37,10]  = "-";
				
					}
			
			if (poczatek_nas + 18 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 18].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[38,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[38,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[38,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[38,7]  ="-";
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[38,8]  = dane_digit;
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[38,10]  = dane_digit;
				
					}
			
			if (poczatek_nas + 19 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 19].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[39,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[39,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[39,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[39,7] = dane_digit;
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[39,8]  = "-";
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[39,10]  = "-";
				
					}
			
			 postep = 50;
            backgroundWorker1.ReportProgress(postep);
            
            
            
            //zapis konsolidacji
            
            
         		     	if (poczatek_kons  + 1 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 1].ToString();
						
						temp1 = temp.Split(':');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[13,4] = dane_digit;
           				 }
						
						if (poczatek_kons  + 2 < koniec_kons   ) {
						string dane;
						
						
						temp = dane_z_pliku[poczatek_kons + 2].ToString();
						
						temp1 = temp.Split(':');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[14,4] = dane_digit;
						
            			}
						
						if (poczatek_kons  + 3 < koniec_kons   ) {
						string dane;
										
						
						
						temp = dane_z_pliku[poczatek_kons + 3].ToString();
						
						temp1 = temp.Split(':');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[15,4] = dane_digit;
						
          				}
						
						if (poczatek_kons  + 4 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 4].ToString();
						
						temp1 = temp.Split(':');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[16,4] = dane_digit;
				
					}
            
            
            			if (poczatek_kons  + 6 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 6].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[13,10] = dane_digit;
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[13,8] = "0";
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[13,12] = dane_digit;
            
          		    }
            
            
            		if (poczatek_kons  + 7 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 7].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[14,10] = dane_digit;
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[14,8] = dane_digit;
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[14,12] = dane_digit;
            
          		    }
            
         			   if (poczatek_kons  + 8 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 8].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[15,10] = dane_digit;
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[15,8] = dane_digit;
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[15,12] = dane_digit;
            
          		    }
            
         			   if (poczatek_kons  + 9 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 9].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[16,10] = dane_digit;
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[16,8] = dane_digit;
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[16,12] = dane_digit;
            
          		    }
            
           			 if (poczatek_kons  + 10 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 10].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[17,10] = dane_digit;
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[17,8] = dane_digit;
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[17,12] = dane_digit;
            
          		    }
            
            
            		if (poczatek_kons  + 11 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 11].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[18,10] = dane_digit;
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[18,8] = dane_digit;
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[18,12] = dane_digit;
            
          		    }
            
            
          			  if (poczatek_kons  + 12 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 12].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[19,10] = dane_digit;
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[19,8] = dane_digit;
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[19,12] = dane_digit;
            
          		    }
            
           				 if (poczatek_kons  + 13 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 13].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[20,10] = dane_digit;
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[20,8] = dane_digit;
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[20,12] = dane_digit;
            
          		    }
            
             			if (poczatek_kons  + 14 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 14].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[21,10] = dane_digit;
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[21,8] = dane_digit;
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[21,12] = dane_digit;
            
          		    }
            
            			 if (poczatek_kons  + 15 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 15].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[22,10] = dane_digit;
						
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[22,8] = dane_digit;
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[22,12] = dane_digit;
            
          		    }
            
            		 if (poczatek_kons  + 16 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 16].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[23,10] = dane_digit;
						
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[23,8] = dane_digit;
						
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[23,12] = dane_digit;
            
          		    }
            
           			  if (poczatek_kons  + 17 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 17].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[24,10] = dane_digit;
						
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[24,8] = dane_digit;
						
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[24,12] = dane_digit;
            
          		    }
            
          			   if (poczatek_kons  + 18 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 18].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[25,10] = dane_digit;
						
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[25,8] = dane_digit;
						
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[25,12] = dane_digit;
            
          		    }
           	 
          			   if (poczatek_kons  + 19 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 19].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[26,10] = dane_digit;
						
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[26,8] = dane_digit;
						
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[26,12] = dane_digit;
            
          		    }
            
           			  if (poczatek_kons  + 20 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 20].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[27,10] = dane_digit;
						
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[27,8] = dane_digit;
						
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[27,12] = dane_digit;
            
          		    }
            
           			  if (poczatek_kons  + 21 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 21].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[28,10] = dane_digit;
						
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[28,8] = dane_digit;
						
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[28,12] = dane_digit;
            
          		    }
            
          		   if (poczatek_kons  + 22 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 22].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[29,10] = dane_digit;
						
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[29,8] = dane_digit;
						
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[29,12] = dane_digit;
            
          		    }
            //zapis koncowy do pliku
			
            postep = 75;
            backgroundWorker1.ReportProgress(postep);
            

            nazwa_temp = sciezka;


            dl = nazwa_pliku.Length;

            max_dl = dl - 4;

            temp = nazwa_pliku;

            temp = temp.Substring(0, max_dl);

            nazwa_temp = nazwa_temp + temp;

			if (radioButton1.Checked == true) {
				
			
            nazwa_temp = nazwa_temp + ".xlsx";
            }
            else
            {
            	
            	
            nazwa_temp = nazwa_temp + " EN.xlsx";	
            	
            }
			
            pe.SaveAs(nazwa_temp);

            pe.Close(true);

            mojeexcel.Quit();
			
            
            backgroundWorker1.CancelAsync();
            

           					
		
						
}
		void BackgroundWorker1ProgressChanged(object sender, ProgressChangedEventArgs e)
		{
			
			if (0 == (e.ProgressPercentage % 5))
    {
        progressBar1.Value = e.ProgressPercentage;
    }
  
    
			
			
		}
		void BackgroundWorker1RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
		{
			pictureBox1.Visible = false;
			
			label1.Text = "Formularz wygenerowany";
			
			progressBar1.Value = 0;
			
			
			
			button1.Enabled = true;
			
		}
		void InfoToolStripMenuItemClick(object sender, EventArgs e)
		{
			MessageBox.Show("Wersja 0.5" + Environment.NewLine +  "Status wersji: beta" + Environment.NewLine + "Autor: Marcin Witowski - 2015");
		}
		void MainFormFormClosing(object sender, FormClosingEventArgs e)
		{
			if (MessageBox.Show("Zamknąć aplikację ?"," Wyjście",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == DialogResult.No)
			{
				
				e.Cancel = true;
				
			}
			
			
			
		}
		void BackgroundWorker2DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
		{
			
			int postep = 0;
			string nazwa_temp;
            string temp;
            string[] temp1;
            float dane_digit;
            int dl;
            int max_dl;
            string sciezka_form;
            
			int Physical_Properties;
			int Cell_Pressure_Increment0;
			int Back_Pressure_Increment0;
			int Cell_Pressure_Increment1;
			int Back_Pressure_Increment1;
			int Cell_Pressure_Increment2;
			int Back_Pressure_Increment2;
			int Cell_Pressure_Increment3;
			int Back_Pressure_Increment3;
			int Cell_Pressure_Increment4;
			int Back_Pressure_Increment4;
			int Cell_Pressure_Increment5;
			int Back_Pressure_Increment5;
			int Cell_Pressure_Increment6;
			int Back_Pressure_Increment6;
			int Cell_Pressure_Increment7;
			int Back_Pressure_Increment7;
			int Cell_Pressure_Increment8;
			int Back_Pressure_Increment8;
			int Cell_Pressure_Increment9;
			int Back_Pressure_Increment9;
			int Cell_Pressure_Increment10;
			int Back_Pressure_Increment10;
			int Cell_Pressure_Increment11;
			int Back_Pressure_Increment11;
			int Consolidation;
			int Compression;
			
            string dane;
            float srednica;
            float powierzchnia;
            float wysokosc;
            float objetosc;
            float masa;
            float gest;
            
            float cell_pressure;
            float pore_pressure;
            float b_skempton;
            float delta_h_sat;
            float delta_v_sat;
            
            string czas_0 = "0.000000";
            string czas_1 = "0.447214";
            string czas_2 = "0.562731";
            string czas_3 = "0.707107";
            string czas_4 = "0.795822";
            string czas_5 = "0.894427";
            string czas_6 = "1.000000";
            string czas_7 = "1.414214";
            string czas_8 = "2.000000";
            string czas_9 = "2.828427";
            string czas_10 = "3.872983";
            string czas_11 = "5.477226";
            string czas_12 = "7.745967";
            string czas_13 = "10.954451";
            string czas_14 = "15.491933";
            string czas_15 = "21.908902";
            string czas_16 = "37.947332";
            
            
            
            Excel.Application  mojeexcel = new Excel.Application();

            mojeexcel.Visible = false;
            
            sciezka_form  = Path.GetDirectoryName(Application.ExecutablePath);
            
        	sciezka_form = sciezka_form + "\\";
        	
        	sciezka_form = sciezka_form + "TRX_Formularze.xlsx";
			
            var pe = mojeexcel.Workbooks.Open(sciezka_form );

            var ae = (Excel.Worksheet)pe.Sheets[1];

            var ae1 = (Excel.Worksheet)pe.Sheets[2];

            var ae2 = (Excel.Worksheet)pe.Sheets[3];

			//dane poczatkowe
			
			
			Physical_Properties   = Array.IndexOf(dane_z_pliku,"Physical Properties");
			
			//masa
			
			temp = dane_z_pliku[Physical_Properties + 1 ].ToString();
			
			temp1 = temp.Split('	');
			
			dane = temp1[1].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
			
			
			
			ae.Cells[30,9] = dane_digit;
            ae.Cells[35,4] = dane_digit;
            
            masa = dane_digit;
            //srednica
            
            temp = dane_z_pliku[Physical_Properties + 2 ].ToString();
			
			temp1 = temp.Split('	');
			
			dane = temp1[1].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
			
			ae.Cells[30,6] = dane_digit;
            
			srednica = dane_digit;
			
			powierzchnia = (float) ((3.1415 * (srednica * srednica)) / 4)/100  ;
			
			ae.Cells[32,2] = powierzchnia;
			
            
            //wysokosc  gestosc objetosc
            
            temp = dane_z_pliku[Physical_Properties + 3 ].ToString();
			
			temp1 = temp.Split('	');
			
			dane = temp1[1].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
			
			ae.Cells[27,9] = dane_digit;
			
			ae1.Cells[15,1] = dane_digit;
			
			wysokosc = dane_digit;
						
			objetosc = (float) (((3.1415 * (srednica * srednica))/4)*wysokosc) / 1000;
			
			ae.Cells[32,5] = objetosc;
			
			ae1.Cells[15,6] = objetosc;
			
			gest = masa / objetosc;
			
			ae.Cells[32,8] = gest;
			
			
			//nasycanie
			
			
			
			Cell_Pressure_Increment0   = Array.IndexOf(dane_z_pliku,"Saturation (Cell Pressure Increment)");
			
			Cell_Pressure_Increment1   = Array.IndexOf(dane_z_pliku,"Saturation (Cell Pressure Increment)1");
			
			Cell_Pressure_Increment2   = Array.IndexOf(dane_z_pliku,"Saturation (Cell Pressure Increment)2");
			
			Cell_Pressure_Increment3   = Array.IndexOf(dane_z_pliku,"Saturation (Cell Pressure Increment)3");
			
			Cell_Pressure_Increment4   = Array.IndexOf(dane_z_pliku,"Saturation (Cell Pressure Increment)4");
			
			Cell_Pressure_Increment5   = Array.IndexOf(dane_z_pliku,"Saturation (Cell Pressure Increment)5");
			
			Cell_Pressure_Increment6   = Array.IndexOf(dane_z_pliku,"Saturation (Cell Pressure Increment)6");
			
			Cell_Pressure_Increment7   = Array.IndexOf(dane_z_pliku,"Saturation (Cell Pressure Increment)7");
			
			Cell_Pressure_Increment8   = Array.IndexOf(dane_z_pliku,"Saturation (Cell Pressure Increment)8");
			
			Cell_Pressure_Increment9   = Array.IndexOf(dane_z_pliku,"Saturation (Cell Pressure Increment)9");
			
			Cell_Pressure_Increment10   = Array.IndexOf(dane_z_pliku,"Saturation (Cell Pressure Increment)10");
			
			Cell_Pressure_Increment11   = Array.IndexOf(dane_z_pliku,"Saturation (Cell Pressure Increment)11");
			
			Back_Pressure_Increment0 = Array.IndexOf(dane_z_pliku,"Saturation (Back Pressure Increment)");
			
			Back_Pressure_Increment1 = Array.IndexOf(dane_z_pliku,"Saturation (Back Pressure Increment)1");
			
			Back_Pressure_Increment2 = Array.IndexOf(dane_z_pliku,"Saturation (Back Pressure Increment)2");
			
			Back_Pressure_Increment3 = Array.IndexOf(dane_z_pliku,"Saturation (Back Pressure Increment)3");
			
			Back_Pressure_Increment4 = Array.IndexOf(dane_z_pliku,"Saturation (Back Pressure Increment)4");
			
			Back_Pressure_Increment5 = Array.IndexOf(dane_z_pliku,"Saturation (Back Pressure Increment)5");
			
			Back_Pressure_Increment6 = Array.IndexOf(dane_z_pliku,"Saturation (Back Pressure Increment)6");
			
			Back_Pressure_Increment7 = Array.IndexOf(dane_z_pliku,"Saturation (Back Pressure Increment)7");
			
			Back_Pressure_Increment8 = Array.IndexOf(dane_z_pliku,"Saturation (Back Pressure Increment)8");
			
			Back_Pressure_Increment9 = Array.IndexOf(dane_z_pliku,"Saturation (Back Pressure Increment)9");
			
			Back_Pressure_Increment10 = Array.IndexOf(dane_z_pliku,"Saturation (Back Pressure Increment)10");
			
			Back_Pressure_Increment11 = Array.IndexOf(dane_z_pliku,"Saturation (Back Pressure Increment)11");
			
			Consolidation = Array.IndexOf(dane_z_pliku,"Consolidation");
			
			Compression = Array.IndexOf(dane_z_pliku,"Compression");
			
			
			postep = 25;
			
			
			//etap start
			
			
			
			
			temp = dane_z_pliku[Cell_Pressure_Increment0 + 1  ].ToString();
			
			temp1 = temp.Split('	');
			
			//cell pressure
			
			dane = temp1[1].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
			ae1.Cells[22,1] = dane_digit;
            
			cell_pressure = dane_digit;
			
			
			//pore pressure
			
			temp = dane_z_pliku[Cell_Pressure_Increment0 + 2  ].ToString();
			
			temp1 = temp.Split('	');
			
			dane = temp1[1].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
						
			ae1.Cells[23,5] = dane_digit;
			
						
			
			//etap 0 cell
			
			if (Back_Pressure_Increment0 > 0) {
				
				temp = dane_z_pliku[Back_Pressure_Increment0 - 2 ].ToString();
			}
			
			else{
				
				temp = dane_z_pliku[Consolidation -2].ToString();
				
			}
			
			
			temp1 = temp.Split('	');
			
			//cell pressure
			
			dane = temp1[2].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
			ae1.Cells[23,1] = dane_digit;
            
			cell_pressure = dane_digit;
			
			
			//pore pressure
			
			dane = temp1[3].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
						
			ae1.Cells[23,5] = dane_digit;
			
			pore_pressure = dane_digit;
			
			b_skempton = (pore_pressure / cell_pressure);
			
			ae1.Cells[23,7] =(float) Math.Round(b_skempton, 2);
			
			
			if (Back_Pressure_Increment1 < 0) {
				
				goto konsolidacja;
			}
			
			
			
			//etap 0 back
			
			
			temp = dane_z_pliku[Cell_Pressure_Increment1 - 2 ].ToString();
			
			temp1 = temp.Split('	');
			
			//cell pressure
			
			dane = temp1[2].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
			ae1.Cells[24,1] = dane_digit;
            
			cell_pressure = dane_digit;
			
			//back pressure
					
			dane = temp1[4].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
						
			ae1.Cells[24,3] = dane_digit;
           
			//pore pressure
			
			dane = temp1[3].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
						
			ae1.Cells[24,5] = dane_digit;
			
			pore_pressure = dane_digit;
			
			//delta v
			
			dane = temp1[6].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
						
			ae1.Cells[24,10] = Math.Abs(dane_digit);
			
			delta_v_sat = Math.Abs(dane_digit);
			
			//delta h
			
			delta_h_sat = ((float) Math.Round((delta_v_sat * (wysokosc /10)) / (3 * objetosc),2)) * 10;
			
			ae1.Cells[24,8] = delta_h_sat;
			
			
			
			
			
			//etap 1 cell
			
			if (Back_Pressure_Increment1 > 0) {
				
				temp = dane_z_pliku[Back_Pressure_Increment1 - 2 ].ToString();
			}
			
			else{
				
				temp = dane_z_pliku[Consolidation -2].ToString();
				
			}
			
						
			
			temp1 = temp.Split('	');
			
			//cell pressure
			
			dane = temp1[2].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
			ae1.Cells[25,1] = dane_digit;
            
			cell_pressure = dane_digit;
			
			
			//pore pressure
			
			dane = temp1[3].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
						
			ae1.Cells[25,5] = dane_digit;
			
			pore_pressure = dane_digit;
			
			b_skempton = (pore_pressure / cell_pressure);
			
			ae1.Cells[25,7] =(float) Math.Round(b_skempton, 2);
			
			if (Back_Pressure_Increment1 < 0) {
				
				goto konsolidacja;
			}
			
			
			
			//etap 1 back
			
			
			temp = dane_z_pliku[Cell_Pressure_Increment2 - 2 ].ToString();
			
			temp1 = temp.Split('	');
			
			//cell pressure
			
			dane = temp1[2].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
			ae1.Cells[26,1] = dane_digit;
            
			cell_pressure = dane_digit;
			
			//back pressure
					
			dane = temp1[4].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
						
			ae1.Cells[26,3] = dane_digit;
           
			//pore pressure
			
			dane = temp1[3].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
						
			ae1.Cells[26,5] = dane_digit;
			
			pore_pressure = dane_digit;
			
			//delta v
			
			dane = temp1[6].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
						
			ae1.Cells[26,10] = Math.Abs(dane_digit);
			
			delta_v_sat = Math.Abs(dane_digit);
			
			//delta h
			
			delta_h_sat = ((float) Math.Round((delta_v_sat * (wysokosc /10)) / (3 * objetosc),2)) * 10;
			
			ae1.Cells[26,8] = Math.Abs(delta_h_sat);
			
			
			
			
			
			//etap 2 cell
			
			if (Back_Pressure_Increment2 > 0) {
				
				temp = dane_z_pliku[Back_Pressure_Increment2 - 2 ].ToString();
			}
			
			else{
				
				temp = dane_z_pliku[Consolidation -2].ToString();
				
			}
			
			temp1 = temp.Split('	');
			
			//cell pressure
			
			dane = temp1[2].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
			ae1.Cells[27,1] = dane_digit;
            
			cell_pressure = dane_digit;
			
				
			dane = temp1[3].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
						
			ae1.Cells[27,5] = dane_digit;
			
			pore_pressure = dane_digit;
			
			b_skempton = (pore_pressure / cell_pressure);
			
			ae1.Cells[27,7] =(float) Math.Round(b_skempton, 2);
			
			if (Back_Pressure_Increment2 < 0) {
				
				goto konsolidacja;
			}
			
			//etap 2 back
			
			
			temp = dane_z_pliku[Cell_Pressure_Increment3 - 2 ].ToString();
			
			temp1 = temp.Split('	');
			
			//cell pressure
			
			dane = temp1[2].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
			ae1.Cells[28,1] = dane_digit;
            
			cell_pressure = dane_digit;
			
			//back pressure
					
			dane = temp1[4].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
						
			ae1.Cells[28,3] = dane_digit;
           
			//pore pressure
			
			dane = temp1[3].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
						
			ae1.Cells[28,5] = dane_digit;
			
			pore_pressure = dane_digit;
			
			//delta v
			
			dane = temp1[6].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
						
			ae1.Cells[28,10] = Math.Abs(dane_digit);
			
			delta_v_sat = Math.Abs(dane_digit);
			
			//delta h
			
			delta_h_sat = ((float) Math.Round((delta_v_sat * (wysokosc /10)) / (3 * objetosc),2)) * 10;
			
			ae1.Cells[28,8] = Math.Abs(delta_h_sat);
			
			
			
			//etap 3 cell
			
			if (Back_Pressure_Increment3 > 0) {
				
				temp = dane_z_pliku[Back_Pressure_Increment3 - 2 ].ToString();
			}
			
			else{
				
				temp = dane_z_pliku[Consolidation -2].ToString();
				
			}
			
			temp1 = temp.Split('	');
			
			//cell pressure
			
			dane = temp1[2].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
			ae1.Cells[27,1] = dane_digit;
            
			cell_pressure = dane_digit;
			
				
			dane = temp1[3].ToString();
			
			dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
						
			ae1.Cells[27,5] = dane_digit;
			
			pore_pressure = dane_digit;
			
			b_skempton = (pore_pressure / cell_pressure);
			
			ae1.Cells[27,7] =(float) Math.Round(b_skempton, 2);
			
			if (Back_Pressure_Increment3 < 0) {
				
				goto konsolidacja;
			}
			
			
			
			
		postep = 50;	
			
			
		konsolidacja:
			
			
			
			//Konsolidacja
			
			
			int konsolidacja = (Compression  - 1) - (Consolidation + 2);
							
							string[] konsolidacja_etap = new string[konsolidacja];
							
							
							
							Array.Copy(dane_z_pliku,(Consolidation + 2),konsolidacja_etap,0,konsolidacja);
			
								
			//czas_0
			
			
			foreach (string x in konsolidacja_etap) {
				
            	string y = czas_0;
            	            	
            	
            	if (x.Contains(y)) {
            		
            		string[] z;
            			
            		z = x.Split('	');
            		
            		
            		dane = z[2].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[14,4] = Math.Abs(dane_digit);
					
					
					dane = z[4].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[15,4] = Math.Abs(dane_digit);
					
					
            		
            		dane = z[6].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[13,9] = Math.Abs(dane_digit);
					
					
					dane = z[3].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[13,10] = Math.Abs(dane_digit);
					
					ae2.Cells[16,4] = Math.Abs(dane_digit);
					
            		
            			
            		}
            		
			}				
							
			//czas_1
			
			
			foreach (string x in konsolidacja_etap) {
				
            	string y = czas_1;
            	            	
            	
            	if (x.Contains(y)) {
            		
            		string[] z;
            			
            		z = x.Split('	');
            		
            						
            		
            		dane = z[6].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[14,9] = Math.Abs(dane_digit);
					
					
					dane = z[3].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[14,10] = Math.Abs(dane_digit);
					
            		
            			
            		}
            		
			}
			
			
			//czas_2
			
			
			foreach (string x in konsolidacja_etap) {
				
            	string y = czas_2;
            	            	
            	
            	if (x.Contains(y)) {
            		
            		string[] z;
            			
            		z = x.Split('	');
            		
            						
            		
            		dane = z[6].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[15,9] = Math.Abs(dane_digit);
					
					
					dane = z[3].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[15,10] = Math.Abs(dane_digit);
					
            		
            			
            		}
            		
			}

			//czas_3
			
			
			foreach (string x in konsolidacja_etap) {
				
            	string y = czas_3;
            	            	
            	
            	if (x.Contains(y)) {
            		
            		string[] z;
            			
            		z = x.Split('	');
            		
            						
            		
            		dane = z[6].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[16,9] = Math.Abs(dane_digit);
					
					
					dane = z[3].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[16,10] = Math.Abs(dane_digit);
					
            		
            			
            		}
            		
			}
			//czas_4
			
			
			foreach (string x in konsolidacja_etap) {
				
            	string y = czas_4;
            	            	
            	
            	if (x.Contains(y)) {
            		
            		string[] z;
            			
            		z = x.Split('	');
            		
            						
            		
            		dane = z[6].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[17,9] = Math.Abs(dane_digit);
					
					
					dane = z[3].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[17,10] = Math.Abs(dane_digit);
					
            		
            			
            		}
            		
			}
			
			//czas_5
			
			
			foreach (string x in konsolidacja_etap) {
				
            	string y = czas_5;
            	            	
            	
            	if (x.Contains(y)) {
            		
            		string[] z;
            			
            		z = x.Split('	');
            		
            						
            		
            		dane = z[6].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[18,9] = Math.Abs(dane_digit);
					
					
					dane = z[3].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[18,10] = Math.Abs(dane_digit);
					
            		
            			
            		}
            		
			}
			
			//czas_6
			
			
			foreach (string x in konsolidacja_etap) {
				
            	string y = czas_6;
            	            	
            	
            	if (x.Contains(y)) {
            		
            		string[] z;
            			
            		z = x.Split('	');
            		
            						
            		
            		dane = z[6].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[19,9] = Math.Abs(dane_digit);
					
					
					dane = z[3].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[19,10] = Math.Abs(dane_digit);
					
            		
            			
            		}
            		
			}
			
			//czas_7
			
			
			foreach (string x in konsolidacja_etap) {
				
            	string y = czas_7;
            	            	
            	
            	if (x.Contains(y)) {
            		
            		string[] z;
            			
            		z = x.Split('	');
            		
            						
            		
            		dane = z[6].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[20,9] = Math.Abs(dane_digit);
					
					
					dane = z[3].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[20,10] = Math.Abs(dane_digit);
					
            		
            			
            		}
            		
			}
			
			//czas_8
			
			
			foreach (string x in konsolidacja_etap) {
				
            	string y = czas_8;
            	            	
            	
            	if (x.Contains(y)) {
            		
            		string[] z;
            			
            		z = x.Split('	');
            		
            						
            		
            		dane = z[6].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[21,9] = Math.Abs(dane_digit);
					
					
					dane = z[3].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[21,10] = Math.Abs(dane_digit);
					
            		
            			
            		}
            		
			}
			
			
			//czas_9
			
			
			foreach (string x in konsolidacja_etap) {
				
            	string y = czas_9;
            	            	
            	
            	if (x.Contains(y)) {
            		
            		string[] z;
            			
            		z = x.Split('	');
            		
            						
            		
            		dane = z[6].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[22,9] = Math.Abs(dane_digit);
					
					
					dane = z[3].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[22,10] = Math.Abs(dane_digit);
					
            		
            			
            		}
            		
			}
			
			//czas_10
			
			
			foreach (string x in konsolidacja_etap) {
				
            	string y = czas_10;
            	            	
            	
            	if (x.Contains(y)) {
            		
            		string[] z;
            			
            		z = x.Split('	');
            		
            						
            		
            		dane = z[6].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[23,9] = Math.Abs(dane_digit);
					
					
					dane = z[3].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[23,10] = Math.Abs(dane_digit);
					
            		
            			
            		}
            		
			}
			
			
			//czas_11
			
			
			foreach (string x in konsolidacja_etap) {
				
            	string y = czas_11;
            	            	
            	
            	if (x.Contains(y)) {
            		
            		string[] z;
            			
            		z = x.Split('	');
            		
            						
            		
            		dane = z[6].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[24,9] = Math.Abs(dane_digit);
					
					
					dane = z[3].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[24,10] = Math.Abs(dane_digit);
					
            		
            			
            		}
            		
			}
			
			
			//czas_12
			
			
			foreach (string x in konsolidacja_etap) {
				
            	string y = czas_12;
            	            	
            	
            	if (x.Contains(y)) {
            		
            		string[] z;
            			
            		z = x.Split('	');
            		
            						
            		
            		dane = z[6].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[25,9] = Math.Abs(dane_digit);
					
					
					dane = z[3].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[25,10] = Math.Abs(dane_digit);
					
            		
            			
            		}
            		
			}
			
			//czas_13
			
			
			foreach (string x in konsolidacja_etap) {
				
            	string y = czas_13;
            	            	
            	
            	if (x.Contains(y)) {
            		
            		string[] z;
            			
            		z = x.Split('	');
            		
            						
            		
            		dane = z[6].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[26,9] = Math.Abs(dane_digit);
					
					
					dane = z[3].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[26,10] = Math.Abs(dane_digit);
					
            		
            			
            		}
            		
			}
			
			//czas_14
			
			foreach (string x in konsolidacja_etap) {
				
            	string y = czas_14;
            	            	
            	
            	if (x.Contains(y)) {
            		
            		string[] z;
            			
            		z = x.Split('	');
            		
            						
            		
            		dane = z[6].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[27,9] = Math.Abs(dane_digit);
					
					
					dane = z[3].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[27,10] = Math.Abs(dane_digit);
					
            		
            			
            		}
            		
			}
			
			//czas_15
			
			
			foreach (string x in konsolidacja_etap) {
				
            	string y = czas_15;
            	            	
            	
            	if (x.Contains(y)) {
            		
            		string[] z;
            			
            		z = x.Split('	');
            		
            						
            		
            		dane = z[6].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[28,9] = Math.Abs(dane_digit);
					
					
					dane = z[3].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[28,10] = Math.Abs(dane_digit);
					
            		
            			
            		}
            		
			}
			
			//czas_16
			
			
			foreach (string x in konsolidacja_etap) {
				
            	string y = czas_16;
            	            	
            	
            	if (x.Contains(y)) {
            		
            		string[] z;
            			
            		z = x.Split('	');
            		
            						
            		
            		dane = z[6].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[29,9] = Math.Abs(dane_digit);
					
					
					dane = z[3].ToString();
			
					dane_digit = float.Parse(dane,CultureInfo.InvariantCulture);
							
					ae2.Cells[29,10] = Math.Abs(dane_digit);
					
            		
            			
            		}
            		
			}
			
			
			
			
			
			//zapis koncowy do pliku
			
            postep = 100;
            backgroundWorker2.ReportProgress(postep);
            

            nazwa_temp = sciezka;


            dl = nazwa_pliku.Length;

            max_dl = dl - 4;

            temp = nazwa_pliku;

            temp = temp.Substring(0, max_dl);

            nazwa_temp = nazwa_temp + temp;


            nazwa_temp = nazwa_temp + ".xlsx";

			
            pe.SaveAs(nazwa_temp);

            pe.Close(true);

            mojeexcel.Quit();
			
            
            backgroundWorker2.CancelAsync();
            
		}
		void BackgroundWorker2ProgressChanged(object sender, System.ComponentModel.ProgressChangedEventArgs e)
		{
			
			progressBar1.Value = e.ProgressPercentage;
			
		}
		void BackgroundWorker2RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
		{
			pictureBox1.Visible = false;
			
			label1.Text = "Formularz wygenerowany";
			
			progressBar1.Value = 0;
			
			
			
			button1.Enabled = true;
		}
		void ToolStripButton1Click(object sender, EventArgs e)
		{
	
			DialogResult result = openFileDialog1.ShowDialog();
			
			
			if (result == DialogResult.OK)
				
			{
				
			nazwa_pliku = openFileDialog1.FileName;
			sciezka = openFileDialog1.InitialDirectory;
			Read();
			}
			
			else
			{
				return;
			}
			
			
		}
		void BackgroundWorker3DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
		{
	int postep = 0;
			
			string nazwa_temp;
            
            string temp;
            string[] temp1;
            float dane_digit;
            float temp_digit;
			
            int dl;
            int max_dl;
            string sciezka_form;
            
			
            
            
            Excel.Application  mojeexcel = new Excel.Application();

            mojeexcel.Visible = false;
            
            sciezka_form  = Path.GetDirectoryName(Application.ExecutablePath);
            
        	sciezka_form = sciezka_form + "\\";
        	
        	sciezka_form = sciezka_form + "satcon A ang.xlsx";
			
            var pe = mojeexcel.Workbooks.Open(sciezka_form );

            var ae = (Excel.Worksheet)pe.Sheets[1];

            var ae1 = (Excel.Worksheet)pe.Sheets[2];

            var ae2 = (Excel.Worksheet)pe.Sheets[3];

            //zapis danych o probce
			
            		
            dl = dane_z_pliku[0].Length;
            max_dl = dl - 15;

            temp = dane_z_pliku[0].ToString();

            temp = temp.Substring(15, max_dl);
            
            if (temp != ""){
            	
            ae.Cells[10,2] = temp;
            ae1.Cells[10,2] = temp;
            ae2.Cells[9,2] = temp;
            ae2.Cells[62,2] = temp;	
            	
            	
            }
            
			
            
            
            
            
            postep = 1;
            backgroundWorker1.ReportProgress(postep);
            
            dl = dane_z_pliku[1].Length;
            max_dl = dl - 15;

            temp = dane_z_pliku[1].ToString();

            temp = temp.Substring(15, max_dl);
			
            if (temp != "") {
            
            ae.Cells[10,5]= temp;
            ae1.Cells[10,6] = temp;
            ae2.Cells[9,7]= temp;
            ae2.Cells[62,7] = temp;
            
            }
            
			
            
            
            	 postep = 2;
            backgroundWorker1.ReportProgress(postep);

            
            
            dl = dane_z_pliku[2].Length;
            max_dl = dl - 24;

            temp = dane_z_pliku[2].ToString();

            temp = temp.Substring(24, max_dl);
			
             if (temp != "") {
            	
            
            ae.Cells[10,10] = temp;
            ae1.Cells[10,10] = temp;
            ae2.Cells[9,12] = temp;
			ae2.Cells[62,12] = temp;
            }
            dl = dane_z_pliku[3].Length;

            max_dl = dl - 14;

            temp = dane_z_pliku[3].ToString();

            temp = temp.Substring(14, max_dl);
			
            
             if (temp != "") {
            	
            	
            
            ae.Cells[11,2]= temp;
            ae1.Cells[11,2] = temp;
            ae2.Cells[10,2] = temp;
			ae2.Cells[63,2] = temp;
            }
            
            
            dl = dane_z_pliku[4].Length;

            max_dl = dl - 15;

            temp = dane_z_pliku[4].ToString();

            temp = temp.Substring(15, max_dl);
			
             if (temp != "") {
            	
            	
            
            ae.Cells[11,5]= temp;
            ae1.Cells[11,6]= temp;
            ae2.Cells[10,6]= temp;
			ae2.Cells[63,6] = temp;
            }
            
            
            dl = dane_z_pliku[5].Length;

            max_dl = dl - 15;

            temp = dane_z_pliku[5].ToString();

            temp = temp.Substring(15, max_dl);
			
             if (temp != "") {
            ae.Cells[11,10]= temp;
            ae1.Cells[11,10] = temp;
            ae2.Cells[10,11]= temp;
            ae2.Cells[63,11] = temp;
            }
            
            
            dl = dane_z_pliku[6].Length;

            max_dl = dl - 5;

            temp = dane_z_pliku[6].ToString();

            temp = temp.Substring(5, max_dl);
			
             if (temp != "") {
            temp_digit = float.Parse(temp);
			
            ae.Cells[24,6] = temp_digit;
            }


            dl = dane_z_pliku[7].Length;

            max_dl = dl - 5;

            temp = dane_z_pliku[7].ToString();

            temp = temp.Substring(5, max_dl);
			
             if (temp != "") {
            	
            
            temp_digit = float.Parse(temp);

            ae.Cells[25,6]= temp_digit;

            }

            dl = dane_z_pliku[8].Length;

            max_dl = dl - 5;

            temp = dane_z_pliku[8].ToString();
			
             
            temp = temp.Substring(5, max_dl);
			if (temp != "") {

            temp_digit = float.Parse(temp);

            ae.Cells[26,6] = temp_digit;
            }	



            dl = dane_z_pliku[9].Length;

            max_dl = dl - 5;

            temp = dane_z_pliku[9].ToString();

			
            temp = temp.Substring(5, max_dl);
			
            
            if (temp != "") {
            temp_digit = float.Parse(temp);

            ae.Cells[27,6]= temp_digit;
            }
            
            
				 postep = 10;
            	backgroundWorker1.ReportProgress(postep);

            dl = dane_z_pliku[10].Length;

            max_dl = dl - 5;

            temp = dane_z_pliku[10].ToString();

            temp = temp.Substring(5, max_dl);
			
            if (temp != "") {
            temp_digit = float.Parse(temp);

            ae.Cells[28,6] = temp_digit;
            }

            dl = dane_z_pliku[11].Length;

            max_dl = dl - 5;

            temp = dane_z_pliku[11].ToString();

            temp = temp.Substring(5, max_dl);
			
            if (temp != "") {
            temp_digit = float.Parse(temp);


            ae.Cells[29,6] = temp_digit;
            }


            dl = dane_z_pliku[12].Length;

            max_dl = dl - 26;

            temp = dane_z_pliku[12].ToString();

            temp = temp.Substring(26, max_dl);
			
            
            if (temp != "") {
            temp_digit = float.Parse(temp);

            ae.Cells[30,6] = temp_digit;

            }
            dl = dane_z_pliku[13].Length;

            max_dl = dl - 5;

            temp = dane_z_pliku[13].ToString();

            temp = temp.Substring(5, max_dl);
			
            if (temp != "") {
            temp_digit = float.Parse(temp);

            ae.Cells[24,9] = temp_digit;

            }
            dl = dane_z_pliku[14].Length;

            max_dl = dl - 5;

            temp = dane_z_pliku[14].ToString();

            temp = temp.Substring(5, max_dl);
            
			if (temp != "") {
            temp_digit = float.Parse(temp);


            ae.Cells[25,9] = temp_digit;
            }


            dl = dane_z_pliku[15].Length;

            max_dl = dl - 5;

            temp = dane_z_pliku[15].ToString();

            temp = temp.Substring(5, max_dl);
			
            if (temp != "") {
            temp_digit = float.Parse(temp);

            ae.Cells[26,9]= temp_digit;

            }

            dl = dane_z_pliku[16].Length;

            max_dl = dl - 26;

            temp = dane_z_pliku[16].ToString();

            temp = temp.Substring(26, max_dl);
			if (temp != "") {
            temp_digit = float.Parse(temp);

            ae.Cells[27,9] = temp_digit;
            ae1.Cells[15,1]= temp_digit;
            }


            dl = dane_z_pliku[17].Length;

            max_dl = dl - 29;

            temp = dane_z_pliku[17].ToString();

            temp = temp.Substring(29, max_dl);

            temp_digit = float.Parse(temp);
            
			if (temp != "") {
            ae.Cells[30,9] = temp_digit;
            ae.Cells[35,4] = temp_digit;
            }
            	
            	
            	
            
            dl = dane_z_pliku[18].Length;

            max_dl = dl - 26;

            temp = dane_z_pliku[18].ToString();

            temp = temp.Substring(26, max_dl);
			if (temp != "") {
            temp_digit = float.Parse(temp);

            ae.Cells[32,2] = temp_digit;
            }

            dl = dane_z_pliku[19].Length;

            max_dl = dl - 18;

            temp = dane_z_pliku[19].ToString();

            temp = temp.Substring(18, max_dl);
			
            if (temp != "") {
            temp_digit = float.Parse(temp);

            ae.Cells[32,5]= temp_digit;
            ae1.Cells[15,6]= temp_digit;
            }
			 postep = 15;
            backgroundWorker1.ReportProgress(postep);

            dl = dane_z_pliku[20].Length;

            max_dl = dl - 30;

            temp = dane_z_pliku[20].ToString();

            temp = temp.Substring(30, max_dl);
			
            if (temp != "") {
            temp_digit = float.Parse(temp);

            ae.Cells[32,8]= temp_digit;
            
            }
			 postep = 20;
            backgroundWorker1.ReportProgress(postep);
            
            
            					
			//zapis danych nasycaniu
			
			
			
			poczatek_nas = Array.IndexOf(dane_z_pliku,"Nasycanie");
			
			poczatek_kons = Array.IndexOf(dane_z_pliku, "Konsolidacja izotropowa");
			
			koniec_kons = Array.IndexOf(dane_z_pliku, "konsolidacja_koniec");
			
						
					
						if (poczatek_nas + 2 < poczatek_kons  ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_nas + 2].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
										
						ae1.Cells[22,1] = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[22,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[22,5]  = dane_digit;
						
												
						ae1.Cells[22,7]  = "-";
						
						
						
						ae1.Cells[22,8] = "-";
						
						
						
						ae1.Cells[22,10]  = "-";
				
					}
			
					if (poczatek_nas + 3 < poczatek_kons  ) {
						
						string dane;
						
						
						
						temp = dane_z_pliku[poczatek_nas + 3].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						
						
						ae1.Cells[23,1] = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[23,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[23,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[23,7]  = dane_digit;
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[23,8] = "-";
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[23,10]  = "-";
				
					}
					
						
			
					if (poczatek_nas + 4 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 4].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[24,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[24,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[24,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[24,7]  = "-";
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[24,8]= dane_digit;
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						dane_digit = (float) Math.Round(dane_digit,2);
						
						ae1.Cells[24,10]  = dane_digit;
				
					}
			
					if (poczatek_nas + 5 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 5].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[25,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[25,3] = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[25,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[25,7]  = dane_digit;
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[25,8]  = "-";
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[25,10]  = "-";
				
					}
			
					if (poczatek_nas + 6 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas +6].ToString();
						
						temp1 = temp.Split(';');
						
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[26,1] = dane_digit;
						
						dane = temp1[2].ToString();
						
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[26,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[26,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[26,7]  = "-";
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[26,8]  = dane_digit;
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[26,10]  = dane_digit;
				
					}
			
					if (poczatek_nas + 7 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 7].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[27,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[27,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[27,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[27,7]  = dane_digit;
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[27,8]  = "-";
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[27,10]  = "-";
				
					}
			
					if (poczatek_nas + 8 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 8].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[28,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[28,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[28,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[28,7]  = "-";
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[28,8]  = dane_digit;
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[28,10]  = dane_digit;
				
					}
					
					if (poczatek_nas + 9 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 9].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[29,1] = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[29,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[29,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[29,7]  = dane_digit;
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[29,8]  = "-";
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[29,10] = "-";
				
					}
			
					if (poczatek_nas + 10 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 10].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[30,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[30,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[30,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[30,7] = "-";				
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[30,8]  = dane_digit;
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[30,10]  = dane_digit;
				
					}
			
					if (poczatek_nas + 11 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 11].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[31,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[31,3] = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[31,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[31,7]  = dane_digit;
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[31,8]  = "-";
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[31,10]  = "-";
				
					}
			
					if (poczatek_nas + 12 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 12].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[32,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane); 
						
						
						ae1.Cells[32,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[32,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[32,7]  = "-";
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[32,8]  = dane_digit;
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[32,10]  = dane_digit;
				
					}
			
			
					if (poczatek_nas + 13 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 13].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[33,1]= dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[33,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[33,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[33,7]  = dane_digit;
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[33,8]  = "-";
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[33,10]  = "-";
				
					}
			
					if (poczatek_nas + 14 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 14].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[34,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[34,3] = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[34,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[34,7]  = "-";
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[34,8]  = dane_digit;
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[34,10]  = dane_digit;
				
					}
					if (poczatek_nas + 15 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 15].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[35,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[35,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[35,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[35,7]  = dane_digit;
						
						dane = temp1[5].ToString();
						
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[35,8]  = "-";
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[35,10]  = "-";
				
					}
			
					if (poczatek_nas + 16 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 16].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[36,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[36,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[36,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[36,7]  = "-";
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[36,8] = dane_digit;
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[36,10]  = dane_digit;
				
					}
			
			if (poczatek_nas + 17 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 17].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[37,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[37,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[37,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[37,7]  = dane_digit;
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[37,8]  = "-";
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[37,10]  = "-";
				
					}
			
			if (poczatek_nas + 18 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 18].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[38,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[38,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[38,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[38,7]  ="-";
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[38,8]  = dane_digit;
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[38,10]  = dane_digit;
				
					}
			
			if (poczatek_nas + 19 < poczatek_kons  ) {
						
						string dane;
						
						temp = dane_z_pliku[poczatek_nas + 19].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[39,1]  = dane_digit;
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[39,3]  = dane_digit;
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[39,5]  = dane_digit;
						
						dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[39,7] = dane_digit;
						
						dane = temp1[5].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[39,8]  = "-";
						
						dane = temp1[6].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae1.Cells[39,10]  = "-";
				
					}
			
			 postep = 50;
            backgroundWorker1.ReportProgress(postep);
            
            
            
            //zapis konsolidacji
            
            
         		     	if (poczatek_kons  + 1 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 1].ToString();
						
						temp1 = temp.Split(':');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[13,4] = dane_digit;
           				 }
						
						if (poczatek_kons  + 2 < koniec_kons   ) {
						string dane;
						
						
						temp = dane_z_pliku[poczatek_kons + 2].ToString();
						
						temp1 = temp.Split(':');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[14,4] = dane_digit;
						
            			}
						
						if (poczatek_kons  + 3 < koniec_kons   ) {
						string dane;
										
						
						
						temp = dane_z_pliku[poczatek_kons + 3].ToString();
						
						temp1 = temp.Split(':');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[15,4] = dane_digit;
						
          				}
						
						if (poczatek_kons  + 4 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 4].ToString();
						
						temp1 = temp.Split(':');
						
						dane = temp1[1].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[16,4] = dane_digit;
				
					}
            
            
            			if (poczatek_kons  + 6 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 6].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[13,10] = dane_digit;
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[13,9] = "0";
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[13,12] = dane_digit;
            
          		    }
            
            
            		if (poczatek_kons  + 7 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 7].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[14,10] = dane_digit;
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[14,9] = dane_digit;
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[14,12] = dane_digit;
            
          		    }
            
         			   if (poczatek_kons  + 8 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 8].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[15,10] = dane_digit;
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[15,9] = dane_digit;
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[15,12] = dane_digit;
            
          		    }
            
         			   if (poczatek_kons  + 9 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 9].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[16,10] = dane_digit;
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[16,9] = dane_digit;
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[16,12] = dane_digit;
            
          		    }
            
           			 if (poczatek_kons  + 10 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 10].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[17,10] = dane_digit;
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[17,9] = dane_digit;
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[17,12] = dane_digit;
            
          		    }
            
            
            		if (poczatek_kons  + 11 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 11].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[18,10] = dane_digit;
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[18,9] = dane_digit;
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[18,12] = dane_digit;
            
          		    }
            
            
          			  if (poczatek_kons  + 12 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 12].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[19,10] = dane_digit;
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[19,9] = dane_digit;
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[19,12] = dane_digit;
            
          		    }
            
           				 if (poczatek_kons  + 13 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 13].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[20,10] = dane_digit;
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[20,9] = dane_digit;
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[20,12] = dane_digit;
            
          		    }
            
             			if (poczatek_kons  + 14 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 14].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[21,10] = dane_digit;
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[21,9] = dane_digit;
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[21,12] = dane_digit;
            
          		    }
            
            			 if (poczatek_kons  + 15 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 15].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[22,10] = dane_digit;
						
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[22,9] = dane_digit;
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[22,12] = dane_digit;
            
          		    }
            
            		 if (poczatek_kons  + 16 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 16].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[23,10] = dane_digit;
						
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[23,9] = dane_digit;
						
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[23,12] = dane_digit;
            
          		    }
            
           			  if (poczatek_kons  + 17 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 17].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[24,10] = dane_digit;
						
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[24,9] = dane_digit;
						
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[24,12] = dane_digit;
            
          		    }
            
          			   if (poczatek_kons  + 18 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 18].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[25,10] = dane_digit;
						
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[25,9] = dane_digit;
						
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[25,12] = dane_digit;
            
          		    }
           	 
          			   if (poczatek_kons  + 19 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 19].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[26,10] = dane_digit;
						
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[26,9] = dane_digit;
						
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[26,12] = dane_digit;
            
          		    }
            
           			  if (poczatek_kons  + 20 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 20].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[27,10] = dane_digit;
						
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[27,9] = dane_digit;
						
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[27,12] = dane_digit;
            
          		    }
            
           			  if (poczatek_kons  + 21 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 21].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[28,10] = dane_digit;
						
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[28,9] = dane_digit;
						
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[28,12] = dane_digit;
            
          		    }
            
          		   if (poczatek_kons  + 22 < koniec_kons   ) {
						
						string dane;						
						
						
						temp = dane_z_pliku[poczatek_kons + 22].ToString();
						
						temp1 = temp.Split(';');
						
						dane = temp1[2].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[29,10] = dane_digit;
						
						
						
						dane = temp1[3].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[29,9] = dane_digit;
						
						
						
				
           				dane = temp1[4].ToString();
						
						dane_digit = float.Parse(dane);
						
						ae2.Cells[29,12] = dane_digit;
            
          		    }
            //zapis koncowy do pliku
			
            postep = 75;
            backgroundWorker1.ReportProgress(postep);
            

            nazwa_temp = sciezka;


            dl = nazwa_pliku.Length;

            max_dl = dl - 4;

            temp = nazwa_pliku;

            temp = temp.Substring(0, max_dl);

            nazwa_temp = nazwa_temp + temp;


            nazwa_temp = nazwa_temp + ".xlsx";

			
            pe.SaveAs(nazwa_temp);

            pe.Close(true);

            mojeexcel.Quit();
			
            
            backgroundWorker3.CancelAsync();
            

			
			
			
		}
		
		
	}
}
