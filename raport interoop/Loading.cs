/*
 * Utworzone przez SharpDevelop.
 * Użytkownik: m_witowski
 * Data: 2014-04-22
 * Godzina: 11:34
 * 
 * Do zmiany tego szablonu użyj Narzędzia | Opcje | Kodowanie | Edycja Nagłówków Standardowych.
 */
using System;
using System.Drawing;
using System.Windows.Forms;

namespace raport_interoop
{
	/// <summary>
	/// Description of Loading.
	/// </summary>
	public partial class Loading : Form
	{
		public Loading()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
		}
		
		void Timer1Tick(object sender, EventArgs e)
		{
			progressBar1.Increment(1);
			if(progressBar1.Value == 100){
				
				timer1.Stop();
				this.Close();
				
			}
		}
	}
}
