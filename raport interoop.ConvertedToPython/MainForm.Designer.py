import clr

# 
# * Utworzone przez SharpDevelop.
# * Użytkownik: m_witowski
# * Data: 2014-04-15
# * Godzina: 09:47
# *
# * Do zmiany tego szablonu użyj Narzędzia | Opcje | Kodowanie | Edycja Nagłówków Standardowych.
# 
class MainForm(object):
	def __init__(self):
		# <summary>
		# Designer variable used to keep track of non-visual components.
		# </summary>
		self._components = None

	def Dispose(self, disposing):
		""" <summary>
		 Disposes resources used by the form.
		 </summary>
		 <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		"""
		if disposing:
			if self._components != None:
				self._components.Dispose()
		self.Dispose(disposing)

	def InitializeComponent(self):
		""" <summary>
		 This method is required for Windows Forms designer support.
		 Do not change the method contents inside the source code editor. The Forms designer might
		 not be able to load this method if it was changed manually.
		 </summary>
		"""
		resources = System.ComponentModel.ComponentResourceManager(clr.GetClrType(MainForm))
		self._menuStrip1 = System.Windows.Forms.MenuStrip()
		self._plikToolStripMenuItem = System.Windows.Forms.ToolStripMenuItem()
		self._otwórzToolStripMenuItem = System.Windows.Forms.ToolStripMenuItem()
		self._infoToolStripMenuItem = System.Windows.Forms.ToolStripMenuItem()
		self._openFileDialog1 = System.Windows.Forms.OpenFileDialog()
		self._textBox1 = System.Windows.Forms.TextBox()
		self._button1 = System.Windows.Forms.Button()
		self._backgroundWorker1 = System.ComponentModel.BackgroundWorker()
		self._progressBar1 = System.Windows.Forms.ProgressBar()
		self._label1 = System.Windows.Forms.Label()
		self._menuStrip1.SuspendLayout()
		self.SuspendLayout()
		# 
		# menuStrip1
		# 
		self._menuStrip1.Items.AddRange(Array[System.Windows.Forms.ToolStripItem]((self._plikToolStripMenuItem, self._infoToolStripMenuItem)))
		self._menuStrip1.Location = System.Drawing.Point(0, 0)
		self._menuStrip1.Name = "menuStrip1"
		self._menuStrip1.Size = System.Drawing.Size(1186, 24)
		self._menuStrip1.TabIndex = 0
		self._menuStrip1.Text = "menuStrip1"
		# 
		# plikToolStripMenuItem
		# 
		self._plikToolStripMenuItem.DropDownItems.AddRange(Array[System.Windows.Forms.ToolStripItem]((self._otwórzToolStripMenuItem)))
		self._plikToolStripMenuItem.Name = "plikToolStripMenuItem"
		self._plikToolStripMenuItem.Size = System.Drawing.Size(38, 20)
		self._plikToolStripMenuItem.Text = "Plik"
		# 
		# otwórzToolStripMenuItem
		# 
		self._otwórzToolStripMenuItem.Name = "otwórzToolStripMenuItem"
		self._otwórzToolStripMenuItem.Size = System.Drawing.Size(112, 22)
		self._otwórzToolStripMenuItem.Text = "Otwórz"
		self._otwórzToolStripMenuItem.Click += self._OtwórzToolStripMenuItemClick
		# 
		# infoToolStripMenuItem
		# 
		self._infoToolStripMenuItem.Name = "infoToolStripMenuItem"
		self._infoToolStripMenuItem.Size = System.Drawing.Size(40, 20)
		self._infoToolStripMenuItem.Text = "Info"
		self._infoToolStripMenuItem.Click += self._InfoToolStripMenuItemClick
		# 
		# openFileDialog1
		# 
		self._openFileDialog1.Filter = "txt|*.txt"
		# 
		# textBox1
		# 
		self._textBox1.Location = System.Drawing.Point(12, 39)
		self._textBox1.Multiline = True
		self._textBox1.Name = "textBox1"
		self._textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both
		self._textBox1.Size = System.Drawing.Size(904, 462)
		self._textBox1.TabIndex = 1
		# 
		# button1
		# 
		self._button1.Font = System.Drawing.Font("Microsoft Sans Serif", 14.25f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((238)))
		self._button1.Location = System.Drawing.Point(947, 141)
		self._button1.Name = "button1"
		self._button1.Size = System.Drawing.Size(203, 91)
		self._button1.TabIndex = 2
		self._button1.Text = "Generuj formularz"
		self._button1.UseVisualStyleBackColor = True
		self._button1.Click += self._Button1Click
		# 
		# backgroundWorker1
		# 
		self._backgroundWorker1.WorkerReportsProgress = True
		self._backgroundWorker1.WorkerSupportsCancellation = True
		self._backgroundWorker1.DoWork += self._BackgroundWorker1DoWork
		self._backgroundWorker1.ProgressChanged += self._BackgroundWorker1ProgressChanged
		self._backgroundWorker1.RunWorkerCompleted += self._BackgroundWorker1RunWorkerCompleted
		# 
		# progressBar1
		# 
		self._progressBar1.Location = System.Drawing.Point(0, 515)
		self._progressBar1.Name = "progressBar1"
		self._progressBar1.Size = System.Drawing.Size(1186, 23)
		self._progressBar1.TabIndex = 3
		# 
		# label1
		# 
		self._label1.Font = System.Drawing.Font("Microsoft Sans Serif", 12f, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((238)))
		self._label1.Location = System.Drawing.Point(947, 48)
		self._label1.Name = "label1"
		self._label1.Size = System.Drawing.Size(203, 43)
		self._label1.TabIndex = 4
		self._label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
		# 
		# MainForm
		# 
		self._AutoScaleDimensions = System.Drawing.SizeF(6f, 13f)
		self._AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		self._ClientSize = System.Drawing.Size(1186, 538)
		self._Controls.Add(self._label1)
		self._Controls.Add(self._progressBar1)
		self._Controls.Add(self._button1)
		self._Controls.Add(self._textBox1)
		self._Controls.Add(self._menuStrip1)
		self._Icon = ((resources.GetObject("$this.Icon")))
		self._MainMenuStrip = self._menuStrip1
		self._Name = "MainForm"
		self._Text = "Formularz TRX"
		self._menuStrip1.ResumeLayout(False)
		self._menuStrip1.PerformLayout()
		self.ResumeLayout(False)
		self.PerformLayout()