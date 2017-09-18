# 
# * Utworzone przez SharpDevelop.
# * Użytkownik: m_witowski
# * Data: 2014-04-15
# * Godzina: 09:47
# *
# * Do zmiany tego szablonu użyj Narzędzia | Opcje | Kodowanie | Edycja Nagłówków Standardowych.
# 
from System import *
from System.Windows.Forms import *

class Program(object):
	""" <summary>
	 Class with program entry point.
	 </summary>
	"""
	def Main(args):
		# <summary>
		# Program entry point.
		# </summary>
		Application.EnableVisualStyles()
		Application.SetCompatibleTextRenderingDefault(False)
		Application.Run(MainForm())

	Main = staticmethod(Main)

Program.Main(None)