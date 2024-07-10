import openpyxl

class DataSet() :
	
	def __init__(self) :
		
		# Egy adatkészlet tárolója
		self.dataset = {}
				
	def load(self, f, key, pat, ser, cnt, start, stop) :
		
		# .XLSX megnyitása, aktív munkalap kiválasztása, az adatkészlet tároló
		# ürítése, adatok beolvasása, tárolása
		
		self.workbook = openpyxl.load_workbook(f)
		self.worksheet = self.workbook.active
		self.dataset.clear()
		
		for i in range(start, stop) :
			if ( pat in str(self.worksheet[key+str(i)].value) ) :
				self.dataset[ self.worksheet[ser+str(i)].value ] = self.worksheet[cnt+str(i)].value
			
	def get(self) :
		return(self.dataset)
		
