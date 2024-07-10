from dataset import *

class Manager() :
	
	def __init__(self) :
		
		# Lista készítése, elemei az adatkészletek
		dataset1 = DataSet()
		dataset2 = DataSet()
		self.datasets = [dataset1, dataset2]
		
		# Lista az eredmény táblázatnak
		self.comptab = []
	
	def compare(self,
		f0, key0, pat0, ser0, cnt0, start0, stop0,
		f1, key1, pat1, ser1, cnt1, start1, stop1) :
		
		# Az eredmény táblázat kiürítése
		self.comptab.clear()
		
		# Adatok betöltése az adatkészletekbe
		
		self.datasets[0].load(f0, key0, pat0, ser0, cnt0, start0, stop0)
		self.datasets[1].load(f1, key1, pat1, ser1, cnt1, start1, stop1)
		
		# Összehasonlító tábla készítése
		dataset1 = self.datasets[0].get()
		dataset2 = self.datasets[1].get()
		
		for serA, cntA in dataset1.items() :
			try :
				cntB = dataset2[serA]
			except:
				cntB = '0'
			self.comptab.append([serA, cntA, cntB])
		
		return(self.comptab)
