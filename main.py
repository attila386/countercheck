################################################################################
#
# A program két adatkészletet épít, "balt"-t és "jobbot", két, a felhasználó
# által betallózott .XLSX állományból.
# 
# A bal készlet felépítése a felhasználó által megadott kereső kifejezésen
# alapul, amelyet egy megadott oszlop celláiban keres, mint töredéket.
# Amennyiben egyezést talál, annak a sornak a további, meghatározott celláiból
# eltárolja a gyári szám és számláló értékeket. A bal készlet építéséhez a bal
# felső vezérlőelemeket használja.
# 
# A jobb készlet felépítése ugyanígy történik, csak a jobb felső vezérlőelemek
# felhasználásával.
# 
# Az összehasonlításhoz a két készletből készül egy eredménytábla, amelybe
# bekerül minden sor a bal készletből (gyári szám és számláló) és mellé kerül az
# ugyanehhez a gyári számhoz tartozó számláló a jobb oldaliból. Ha nincs ilyen,
# akkor 0 kerül be jobb oldali számlálóként.
# 
# Az eredménytábla kiírásakor az olyan sorok, amelyek két számláló értéke eltér,
# színes kiemelést kap.
# 
# Ha a betallózott fájlok nyitásakor, betöltésekor, vagy adatfeldolgozáskor hiba
# lép fel, a program egy általános hibaüzenetet ad.
#
# A program nem ellenőrzi a beolvasott adatok megfelelőségét.
# 
# A program automatikusan elmenti és betölti a felhasználó által beállított
# paramétereket egy .json fájlba.
#
################################################################################

from tkinter import ttk
from tkinter import *
from tkinter import filedialog as fd
from datetime import datetime
import json
from manager import *

# Programverzió
version = '0.92'

# ABC a spinboxokhoz
spin_alphabet=('A', 'B', 'C', 'D', 'E',
'F', 'G', 'H', 'I', 'J',
'K', 'L', 'M', 'N', 'O',
'P', 'Q', 'R', 'S', 'T',
'U', 'V', 'W', 'X', 'Y',
'Z' )

# A teljes fájl útvonalakat tartalmazó lista
filepath0 = ''
filepath1 = ''
filepaths = [filepath0, filepath1]

################################################################################

def load_userprefs() :

	# A felhasználó által utoljára használt beállítások betöltése. Ha sikertelen,
	# a program alapértelmezett értékeket állít be.

	userprefs = {}
	
	try :
		with open('userprefs.json') as f :
			userprefs = json.load(f)
			
			c_search_spin0.set(userprefs['key0'])
			pat_entry0.delete(0, END)
			pat_entry0.insert(0, userprefs['pat0'])
			c_ser_spin0.set(userprefs['ser0'])
			c_cnt_spin0.set(userprefs['cnt0'])
			start_spin0.set(userprefs['start0'])
			stop_spin0.set(userprefs['stop0'])
			c_search_spin1.set(userprefs['key1'])
			pat_entry1.delete(0, END)
			pat_entry1.insert(0, userprefs['pat1'])
			c_ser_spin1.set(userprefs['ser1'])
			c_cnt_spin1.set(userprefs['cnt1'])
			start_spin1.set(userprefs['start1'])
			stop_spin1.set(userprefs['stop1'])
	except:
		pass

def app_close() :
	
	# A felhasználó által megadott paraméterek mentése. Ha ez sikertelen
	# az alkalmazás jelzés nélkül kilép.
	userprefs = {}
	
	userprefs.clear()
	userprefs['key0'] = c_search_spin0.get()
	userprefs['pat0'] = pat_entry0.get()
	userprefs['ser0'] = c_ser_spin0.get()
	userprefs['cnt0'] = c_cnt_spin0.get()
	userprefs['start0'] = int(start_spin0.get())
	userprefs['stop0'] = int(stop_spin0.get())
	userprefs['key1'] = c_search_spin1.get()
	userprefs['pat1'] = pat_entry1.get()
	userprefs['ser1'] = c_ser_spin1.get()
	userprefs['cnt1'] = c_cnt_spin1.get()
	userprefs['start1'] = int(start_spin1.get())
	userprefs['stop1'] = int(stop_spin1.get())
	
	try :
		with open('userprefs.json', 'w') as f :
			json.dump(userprefs, f)
	except:
		pass
	
	root.destroy()

def browseFile(idx) :
	
	# Nyit egy fájl tallózó ablakot, a a fájlt a teljes útvonallal együtt
	# elmenti. Emellett kinyeri a fájlnevet, amit kiír. Ha a felhasználó
	# törli a műveletet, akkor a fájlnév helyére a 'Tallózz ...' kerül.
	
	f = fd.askopenfilename()
	try :
		filepaths[idx] = f
		if f != '' :
			filename = f[f.rfind('/')+1 : ]
		else :
			filename = 'Tallózz ...'
		fname_labels[idx].set(filename)
	except:
		pass

def compare() :
	
	# Összeszedi a felhasználó által beállított paramétereket, lepasszolja a
	# menedzser objektumnak, a visszakapott eredmény táblát feldolgozza. Kiírja
	# a sorokat, megjelöli a kívántakat.
	
	key0 = c_search_spin0.get()
	pat0 = pat_entry0.get()
	ser0 = c_ser_spin0.get()
	cnt0 = c_cnt_spin0.get()
	start0 = int(start_spin0.get())
	stop0 = int(stop_spin0.get())
	key1 = c_search_spin1.get()
	pat1 = pat_entry1.get()
	ser1 = c_ser_spin1.get()
	cnt1 = c_cnt_spin1.get()
	start1 = int(start_spin1.get())
	stop1 = int(stop_spin1.get())
	
	tview1.delete(*tview1.get_children())

	count = 0
	differences = 0	
	
	now = datetime.now()
	current_time = now.strftime("%H:%M:%S")
	
	try:
		comptab = oManager.compare(filepaths[0], key0, pat0, ser0, cnt0, start0, stop0,
			filepaths[1], key1, pat1, ser1, cnt1, start1, stop1)
		
		for record in comptab :
			if record[1] != record[2] :
				tview1.insert(parent='', index='end', iid=count,
					values=(count+1, record[0], record[1], record[2]), tags=('diff'))
				differences += 1
			else :
				tview1.insert(parent='', index='end', iid=count,
					values=(count+1, record[0], record[1], record[2]))
			count += 1
			
		status_bar.config(text=f'[{current_time}] Eltérések száma: {str(differences)}')
	except:
		status_bar.config(text=f'[{current_time}] Nem sikerült a betöltés vagy a feldolgozás.')

def panelswap() :
	
	# Felcseréli a bal és jobb panel beállításait.
	
	filepaths[0], filepaths [1] = filepaths[1], filepaths[0]
	
	t = fname_labels[0].get()
	fname_labels[0].set(fname_labels[1].get())
	fname_labels[1].set(t)
	
	t = c_search_spin0.get()
	c_search_spin0.set(c_search_spin1.get())
	c_search_spin1.set(t)
	
	t = pat_entry0.get()
	pat_entry0.delete(0, END)
	pat_entry0.insert(0,pat_entry1.get())
	pat_entry1.delete(0, END)
	pat_entry1.insert(0, t)
	
	t = c_ser_spin0.get()
	c_ser_spin0.set(c_ser_spin1.get())
	c_ser_spin1.set(t)
	
	t = c_cnt_spin0.get()
	c_cnt_spin0.set(c_cnt_spin1.get())
	c_cnt_spin1.set(t)
	
	t = start_spin0.get()
	start_spin0.set(start_spin1.get())
	start_spin1.set(t)
	
	t = stop_spin0.get()
	stop_spin0.set(stop_spin1.get())
	stop_spin1.set(t)
			
################################################################################

# Instantiate Tkinter root
root = Tk()
root.minsize(720, 640)
root.geometry('800x600')
root.title('Számláló összehasonlítás')
root.columnconfigure(0, weight=1)
root.rowconfigure(0, weight=1)

# Kezelő a root WM_DELETE_WINDOW eseményéhez
root.protocol('WM_DELETE_WINDOW', app_close)

# Kép a tallózó gombhoz
browse_img = PhotoImage(file = r'pix/folder24.png')

# Kép a csere gombhoz
swap_img = PhotoImage(file = r'pix/swap24.png')

# Kép az összehasonlító gombhoz
play_img = PhotoImage(file = r'pix/play24.png')

# Style widget
s = ttk.Style()

mainframe=ttk.Frame(root, height=480, width=480)
mainframe.grid(column=0, row=0, sticky='nwse')
# A mainframe tartalmazza a bal és jobb panelokat, a balt a 0 a jobbot az 1
# oszlopban
mainframe.columnconfigure(0, weight=1)
mainframe.columnconfigure(1, weight=1)
# A mainframe 0 sorának nincs külön konfigja (egyelőre). Az 1 sorban vannak a
# gombok, ez nem nyúlik, a 2 sorban van a fő kijelző
mainframe.rowconfigure(1, weight=0)
mainframe.rowconfigure(2, weight=1)

fname_lbl0_txt = StringVar()
fname_lbl0_txt.set('Tallózz ...')
fname_lbl1_txt = StringVar()
fname_lbl1_txt.set('Tallózz ...')

fname_labels = [fname_lbl0_txt, fname_lbl1_txt]

################ A bal oldali paraméter blokk ################

# nav0 Frame
nav0 = ttk.Frame(mainframe)
nav0.grid(column=0, row=0, sticky='nwe')
nav0.columnconfigure(0, weight=1)

# nav0 fájl tallózó Frame
ublk0 = ttk.Frame(nav0, borderwidth=2, relief='groove')
ublk0.grid(column=0, row=0, sticky='nwe')
ublk0.columnconfigure(0, weight=1)

# nav0 keresési paraméterek Frame
lblk0 = ttk.Frame(nav0, borderwidth=2, relief='groove')
lblk0.grid(column=0, row=1, sticky='nwe')
lblk0.columnconfigure(1, weight=1)

# nav0 fájl tallózó
fname_lbl0=ttk.Label(ublk0, textvariable=fname_labels[0], width=16)
browse_btn0=ttk.Button(ublk0, image=browse_img, command=lambda: browseFile(0))

# nav0 fájl tallózó kirakása a képernyőre
fname_lbl0.grid(column=0, row=0, padx=2, pady=2, sticky='we')
browse_btn0.grid(column=1, row=0, padx=2, pady=2)

# nav0 keresési paraméterek
pat_lbl0=ttk.Label(lblk0, text='Keresd:')
pat_entry0=ttk.Entry(lblk0,	justify='right')
pat_entry0.insert(0, '')
c_search_lbl0=ttk.Label(lblk0, text='Oszlopban:')
c_search_spin0=ttk.Spinbox(lblk0, width=2, format='%c', values=spin_alphabet, wrap=True)
c_search_spin0.insert(0, 'A')
c_ser_lbl0=ttk.Label(lblk0, text='Gyári számok:')
c_ser_spin0=ttk.Spinbox(lblk0, width=2, format='%c', values=spin_alphabet, wrap=True)
c_ser_spin0.insert(0, 'A')
c_cnt_lbl0=ttk.Label(lblk0, text='Számlálók:')
c_cnt_spin0=ttk.Spinbox(lblk0, width=2, format='%c', values=spin_alphabet, wrap=True)
c_cnt_spin0.insert(0, 'A')
start_lbl0=ttk.Label(lblk0, text='Első sor:', anchor='w')
start_spin0=ttk.Spinbox(lblk0, width=3, from_=1, to=500, wrap=True)
start_spin0.insert(0, '1')
stop_lbl0=ttk.Label(lblk0, text='Záró sor:')
stop_spin0=ttk.Spinbox(lblk0, width=3, from_=1, to=500, wrap=True)
stop_spin0.insert(0, '500')

# nav0 keresési paraméterek kirakása a képernyőre
pat_lbl0.grid(column=0, row=0, padx=10, pady=2, sticky='we')
pat_entry0.grid(column=1, row=0, padx=10, sticky='we')
c_search_lbl0.grid(column=0, row=1, padx=10, pady=2, sticky='we')
c_search_spin0.grid(column=1, row=1, padx=10, sticky='e')
c_ser_lbl0.grid(column=0, row=2, padx=10, pady=2, sticky='we')
c_ser_spin0.grid(column=1, row=2, padx=10, sticky='e')
c_cnt_lbl0.grid(column=0, row=3, padx=10, pady=2, sticky='we')
c_cnt_spin0.grid(column=1, row=3, padx=10, sticky='e')
start_lbl0.grid(column=0, row=4, padx=10, pady=2, sticky='we')
start_spin0.grid(column=1, row=4, padx=10, sticky='e')
stop_lbl0.grid(column=0, row=5, padx=10, pady=2, sticky='we')
stop_spin0.grid(column=1, row=5, padx=10, sticky='e')

################ A jobb oldali paraméter blokk ################

# nav1 Frame
nav1 = ttk.Frame(mainframe)
nav1.grid(column=1, row=0, sticky='nwe')
nav1.columnconfigure(0, weight=1)

# nav1 fájl tallózó Frame
ublk1 = ttk.Frame(nav1, borderwidth=2, relief='groove')
ublk1.grid(column=0, row=0, sticky='nwe')
ublk1.columnconfigure(0, weight=1)

# nav1 keresési paraméterek Frame
lblk1 = ttk.Frame(nav1, borderwidth=2, relief='groove')
lblk1.grid(column=0, row=1, sticky='nwe')
lblk1.columnconfigure(1, weight=1)

# nav1 fájl tallózó
fname_lbl1=ttk.Label(ublk1, textvariable=fname_labels[1], width=16)
browse_btn1=ttk.Button(ublk1, image=browse_img, command=lambda: browseFile(1))

# nav1 fájl tallózó kirakása a képernyőre
fname_lbl1.grid(column=0, row=0, padx=2, pady=2, sticky='we')
browse_btn1.grid(column=1, row=0, padx=2, pady=2)

# nav1 keresési paraméterek
pat_lbl1=ttk.Label(lblk1, text='Keresd:')
pat_entry1=ttk.Entry(lblk1, justify='right')
pat_entry1.insert(0, '')
c_search_lbl1=ttk.Label(lblk1, text='Oszlopban:')
c_search_spin1=ttk.Spinbox(lblk1, width=2, format='%c', values=spin_alphabet, wrap=True)
c_search_spin1.insert(0, 'A')
c_ser_lbl1=ttk.Label(lblk1, text='Gyári számok:')
c_ser_spin1=ttk.Spinbox(lblk1, width=2, format='%c', values=spin_alphabet, wrap=True)
c_ser_spin1.insert(0, 'A')
c_cnt_lbl1=ttk.Label(lblk1, text='Számlálók:')
c_cnt_spin1=ttk.Spinbox(lblk1, width=2, format='%c', values=spin_alphabet, wrap=True)
c_cnt_spin1.insert(0, 'A')
start_lbl1=ttk.Label(lblk1, text='Első sor:', anchor='w')
start_spin1=ttk.Spinbox(lblk1, width=3, from_=1, to=500, wrap=True)
start_spin1.insert('0', '1')
stop_lbl1=ttk.Label(lblk1, text='Záró sor:')
stop_spin1=ttk.Spinbox(lblk1, width=3, from_=1, to=500, wrap=True)
stop_spin1.insert(0, '500')

# nav1 keresési paraméterek kirakása a képernyőre
pat_lbl1.grid(column=0, row=0, padx=10, pady=2, sticky='we')
pat_entry1.grid(column=1, row=0, padx=10, sticky='we')
c_search_lbl1.grid(column=0, row=1, padx=10, pady=2, sticky='we')
c_search_spin1.grid(column=1, row=1, padx=10, sticky='e')
c_ser_lbl1.grid(column=0, row=2, padx=10, pady=2, sticky='we')
c_ser_spin1.grid(column=1, row=2, padx=10, sticky='e')
c_cnt_lbl1.grid(column=0, row=3, padx=10, pady=2, sticky='we')
c_cnt_spin1.grid(column=1, row=3, padx=10, sticky='e')
start_lbl1.grid(column=0, row=4, padx=10, pady=2, sticky='we')
start_spin1.grid(column=1, row=4, padx=10, sticky='e')
stop_lbl1.grid(column=0, row=5, padx=10, pady=2, sticky='we')
stop_spin1.grid(column=1, row=5, padx=10, sticky='e')

# Keret a funkciógomboknak
fnframe = ttk.Frame(mainframe)
fnframe.grid(column=0, row=1, columnspan=2)

# Panelcsere gomb
swap_btn = ttk.Button(fnframe, image=swap_img, command=panelswap)
swap_btn.grid(column=0, row=0, padx=2, pady=4)

# Végrehajtás gomb
run_btn = ttk.Button(fnframe, image=play_img, command=compare)
#run_btn.grid(column=0, row=1, columnspan=2, pady=4)
run_btn.grid(column=1, row=0, padx=2, pady=4)

# Keret a fő kijelzőnek
dispframe=ttk.Frame(mainframe, borderwidth=2, relief='groove')
dispframe.grid(column=0, row=2, columnspan=2, sticky='nwse')
dispframe.columnconfigure(0, weight=1)
dispframe.rowconfigure(0, weight=1)

# Fő kijelző
tview1 = ttk.Treeview(dispframe)
tview1['columns']=('NR', 'SN1', 'CN1', 'CN2')
# Columns
tview1.column('#0', width=0, stretch=NO)
tview1.column('NR',  width=64, stretch=NO, anchor='center')
tview1.column('SN1', anchor='w')
tview1.column('CN1', anchor='e')
tview1.column('CN2', anchor='e')
# Headings
tview1.heading('NR', text='#')
tview1.heading('SN1', text='Gyári szám')
tview1.heading('CN1', text='Bal számláló')
tview1.heading('CN2', text='Jobb számláló')
# Display them
tview1.grid(column=0, row=0, sticky='nwse')

# Jelölő beállítása sorok jelöléséhez
tview1.tag_configure('diff', background='orange', foreground='white')

# Gördítősáv a fő kijelzőhöz
tscroll1 = ttk.Scrollbar(dispframe, command=tview1.yview)
tscroll1.grid(column=1, row=0, sticky='ns')
tview1.configure(yscrollcommand=tscroll1.set)

# Állapot sor, üzenet sor
status_bar = ttk.Label(mainframe, text='', anchor='w')
status_bar.grid(column=0, row=3, columnspan=2, sticky='we', padx=8)

################################################################################

# A verziószám kiírása
status_bar.config(text=f'Programverzió {version}')

# A felhasználói paraméterek betöltése
load_userprefs()

# Menedzser objektum készítése
oManager = Manager()

# Fő program futtatása
root.mainloop()
