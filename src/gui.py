import ctypes
import sweep
import sys
import time
import tkinter as tk
from tkinter import ttk

WIDTH = 520
HEIGHT = 320
CELL_WIDTH = 6
TEXT_CELL_WIDTH = 12
PROGRESS_WIDTH = 300
DEBUG = [False]

myappid = 'mycompany.myproduct.subproduct.version' # arbitrary string
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)

def EXIT():
	sys.exit()

def PRINT(*args, **kwargs):
	if(DEBUG[0]):
		return __builtins__.print(*args, **kwargs)
		
def progress_cliked(progress : bool) -> None:
	progress[0] = not progress[0]

def on_select(event):
    selected_item = combo_box.get()
    label.config(text="Selected Item: " + selected_item)

def Show_parameters_bis(freq_start : float, freq_stop : float, nb_points : int, dwel : float, amplitude : float, debug : int, progress : bool) -> None:
	PRECISION = 3
	DEBUG[0] = bool(debug)

	PRINT('\nSTARTING ACQUISITION WITH PARAMETERS :')
	print('fstart :', ('{:.%df}' % PRECISION).format(freq_start), 'Ghz')
	print('fstop  :', ('{:.%df}' % PRECISION).format(freq_stop), 'Ghz')
	print('points : ' + str(nb_points))
	print('dwel   :', ('{:.%df}' % PRECISION).format(dwel), 'ms')
	print('amp    :', ('{:.%df}' % PRECISION).format(amplitude), 'dBm')
	print('debug  : ' + str(debug))
	progress_cliked(progress)
	print(progress)
	print(DEBUG)
	print('\n')

def Sweep_freq(progress, window, power_meter, signal_source, freq_start : float, freq_stop : float, nb_points : int, dwel : float, amplitude : float):

	# signal_source.write('FREQ:STAR %f GHz' % freq_start)
	# print(signal_source.query('FREQ:STAR?'))

	# signal_source.write('FREQ:STOP %f GHz' % freq_stop)
	# print(signal_source.query('FREQ:STOP?'))

	# signal_source.write('SWE:POIN %d' % nb_points)
	# signal_source.write('SWE:DWEL %f MS' % dwel)

	# signal_source.write('POW:AMPL %f dBm' % amplitude)

	# signal_source.write('LIST:TRIG:SOUR BUS')
	# signal_source.write('OUTP:STAT ON')
	# signal_source.write('INIT:CONT ON')

	signal_source.write('OUTP ON')
	print(signal_source.query('*OPC?'))
	signal_source.write('POW %f dBm' % amplitude)

	# Load the Excel workbook
	excel = openpyxl.load_workbook(filename = 'E8257D-67 2024.xlsx')
	PRINT(excel.sheetnames)
	sheet = excel[excel.sheetnames[1]]

	freq = freq_start
	precision_string = Float_precision_str(PRECISION)

	progress.start()
	for i in range(nb_points):
		signal_source.write('FREQ:CW %f GHz' % freq)

		power_meter.write('*CLS')
		power_meter.write('FREQ ' + str(freq) + ' GHz')

		#Measure level
		power_meter.write('TRIG:DEL:AUTO ON')
		power_meter.write('INIT:CONT OFF') 
		power_meter.write('TRIG:SOUR IMM') 
		power_meter.write('INIT') 

		power_meter.write('*OPC')

		start = time.time()
		STB_polling(power_meter, signal_source, timeout = 20, sleepTime = 0.15)
		end = time.time()

		#Clear ESE
		power_meter.query('*ESR?') 

		#Read Level
		level = power_meter.query('FETCH?')

		#Progress Bar GUI
		progress['value'] = i 
		window.update_idletasks()

		print(str(i) + ':', ('{:.%df}' % PRECISION).format(float(freq)), 'GHz | ', ('{:.%df}' % PRECISION).format(float(level)) , 'dBm | ' , ('{:.%df}' % PRECISION).format(end - start), 's')

		#Excel array
		sheet['B' + str(i+3)] = freq
		sheet['C' + str(i+3)] = float(('{:.6f}').format(float(level)))
		sheet['B' + str(i+3)].number_format = precision_string
		sheet['C' + str(i+3)].number_format = precision_string

		# signal_source.write('*TRG')

		freq += (freq_stop - freq_start) / (nb_points - 1)

	# signal_source.write('INIT:CONT OFF')
	# signal_source.write('OUTP:STAT OFF')
	signal_source.write('OUTP OFF')

	# Save the workbook
	new_workbook_name = Excel_name('Giga-2420B', 0, freq_start, freq_stop, nb_points, dwel, amplitude)	
	excel.save(new_workbook_name)

	progress.stop()
	return excel, new_workbook_name

def Sweep(freq_start : float, freq_stop : float, nb_points : int, dwel : float, amplitude : float) -> int:
	rm, power_meter, signal_source = Gpid_devices_open()

	Signal_source_init(signal_source)
	Power_meter_init(power_meter)

	start_tot = time.time()

	print('\nSTART ACQUISITION SWEEP FREQ')
	Show_parameters(freq_start, freq_stop, nb_points, dwel, amplitude)
	
	excel, excel_name = Sweep_freq(power_meter, signal_source, freq_start, freq_stop, nb_points, dwel, amplitude)

	end_tot = time.time()
	print('\nTOTAL ACQUISITION TIME : ', ('{:.%df}' % PRECISION).format(end_tot - start_tot), 's')
	print('Saved in : ', excel_name)

	CLOSE_ALL(signal_source, power_meter, excel, rm)
	return 0

def Main_window():
	window = tk.Tk()
	window.title("CIM AUTOMATION PROGRAM")

	ws = window.winfo_screenwidth() # width of the screen
	hs = window.winfo_screenheight() # height of the screen
	x = (ws/2) - (WIDTH/2)
	y = (hs/2) - (HEIGHT/2)

	window.resizable(True, True)
	window.geometry('%dx%d+%d+%d' % (WIDTH, HEIGHT, x, y))
	window.iconphoto(False, tk.PhotoImage(file = 'CIM_Logo_OpenStar_PNG.png'))
	window.configure(background = "ivory")

	return window

def Menu_window(window):
	menu = tk.Menu(window)
	window.config(menu = menu)
	filemenu = tk.Menu(menu)
	menu.add_cascade(label = 'File', menu = filemenu)
	filemenu.add_command(label = 'New')
	filemenu.add_command(label = 'Open...')
	filemenu.add_separator()
	filemenu.add_command(label = 'Exit', command = window.quit)
	helpmenu = tk.Menu(menu)
	menu.add_cascade(label = 'Help', menu = helpmenu)
	helpmenu.add_command(label = 'About')
	
	return menu, filemenu, helpmenu

def main() -> int:
	# Open WINDOW
	window = Main_window()
	Menu_window(window)

	freq_start = tk.DoubleVar()
	freq_stop = tk.DoubleVar()
	nb_points = tk.IntVar()
	dwel = tk.DoubleVar()
	amplitude = tk.DoubleVar()
	debug = tk.IntVar()
	progress_click = [False]

	tk.Label(window, text = "Freq Start      (GHz):").grid(row = 0, column = 0, sticky = tk.W)
	tk.Label(window, text = "Freq Stop      (GHz):").grid(row = 1, column = 0, sticky = tk.W)
	tk.Label(window, text = "Nb Points               :").grid(row = 2, column = 0, sticky = tk.W)
	tk.Label(window, text = "Dwel               (MS):").grid(row = 3, column = 0, sticky = tk.W)
	tk.Label(window, text = "Ampltiude  (dBm):").grid(row = 4, column = 0, sticky = tk.W)

	tk.Entry(window, width = CELL_WIDTH, textvariable = freq_start).grid(row = 0, column = 1)
	tk.Entry(window, width = CELL_WIDTH, textvariable = freq_stop).grid(row = 1, column = 1)
	tk.Entry(window, width = CELL_WIDTH, textvariable = nb_points).grid(row = 2, column = 1)
	tk.Entry(window, width = CELL_WIDTH, textvariable = dwel).grid(row = 3, column = 1,)
	tk.Entry(window, width = CELL_WIDTH, textvariable = amplitude).grid(row = 4, column = 1)

	freq_sweep_button = tk.Button(window, text = "Freq Sweep", command = lambda: (Show_parameters_bis(freq_start.get(), freq_stop.get(), nb_points.get(), dwel.get(), amplitude.get(), debug.get(), progress_click), print()), width = 15, height = 2)
	freq_sweep_button.place(x = 390, y = 50)	

	cancel_button = tk.Button(window, text = "Cancel", command = EXIT, width = 15, height = 2)
	cancel_button.place(x = 390, y = 100)

	debug_button = tk.Checkbutton(window, text = 'DEBUG', variable = debug)
	debug_button.place(x = 420, y = 20)

	progress = ttk.Progressbar(window, orient = "horizontal", length = PROGRESS_WIDTH, mode = "determinate", maximum = nb_points.get())
	progress.place(x = WIDTH/2 - PROGRESS_WIDTH/2, y = HEIGHT - 50)

	window.mainloop()
	return 0

if __name__=="__main__":
    main()