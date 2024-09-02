import pyvisa
import openpyxl
import time

### Constants
PRECISION = 3
NB_ESE_BITS = 60
NB_QUES_BITS = 952
DEBUG = False

def PRINT(*args, **kwargs):
	if(DEBUG):
		return __builtins__.print(*args, **kwargs)

def Float_precision_str(n : int) -> str:
	s : str = '#,##0.'
	for i in range(n):
		s += '0'
	return s

def Gpid_devices_open():
	rm = pyvisa.ResourceManager()
	print(rm.list_resources(), '\n')

	#Open gpid devices
	measure_source = rm.open_resource('GPIB0::3::INSTR')
	signal_source = rm.open_resource('TCPIP::192.168.10.100::1050::SOCKET')

	signal_source.write('SYST:LANG SCPI')
	# measure_source.write('SYST:LANG SCPI')
	time.sleep(0.2)

	# print(measure_source.query('*IDN?'), end = "")
	return rm ,measure_source, signal_source

### Signal Source Init
def Signal_source_init(signal_source) -> None:
	signal_source.write('*RST')
	signal_source.write('FREQ:PULS OFF')

### Measure Source Init
def Measure_source_init(measure_source) -> None:
	measure_source.write('DA')
	measure_source.write('FP')
	measure_source.write('FR')
	measure_source.write('ES')
	measure_source.write('SR33')
	
	measure_source.write('B1')
	measure_source.write('B2')
	measure_source.write('R0')
	print('STB Frequence metre :', measure_source.read_stb())

def Sweep_freq_measure_precision(signal_source, measure_source):
	print('SWEEP FREQ MEASURE PRECISON')
	
	signal_source.write('OUTP ON')
	signal_source.write('POW %f dBm' % amplitude)

	# Load the Excel workbook
	excel = openpyxl.Workbook()
	sheet = excel.active
	sheet.title = "Data"

	freq = [None] * (11 + 34 + 1)
	freq[0] = 0.1
	freq[1] = 0.2
	freq[2] = 0.4
	freq[3] = 0.75

	freq[9] = 2.5
	freq[10] = 3
	freq[45] = 20.1
	for i in range(4, 9):
		freq[i] = freq[i-1] + 0.25
	
	for i in range(11, 45):
		freq[i] = freq[i-1] + 0.5
	
	precision_string = Float_precision_str(PRECISION)
	sheet['B' + str(1)] = 'Freq (GHz)'
	sheet['C' + str(1)] = 'Precision (GHz)'

	for i in range(len(freq)):
		signal_source.write('FREQ:CW %f GHz' % freq[i])

		start = time.time()
		STB_polling(measure_source, signal_source, condition = 1, timeout = 20, sleepTime = 0.15)
		end = time.time()

		precision = 0

		#Clear ESE
		power_meter.query('*ESR?') 

		print(str(i) + ':', ('{:.%df}' % PRECISION).format(float(freq)), 'GHz | ', ('{:.%df}' % PRECISION).format(float(precision)) , 'GHz | ' , ('{:.%df}' % PRECISION).format(end - start), 's')

		#Excel array
		sheet['B' + str(i+3)] = freq[i]
		sheet['C' + str(i+3)] = float(('{:.6f}').format(float(precision)))
		sheet['B' + str(i+3)].number_format = precision_string
		sheet['C' + str(i+3)].number_format = precision_string

		freq += (freq_stop - freq_start) / (nb_points - 1)

	signal_source.write('OUTP OFF')

	# Save the workbook
	excel.save('EIP548-SW_FP')
	new_workbook_name = 'EIP548-SW_FP'
	return excel, new_workbook_name

def CLOSE_ALL(signal_source, measure_source, excel, rm) -> None:
	signal_source.close()
	measure_source.close()
	excel.close()
	rm.close()

def STB_polling(instrument, instrument_bis, condition = 32, timeout = 1.0, sleepTime = 0.3) -> int:
	PRINT('Polling started')
	end_time = time.time() + timeout # compute the maximal end time
	status = False
	error = False
	stb = instrument.read_stb()
	status = (stb & condition) == condition # first condition check, no need to wait if condition already true
	error = (stb & 1) == 1 # check error, no need to wait if already an error is available

	while not status and time.time() < end_time and not error:
		PRINT('STB : ', stb)
		time.sleep(sleepTime)
		stb = instrument.read_stb()
		status = (stb & condition) == condition #check conditon
		error = (stb & 4) == 4 # check bit 4: Error Message Available
	if status:
		PRINT('Polling finished because STB satisfied condition. STB = ', condition)
	elif error:
		print('Polling finished because Error Occured')
		# print(instrument.query('SYST:ERR?'))
		# print(instrument_bis.query('SYST:ERR?'))
	else:
		print('Polling finished because timeout of', timeout, 'seconds reached')
	return status	

def main() -> int:
	rm, measure_source, signal_source = Gpid_devices_open()

	Signal_source_init(signal_source)
	Measure_source_init(measure_source)

	start_tot = time.time()

	print('\nSTART ACQUISITION ', end = '')
	
	# excel, excel_name = Sweep_freq_measure_precision(signal_source, measure_source)

	end_tot = time.time()
	print('\nTOTAL ACQUISITION TIME : ', ('{:.%df}' % PRECISION).format(end_tot - start_tot), 's')
	
	# print('Saved in : ', excel_name)
	# CLOSE_ALL(signal_source, measure_source, excel, rm)
	return 0

if __name__=="__main__":
    main()







