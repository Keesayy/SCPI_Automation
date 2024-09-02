import pyvisa
import openpyxl
import time

### Constants
PRECISION = 3
NB_ESE_BITS = 60
NB_QUES_BITS = 952

def Float_precision_str(n : int) -> str:
	s : str = '#,##0.'
	for i in range(n):
		s += '0'
	return s

def Excel_name(name : str, precision : int, freq_start : float, freq_stop : float, nb_points : int, dwel : float, amplitude : float) -> str:
	s : string = name
	s = s + '-' + ('{:.%df}' % precision).format(freq_start)
	s = s + '-' + ('{:.%df}' % precision).format(freq_stop)
	s = s + '-' + str(nb_points)
	s = s + '-' + ('{:.%df}' % precision).format(dwel)
	s = s + '-' + ('{:.%df}' % precision).format(amplitude)
	s = s + '.xlsx'
	return s;

def Gpid_devices_open():
	rm = pyvisa.ResourceManager()
	print(rm.list_resources(), '\n')

	#Open gpid devices
	power_meter = rm.open_resource('GPIB0::13::INSTR')
	signal_source = rm.open_resource('GPIB0::19::INSTR')

	power_meter.write('SYST:LANG SCPI')
	# signal_source.write('SYST:LANG SCPI')
	signal_source.write('/SCPI')
	time.sleep(0.2)

	print(power_meter.query('*IDN?'), end = "")
	print(signal_source.query('*IDN?'), end = "")
	return rm ,power_meter, signal_source

def Signal_source_init(signal_source) -> None:
	### Signal Source Init
	# signal_source.write('*RST')
	signal_source.write('*CLS')
	signal_source.write('FREQ:MODE LIST')
	signal_source.write('POW:MODE LIST')
	signal_source.write('LIST:TYPE STEP')

	signal_source.write('POW:ATT:AUTO ON')
	signal_source.write('POW:ATT 0 DB')
	signal_source.write('POW:ALC:LEV 0 DB')

def Power_meter_init(power_meter) -> None:
	### Power Meter Init
	NB_DIGIT = 3
	power_meter.write('*CLS')
	power_meter.write('*ESE 1') 
	power_meter.write('UNIT:POW dBm')

	PRINT('ESR : ', power_meter.write('*ESR?'))
	# power_meter.write('SENSE:AVER:COUN:AUTO ON')
	# PRINT('ESR : ', power_meter.query('*ESR?'))

	power_meter.write('DISP:RES %d' % NB_DIGIT)

def Show_parameters(freq_start : float, freq_stop : float, nb_points : int, dwel : float, amplitude : float) -> None:
	print('\nSTARTING ACQUISITION WITH PARAMETERS :')
	print('fstart :', ('{:.%df}' % PRECISION).format(freq_start), 'Ghz')
	print('fstop  : ', ('{:.%df}' % PRECISION).format(freq_stop), 'Ghz')
	print('points :' + str(nb_points))
	print('dwel   :', ('{:.%df}' % PRECISION).format(dwel), 'ms')
	print('amp    :', ('{:.%df}' % PRECISION).format(amplitude), 'dBm')
	print('\n')

def CLOSE_ALL(signal_source, power_meter, excel, rm) -> None:
	signal_source.close()
	power_meter.close()
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
		print(instrument.query('SYST:ERR?'))
		print(instrument_bis.query('SYST:ERR?'))
	else:
		print('Polling finished because timeout of', timeout, 'seconds reached')
	return status	







