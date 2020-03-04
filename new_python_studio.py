import pyvisa
import time
import random


import numpy as np
import pandas as pd 
from struct import unpack
import socket, re 

from pandas import ExcelWriter
from pandas import ExcelFile









class psa_meas():
    def __init__(self, psa_address):
        print(psa_address)
        self.psa_classy = pyvisa.ResourceManager().open_resource('%s' %(psa_address))
        self.psa_classy.read_termination = '\n'
        self.psa_classy.timeout = 2000
        
    def filename_gen(self):
        self.filename = str(int(random.random()*100000000))
        
    def get_marker(self):
        f=self.psa_classy.query('CALC:MARK:X?')
        amp=self.psa_classy.query('CALC:MARK:Y?')
        marker={'freq':[], 'ampl':[]}
        marker['freq'].append(f)
        marker['ampl'].append(amp)
        return marker

    def single(self):
        #psa.write("TRAC:MODE MINH")
        #psa.write("TRAC:MODE MAXH")
        self.psa_classy.write("TRIG:SOUR IMM")    
        self.psa_classy.write("INIT:CONT OFF")
        self.psa_classy.write("INIT:IMM")
        self.waiting_ocp()    
        self.psa_classy.write("DISP:FSCR ON")

    def peak_search(self):
        self.psa_classy.write(":CALC:MARK1:MAX")
        
    def waiting_ocp(self):
        a = 0
        self.psa_classy.write(':STAT:OPER:ENAB 9')
        while a != 1:
            try:
                a = int(self.psa_classy.query('*OPC?'))
            except:
                time.sleep(0.5) 
                print('waiting...')
                print('OK')

        a = 0
    def save(self, local_filename):
        
        self.filename_gen()
        self.psa_classy.write("MMEM:STOR:SCR 'C:\%s.GIF'" %(self.filename))
        self.psa_classy.write(":MMEM:STOR:TRAC TRACE1,'C:\%s.CSV'" %(self.filename))

        self.waiting_ocp()
        capture = self.psa_classy.query_binary_values(message=":MMEM:DATA? 'C:\%s.GIF'" %(self.filename), container=list, datatype='c')
        with open(r".\%s.GIF" %(local_filename), 'wb') as fp:
            for byte in capture:
                fp.write(byte)
            fp.close()
     
        capture = self.psa_classy.query_binary_values(message=":MMEM:DATA? 'C:\%s.CSV'" %(self.filename), container=list, datatype='c')
        with open(r".\%s.CSV" %(local_filename), 'wb') as fp:
            for byte in capture:
                fp.write(byte)
            fp.close()
        
   
        self.psa_classy.write("MMEMory:DELete 'C:\%s.GIF'" %(self.filename))
        self.psa_classy.write("MMEMory:DELete 'C:\%s.CSV'" %(self.filename))
        

    def maxhold_on(self):
        self.psa_classy.write("TRAC:MODE MINH")
        self.psa_classy.write("TRAC:MODE MAXH")
    def set_x_y_att(self, freq_start, freq_stop, rlev, units, points, rbw, att):
        self.psa_classy.write("POW:ATT %ddb" %(att))
        self.psa_classy.write("UNIT:POW %s"%(units))
        self.psa_classy.write("DISP:WIND:TRAC:Y:RLEV %f" %(rlev))
        self.psa_classy.write("BAND %dHz" %(rbw))
        self.psa_classy.write("FREQ:STAR %d Hz" %(freq_start))
        self.psa_classy.write("FREQ:STOP %d Hz" %(freq_stop))
        self.psa_classy.write("SWE:POIN %d" %(points))       

class psg_control():
    def __init__(self, psg_address):
        self.psg_classy = pyvisa.ResourceManager().open_resource('%s' %(psg_address))
        self.psg_classy.read_termination = '\n'
        self.psg_classy.timeout = 20000
        
    def set_psg(self, freq, power, units):
        if "ASRL" in str(self.psg_classy):
            self.psg_classy.write("VOLT:UNIT %s;" %(units))
            self.psg_classy.write("VOLT %.4f;" %(power))
        else:
            #self.psg_classy.write("SOUR:POW:LEV:IMM:AMPL %.4f" %(power))
            #self.psg_classy.write("POW:AMPL %f %s\n" %(power, units))
            self.psg_classy.write("POW:LEVEL %f %s\n" %(power, units))
        self.psg_classy.write("freq %dHz\n" %(freq))
    #    psg.clear()
    def rf_on(self, state):
        self.psg_classy.write(":OUTPUT:STAT %d;" %(state))
    def get_ampl(self):
        time.sleep(0.05)
        return float(self.psg_classy.query("VOLT?;"))


class scope_meas():
    def __init__(self, scope_adress):
        rm = pyvisa.ResourceManager()
        self.scope_classy = pyvisa.ResourceManager().open_resource('%s' %(scope_adress))
        self.scope_ip = scope_adress[8:-14]
    def get_csv(self, ch, name):

        self.scope_classy.write('DATA:SOU CH%d' %(ch))
        self.scope_classy.write('DATA:WIDTH 1')
        self.scope_classy.write('DATA:ENC RPB')


        ymult = float(self.scope_classy.query('WFMPRE:YMULT?'))
        yzero = float(self.scope_classy.query('WFMPRE:YZERO?'))
        yoff = float(self.scope_classy.query('WFMPRE:YOFF?'))
        xincr = float(self.scope_classy.query('WFMPRE:XINCR?'))


        self.scope_classy.write('CURVE?')
        data = self.scope_classy.read_raw()
        headerlen = 2 + int(data[1])
        header = data[:headerlen]
        ADC_wave = data[headerlen:-1]

        ADC_wave = np.array(unpack('%sB' % len(ADC_wave),ADC_wave))

        Volts = (ADC_wave - yoff) * ymult  + yzero

        Time = np.arange(0, xincr * len(Volts), xincr)

        df = pd.DataFrame(Time, Volts)
        df.to_csv("%s.csv" %(name))



    def get_img(self, name):

        input_buffer = 32 * 1024

        # Open a connection to the instrument's webserver
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.connect((self.scope_ip, 80))

        cmd = b"GET /image.png HTTP/1.0\n\n"
        s.send(cmd)

        # Get the HTTP header
        status = s.recv(input_buffer)
        

        #Read the first chunk of data
        data = s.recv(input_buffer)

        # Check if the content is a png image
        if (b"Content-Type: image/png" not in data):
            # Not proper data
            print("Content returned is not image/png")
            imgData = b""
            
        else: # Content is correct so copy the data to a file

            # Find the length of the png data
            searchObj = re.search(b"Content-Length: (\d+)\r\n", data)
            imgSizeLeft = int(searchObj.group(1))
            
            # Pull the image data out of the first buffer
            startIdx = data.find(b"\x89PNG")
            
            # For the TDS3000B Series, the PNG image data may not come out with the
            # HTTP header
            # If the PNG file header was not found then do another read
            if (startIdx == -1):
                data = s.recv(input_buffer)
                imgData = data[data.find(b"\x89PNG"):]
            else:
                imgData = data[startIdx:]
                
            imgSizeLeft = imgSizeLeft - len(imgData)

            # Read the rest of the image data
            data = s.recv(input_buffer)

            while imgSizeLeft > len(data):
                imgData = b"".join([imgData, data])
                imgSizeLeft = imgSizeLeft - len(data)
                data = s.recv(input_buffer)

                # The TDS3000B Series sends the wrong value for Content-Length.  It
                # sends a value much larger than the real length.
                # If there is no more data then break out of the loop
                if (len(data) == 0):
                    break

            # Add the last chunk of data
            imgData = b"".join([imgData, data])

        # Close the connection
        s.close()




        save_file_location = '%s.png'%(str(name))
        saveFile = open(r'%s.png'%(str(name)), "wb")

        saveFile.write(imgData)
        saveFile.close()

    def single(self):
        a = 0
        self.scope_classy.write('ACQuire:STOPAfter SEQ')
        self.scope_classy.write('ACQuire:STATE ON')
        a=int(self.scope_classy.query('BUSY?'))
        while a == 1:
            a=int(self.scope_classy.query('BUSY?'))
            time.sleep(0.5)
            print ("BUSY__")


    def set_hor(self, hor):
        self.scope_classy.write('HORizontal:MAIn:SCAle %.20f' %(hor))
    def get_freq(self, ch):
        self.scope_classy.write('MEASUrement:IMMed:SOUrce CH%d' %(ch))
        self.scope_classy.write('MEASUrement:IMMed:TYPe FREQ')
        return float(self.scope_classy.query('MEASUrement:IMMed:VAL?'))
    def get_amp(self, ch):
        self.scope_classy.write('MEASUrement:IMMed:SOUrce CH%d' %(ch))
        self.scope_classy.write('MEASUrement:IMMed:TYPe AMP')
        return float(self.scope_classy.query('MEASUrement:IMMed:VAL?'))
    def get_rms(self, ch):
        self.scope_classy.write('MEASUrement:IMMed:SOUrce CH%d' %(ch))
        self.scope_classy.write('MEASUrement:IMMed:TYPe RMS')
        return float(self.scope_classy.query('MEASUrement:IMMed:VAL?'))
    def get_max(self, ch):
        self.scope_classy.write('MEASUrement:IMMed:SOUrce CH%d' %(ch))
        self.scope_classy.write('MEASUrement:IMMed:TYPe MAX')
        return float(self.scope_classy.query('MEASUrement:IMMed:VAL?'))
    def get_min(self, ch):
        self.scope_classy.write('MEASUrement:IMMed:SOUrce CH%d' %(ch))
        self.scope_classy.write('MEASUrement:IMMed:TYPe MINI')
        return float(self.scope_classy.query('MEASUrement:IMMed:VAL?'))
    def get_peak2peak(self, ch):
        self.scope_classy.write('MEASUrement:IMMed:SOUrce CH%d' %(ch))
        self.scope_classy.write('MEASUrement:IMMed:TYPe PK2')
        return float(self.scope_classy.query('MEASUrement:IMMed:VAL?'))
    def get_rise(self, ch):
        self.scope_classy.write('MEASUrement:IMMed:SOUrce CH%d' %(ch))
        self.scope_classy.write('MEASUrement:IMMed:TYPe RIS')
        return float(self.scope_classy.query('MEASUrement:IMMed:VAL?'))


    def set_vertical(self, ch):
        a = float(self.scope_classy.query("CH%d:SCAl?" %(ch)))
        s = float(self.get_peak2peak(ch))
        max_peak = self.get_max(ch)
        min_peak = self.get_min(ch)
        count = 0
        self.scope_classy.write("CH%d:POS %f" %(ch, 0))
        div_correct = 0    
        while a*7.8<s or a*3.5>s or max_peak+div_correct > a*4 or min_peak-div_correct < a*4:
            old_scale = float(self.scope_classy.query("CH%i:SCAl?" %(ch)))
            if a*7.8<s:
                self.scope_classy.write("CH%d:SCAl %f"%(ch, a*2))
            if a*3.5>s:
                self.scope_classy.write("CH%i:SCAl %f"%(ch, a/2))
            if a*7.8>s and a*3.5<s:
                max_peak = self.get_max(ch)
                min_peak = self.get_min(ch)
                position = abs(abs(min_peak)-abs(max_peak))
                div_correct = position/(a*2)
                if abs(max_peak)>abs(min_peak):
                    div_correct = div_correct*-1
                    self.scope_classy.write("CH%d:POS %f" %(ch, div_correct))
                if abs(max_peak)<abs(min_peak):                    
                    self.scope_classy.write("CH%d:POS %f" %(ch, div_correct))
            self.single()
            a = float(self.scope_classy.query("CH%d:SCAl?" %(ch)))
            s = float(self.get_peak2peak(ch))
            if count == 30:
                break
                input("to many try")
            if old_scale == a:
                break
                input("tha_same_scale")
        per_div = (s*1.2)/8
        self.scope_classy.write("CH%d:SCAl %f"%(ch, per_div))

        
    def advanced_set_vertical(scope, ch):
        self.scope_classy.write("TRIG:A:EDG:SOU EXT")
        self.scope_classy.write("TRIG:A:LEV %f" %(1))
        self.scope_classy.write("ACQ:MOD ENV")
        self.scope_classy.write("ACQ:NUMENV 8")
        self.single()
        self.set_vertical(ch)
        trigger_level = self.get_max(scope, ch)
        trigger_level = trigger_level-trigger_level*0.1

        self.scope_classy.write("TRIG:A:EDG:SOU CH%d" %(ch))
        self.scope_classy.write("TRIG:A:LEV %f" %(trigger_level))
        self.scope_classy.write("ACQ:MOD AVE")
        self.scope_classy.write("ACQ:NUMAV 4")



class scope_agilent():
    def __init__(self, scope_adress):
        rm = pyvisa.ResourceManager()
        self.scope_agilent = rm.open_resource('%s' %(scope_agilent))

    def check_instrument_errors(command):
        while True:
            error_string = self.scope_agilent.query(":SYSTem:ERRor?;")
            if error_string: # If there is an error string value.
                if error_string.find("+0,", 0, 3) == -1: # Not "No error".
                    print("ERROR: %s, command: '%s'" % (error_string, command))
                    print("Exited because of error.")
                    sys.exit(1)
                else: # "No error"
                    break
            else: # :SYSTem:ERRor? should always return string.
                print("ERROR: :SYSTem:ERRor? returned nothing, cmmand: '%s'" % command)
                print("Exited because of error.")
                sys.exit(1)
        
    def get_csv(self, ch, name):

        self.scope_agilent.write(":WAVeform:POINts:MODE RAW")
        self.scope_agilent.write(":WAVeform:POINts 10240")
        if ch=="MATH":
            self.scope_agilent.write(":WAVeform:SOURce %s" %(ch))
        else:
            self.scope_agilent.write(":WAVeform:SOURce CHANnel%d" %(ch))
        self.scope_agilent.write(":WAVeform:FORMat BYTE")

        x_increment = float(self.scope_agilent.query(":WAVeform:XINCrement?"))
        x_origin = float(self.scope_agilent.query(":WAVeform:XORigin?"))
        y_increment = float(self.scope_agilent.query(":WAVeform:YINCrement?"))
        y_origin = float(self.scope_agilent.query(":WAVeform:YORigin?"))
        y_reference = float(self.scope_agilent.query(":WAVeform:YREFerence?"))

        sData = self.scope_agilent.query_binary_values(":WAVeform:DATA?", datatype='s')
        values = unpack("%dB" % len(sData[0]), sData[0])
        f = open("%s.csv" %(name), "w")
        for i in xrange(0, len(values) - 1):
            time_val = x_origin + (i * x_increment)
            voltage = ((values[i] - y_reference) * y_increment) + y_origin
            f.write("%E, %f\n" % (time_val, voltage))
        f.close()



    def get_img(self, name):
        self.scope_agilent.timeout = 15000
        self.scope_agilent.write(":HARDcopy:INKSaver OFF;")
        sDisplay = self.scope_agilent.query_binary_values(":DISPLAY:DATA? BMP, SCREEN, COLOR;", datatype='s')
        f = open("%s.png" %name, "wb")
        f.write(sDisplay[0])
        f.close()


    def get_img_csv(self, name, ch):
        self.get_img(name)
        self.get_csv(ch, name)


    def single(self, meas_type="NORM"):
        if meas_type != "NORM":
            meas_type = "PEAK"
            meas_type = "HRES"
            meas_type = "AVER"
        a = 0
        self.scope_agilent.write('ACQuire:TYPE %s' %(meas_type))
        if meas_type == "AVER":
            self.scope_agilent.write("ACQuire:COUNt 16")
        self.scope_agilent.write('ACQuire:STATE ON')
        self.scope_agilent.write("SING;")
        a=int(self.scope_agilent.query('*OPC?'))
        while a == 0:
            a=int(self.scope_agilent.query('*OPC?'))
            time.sleep(0.5)
            print ("BUSY__")    


    def set_hor(self, hor):
        self.scope_agilent.write('TIMebase:SCALe %.20f' %(hor))

        
    def get_freq(self, ch):
        self.scope_agilent.write('MEASure:SOURce CHAN%d' %(ch))
        self.scope_agilent.write('MEASure:FREQ')
        return float(self.scope_agilent.query('MEASure:FREQuency?'))
    
    def get_amp(self, ch):
        self.scope_agilent.write('MEASure:SOURce CHAN%d' %(ch))
        self.scope_agilent.write(':MEASure:VAMPlitude;')
        return float(self.scope_agilent.query(':MEASure:VAMPlitude?'))
    def get_rms(self, ch):
        self.scope_agilent.write('MEASure:SOURce CHAN%d' %(ch))
        self.scope_agilent.write(':MEASure:VRMS;')
        return float(self.scope_agilent.query(':MEASure:VRMS?'))
    def get_max(self, ch):
        self.scope_agilent.write('MEASure:SOURce CHAN%d' %(ch))
        self.scope_agilent.write(':MEASURE:VMAX;')
        return float(self.scope_agilent.query(':MEASURE:VMAX?'))
    def get_min(self, ch):
        self.scope_agilent.write('MEASure:SOURce CHAN%d' %(ch))
        self.scope_agilent.write(':MEASURE:VMIN;')
        return float(self.scope_agilent.query(':MEASURE:VMIN?'))
    def get_peak2peak(self, ch):
        self.scope_agilent.write('MEASure:SOURce CHAN%d' %(ch))
        self.scope_agilent.write(':MEASURE:VPP;')
        return float(self.scope_agilent.query(':MEASURE:VPP?'))
    def get_rise(self, ch):
        self.scope_agilent.write('MEASUrement:IMMed:SOUrce CH%d' %(ch))
        self.scope_agilent.write('MEASUrement:IMMed:TYPe RIS')
        return float(self.scope_agilent.query('MEASUrement:IMMed:VAL?'))

    def fftanalyze(self, ch, freq, span):
        #   :FUNCtion:CENTer <frequency>
        #   :FUNCtion:DISPlay? 0|1
        #   :FUNCtion:GOFT:SOURce1
        #   :FUNCtion:OPERation FFT
        #   :FUNCtion:WINDow HANN
        #   :FUNCtion:SPAN
        
        self.scope_agilent.write(":FUNCtion:SOURce1 CHAN%d" %(ch))
        self.scope_agilent.write(":FUNCtion:OPERation FFT")
        self.scope_agilent.write(":FUNCtion:CENTer %f" %freq)
        self.scope_agilent.write(":FUNCtion:SPAN %f" %(span))
        self.scope_agilent.write(":FUNCtion:DISPlay 1")
        self.single("AVER")
        name = "FFT_%f"
        self.get_img_csv(name, "MATH")


    


    def set_vertical(self, ch):
        a = float(self.scope_classy.query(":CHANnel%d:SCALe?" %(ch)))
        s = float(self.get_peak2peak(ch))
        
        max_peak = self.get_max(ch)
        min_peak = self.get_min(ch)

        count = 0

        self.scope_classy.write(":CHANnel%d:OFFSet %f" %(ch, 0))
        div_correct = 0    
        while a*7.8<s or a*3.5>s or max_peak+div_correct > a*4 or min_peak-div_correct < a*4:
            old_scale = float(self.scope_classy.query(":CHANnel%d:SCALe?" %(ch)))


            if a*7>s and a*3<s:
                max_peak = self.get_max(ch)
                min_peak = self.get_min(ch)
                position = abs(abs(min_peak)-abs(max_peak))
                div_correct = position/(a*2)
                if abs(max_peak)>abs(min_peak):
                    div_correct = div_correct*-1
                    self.scope_classy.write(":CHANnel%d:OFFSet %f" %(ch, div_correct))
                if abs(max_peak)<abs(min_peak):                    
                    self.scope_classy.write(":CHANnel%d:OFFSet %f" %(ch, div_correct))
                    
            if a*7<s:
                self.scope_classy.write(":CHANnel%d:SCALe %f"%(ch, s/6.5))
            if a*3>s:
                self.scope_classy.write(":CHANnel%d:SCALe %f"%(ch, s/3.5))
                            

            self.single()
            a = float(self.scope_classy.query(":CHANnel%d:SCALe?" %(ch)))
            s = float(self.get_peak2peak(ch))







            if count == 30:
                break
                input("too many try")
            if old_scale == a:
                break
                input("same scale")
        per_div = (s*1.2)/8
        self.scope_classy.write(":CHANnel%d:SCALe %f"%(ch, per_div))



        
    def advanced_set_vertical(scope, ch):
        self.scope_classy.write("TRIGger:EDGE:SOURce EXT")
        
        self.scope_classy.write("TRIG:EDGE:LEV %f" %(1))
        self.scope_classy.write("ACQ:MOD ENV")
        self.scope_classy.write("ACQ:NUMENV 8")
        self.single()
        self.set_vertical(ch)
        trigger_level = self.get_max(scope, ch)
        trigger_level = trigger_level-trigger_level*0.1

        self.scope_classy.write("TRIG:A:EDG:SOU CH%d" %(ch))
        self.scope_classy.write("TRIG:A:LEV %f" %(trigger_level))
        self.scope_classy.write("ACQ:MOD AVE")
        self.scope_classy.write("ACQ:NUMAV 4")







def constant_ampl(ampl_point, psg_freq, psg_ampl):
    psg_test.set_psg(psg_freq, psg_ampl, "VPP")
    scope_test.set_hor((1.0/psg_freq)/4.0)
    scope_test.single()
    scope_test.set_vertical(1)
    scope_val=float(scope_test.get_peak2peak(1))
    psg_val=float(psg_test.get_ampl())
    while ampl_point > scope_val*1.2 or ampl_point < scope_val*0.8:
        psg_ampl = psg_val/((scope_val)/ampl_point)
#        raw_input(psg_ampl)
        psg_test.set_psg(psg_freq, psg_ampl, "VPP")
        scope_test.single()
        scope_test.set_vertical(1)
        scope_val=float(scope_test.get_peak2peak(1))
    scope_test.get_img(str(psg_freq))
    psg_freq=psg_freq*1.1
    constant_ampl(ampl_point, psg_freq, psg_ampl)        
        

def load_freq():
    return pd.read_excel('frequency.xlsx')

def save_dataframe(dictionary, excel_name, result_order):
    dataframe = pd.DataFrame(dictionary)
    if result_order != [0]:
        dataframe = dataframe[result_order]

    writer = pd.ExcelWriter('%s.xlsx' %(excel_name), engine='xlsxwriter')
    dataframe.to_excel(writer, sheet_name='RESULT')
    writer.save()
    



