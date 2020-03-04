import new_python_studio

rm = new_python_studio.pyvisa.ResourceManager()

psa_address = "TCPIP0::10.204.20.88::inst0::INSTR"
psa = new_python_studio.psa_meas(psa_address)

scope_address = "TCPIP0::10.204.20.96::inst0::INSTR"
scope = new_python_studio.scope_meas(scope_address)

psg_address="ASRL25::INSTR"
psg = new_python_studio.psg_control(psg_address)

freq=new_python_studio.load_freq()



result_order = ["freq", "coil_current", "psg_ampl", "BAND", "LSB", "USB"]
dictionary = {"freq":[], "coil_current":[], "psg_ampl":[], "USB":[], "LSB":[], "BAND":[]}
band=435000000


def constant_current(target, ch, ampl_gain, freq_center, units_psg):
    meas_current=0
    while meas_current>target*1.02 or meas_current<target*0.98:
        scope.set_vertical(ch)
        meas_current=scope.get_rms(ch)
        print("target is %fmA" %(target*1000))
        print("measured value is %fmA"%(meas_current*1000))
        to_add=(target)/meas_current
        psg_ampl=psg.get_ampl()
        print("PSG ampl is %fVpp"%(psg_ampl))
        to_set=psg_ampl*to_add
        if to_set > 5:
            input("to high input")
        print("new PSG ampl will be: %fVpp" %(to_set))
        print(str(freq_center)+"Hz")
        psg.set_psg(freq_center, to_set, units_psg)

def constant_dbuv(target, ampl_gain, freq_center, units_psg):
    psg.set_psg(freq_center, psg.get_ampl(), units_psg)
    psa.single()
    psa.peak_search()
    marker=psa.get_marker()
    meas_dbuv=float(marker["ampl"][0])
    while meas_dbuv>target*1.002 or meas_dbuv<target*0.998:
        psa.single()
        psa.peak_search()
        marker=psa.get_marker()
        meas_dbuv=float(marker["ampl"][0])
        print("target is %fdBuV" %(target))
        print("measured value is %fdBuV"%(meas_dbuv))
        to_add=(10**(target/20))/(10**(meas_dbuv/20))
        psg_ampl=psg.get_ampl()
        print("PSG ampl is %fVpp"%(psg_ampl))
        to_set=psg_ampl*to_add
        print("new PSG ampl will be: %fVpp" %(to_set))
        if to_set>5:
            input("to_HIGH_input")
        print(freq_center)
        psg.set_psg(freq_center, to_set, units_psg)

def rlevel_control():
#    psa.single()
    psa.peak_search()
    field = psa.get_marker()
    rlev=float(field["ampl"][0])+10.0
    psa.psa_classy.write("DISP:WIND:TRAC:Y:RLEV %f" %(rlev))


psg.set_psg(50, 0.2, "VPP")
for x in range(len(freq)):
    freq_start=freq["freq_start"][x]
    freq_stop=freq["freq_stop"][x]
    freq_center=int((freq_start+freq_stop)/2)
    rlev=freq["rlev"][x]
    units=freq["units"][x]
    points=freq["points"][x]
    rbw=freq["rbw"][x]
    att=freq["att"][x]
    current_target=freq["current_target"][x]
    dbuv_target=freq["dBuV_target"][x]
    units_psg=freq["units_psg"][x]
    usb_f=band+freq_center
    lsb_f=band-freq_center



    psg.rf_on(1)

    psg.set_psg(freq_center, psg.get_ampl(), units_psg)
    scope.set_hor(((1/freq_center)*10)/10)


#    psa.set_x_y_att(freq_start, freq_stop, float(rlev), str(units), points, rbw, att)

    ##curent control
    constant_current(current_target, 3, 30, freq_center, units_psg)


    ##dBuV control
#    constant_dbuv(dbuv_target, 30, freq_center, units_psg)  
#    scope.set_vertical(3)
#    scope.set_vertical(1)


#    coil_voltage=scope.get_rms(1)
    coil_current=scope.get_rms(3)           #measuring on CH3
    scope.get_img(str(freq_center))

#band meas    
    psa.set_x_y_att(band-50, band+50, float(rlev), "dBm", points, rbw, att)
    psa.single()
#    rlevel_control()
#    psa.single()
    psa.psa_classy.write("CALC:MARK1:X %d" %band)

    BAND = psa.get_marker()
    psa.save("band_"+str(freq_center))
#lsb meas    
    psa.set_x_y_att(lsb_f-50, lsb_f+50, float(rlev), "dBm", points, rbw, att)
    psa.single()
    rlevel_control()
#    psa.single()
#    psa.peak_search()
    psa.psa_classy.write("CALC:MARK1:X %d" %lsb_f)
    LSB = psa.get_marker()
    psa.save("lsb_"+str(freq_center))
#usb meas    
    psa.set_x_y_att(usb_f-50, usb_f+50, float(rlev), "dBm", points, rbw, att)
    psa.single()
    rlevel_control()
    psa.psa_classy.write("CALC:MARK1:X %d" %usb_f)

#    psa.single()
#    psa.peak_search()
    USB = psa.get_marker()
    psa.save("usb_"+str(freq_center))









    psg.rf_on(0)

    dictionary["freq"].append(freq_center)
#    dictionary["field"].append(float(field["ampl"][0]))
    dictionary["coil_current"].append(coil_current)
#    dictionary["coil_voltage"].append(coil_voltage)
    dictionary["psg_ampl"].append(float(psg.get_ampl()))
    dictionary["USB"].append(float(USB["ampl"][0]))
    dictionary["LSB"].append(float(LSB["ampl"][0]))
    dictionary["BAND"].append(float(BAND["ampl"][0]))
new_python_studio.save_dataframe(dictionary, "FACE_1_1", result_order)
