import os
import time
import pyvisa
import csv


class DMM34461A(object):

    def __init__(self):
        """初始化资源管理器但不连接设备"""
        self.Meas = 0
        self.rm = pyvisa.ResourceManager()
        self.K34461A = None  # 设备连接句柄

    def connect(self, device_address):
        """连接指定设备"""
        self.K34461A = self.rm.open_resource(device_address)
        # 这里可以添加设备初始化配置

    def conf_curr_dc(self):
        #   设置为DC 电流档
        self.K34461A.write('CONF:CURR:DC')
        device.local()

    def conf_curr_ac(self):
        #   设置为AC电流档
        self.K34461A.write('CONF:CURR:AC')
        device.local()

    def conf_volt_dc(self):
        #   设置为DC电压档
        self.K34461A.write('CONF:VOLT:DC')
        device.local()

    def conf_volt_ac(self):
        #   设置为AC电压档
        self.K34461A.write('CONF:VOLT:AC')
        device.local()

    def set_volt_aperture(self, APER):
        #   设置电压测量积分时间
        if APER == 1:
            self.K34461A.write('VOLTage:APERture 2E-02')   #1 PLC
        elif APER == 10:
            self.K34461A.write('VOLTage:APERture 2E-01')   #10 PLC
        elif APER == 100:
            self.K34461A.write('VOLTage:APERture 2E+00')    #100 PLC
        elif APER == 0.2:
            self.K34461A.write('VOLTage:APERture 3E-03')    #0.2 PLC
        elif APER == 0.02:
            self.K34461A.write('VOLTage:APERture 3E-04')    #0.02 PLC
        else:
            print('输入参数错误，请输入0.02/0.2/1/10/100')

        device.local()

    def set_input_Z(self, IMMP):
        if IMMP == '10M':
            self.K34461A.write('VOLT:DC:IMPedance:AUTO 0')
        elif IMMP == 'AUTO':
            self.K34461A.write('VOLT:DC:IMPedance:AUTO 1')
        else:
            print('输入参数错误，请输入"10M"/"AUTO"')

        device.local()

    def get_volt_dc(self):
        #   获取DC电压档电压值
        volt = self.K34461A.query('MEAS:VOLT:DC?')
        # device.local()
        return round(float(volt),6)


    def get_volt_ac(self):
        #   获取AC电压档电压值
        volt = self.K34461A.query('MEAS:VOLT:AC?')
        # device.local()
        return round(float(volt),6)

    def get_curr_dc(self):
        #   获取DC电流档电流值
        curr = self.K34461A.query('MEAS:CURR:DC?')
        # device.local()
        return round(float(curr),6)

    def get_curr_ac(self):
        #   获取AC电流档电流值
        curr = self.K34461A.query('MEAS:CURR:AC?')
        # device.local()
        return round(float(curr),6)

    def get_immp(self):
        #   获取电阻值
        res = self.K34461A.query('MEASUre:RESistance?')
        # device.local()
        return round(float(res), 6)

    def measurement(self, function):
        if function == 'DCV':
            self.Meas = self.K34461A.query('MEASUre:VOLTage:DC?')
        if function == 'DCI':
            self.Meas = self.K34461A.query('MEASUre:CURRent:DC?')
        if function == 'ACV':
            self.Meas = self.K34461A.query('MEASUre:VOLTage:AC?')
        if function == 'ACI':
            self.Meas = self.K34461A.query('MEASUre:CURRent:AC?')
        if function == 'Res':
            self.Meas = self.K34461A.query('MEASUre:RESistance?')
        Meas_Result = float(self.Meas)
        # device.local()
        return Meas_Result

    def local(self):
        self.K34461A.write('SYST:LOC')

    def text_function(self, cmd):
        if '?' in cmd:
            read = self.K34461A.query(cmd)
            return read
        else:
            self.K34461A.write(cmd)
        # device.local()

    def start_scanning(self, interval_time, meas_param):

        print("start_scanning")
        while True:
            if meas_param == '直流电压':
                print(self.get_volt_dc)
            elif meas_param == '交流电压':
                print(self.get_volt_dc)
            elif meas_param == '直流电流':
                print(self.get_volt_dc)
            elif meas_param == '交流电流':
                print(self.get_volt_dc)
            elif meas_param == '电阻':
                print(self.get_volt_dc)
            else:
                pass
            time.sleep(float(interval_time))




    def stop_scanning(self):
        print("stop_scanning")


