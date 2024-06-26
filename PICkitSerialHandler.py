# !/usr/bin/env python
# -*- coding:utf-8 -*-
"""
    使用pythonnet建立與PICkitS.dll溝通
    * 使用.net dll，裡面的參數要用c#的偽參數，在import clr的時候就已有import C# runtime的相
    關type, 因此參數的建立用c# runtime的type來傳入

    TODO:
        PMBUSWriteWithPEC

"""
import enum
import threading
import time
import logging
import typing

# log 設定
LOGLEVEL = logging.WARNING  # CRITICAL > ERROR > WARNING > INFO > DEBUG > NOTSET
formatter = "%(asctime)s.%(msecs)03d %(filename)s[line:%(lineno)4d] %(levelname)s %(message)s"
# logging.basicConfig(level=LOGTYPE,
#                     format='%(asctime)s.%(msecs)03d %(filename)s[line:%(lineno)4d] %(levelname)s %(message)s',
#                     datefmt='%Y.%m.%d %H:%M:%S')
_PKSA_LOG = logging.getLogger(__name__)
_PKSA_LOG.setLevel(LOGLEVEL)
handler = logging.StreamHandler()
handler.setLevel(LOGLEVEL)
handler_formatter = logging.Formatter(formatter)
if LOGLEVEL == logging.DEBUG:
    handler.setFormatter(handler_formatter)
_PKSA_LOG.addHandler(handler)

""" 
    imprt clr: 導入PICKitS.dll的module
    ref: https://github.com/pythonnet/pythonnet#pythonnet---python-for-net
    ref: http://pythonnet.github.io/
    
    * clr是什麼？
    Common Language Runtime（通用语言运行时）是用来管理任意支持的语言编写的程序执行、
    允许他们分享用任意语言编写的通用的面向对象的类的程序。
    簡單來說就是一個界面，任何程式語言都可以自由的存取
    ref: https://blog.csdn.net/lxjames833539/article/details/6453544
    
"""
import sys
import os
path = os.getcwd()

# === 還沒測試好的code（at linux platform）===
if sys.platform == "linux":
    sys.path.append("{}/libs/PICkitS.dll".format(path))
    import pythonnet
    import clr
    from System import Reflection
    from System import Array

    Reflection.Assembly.LoadFile("{}/libs/PICkitS.dll".format(path))
    from PICkitS import *
# === 還沒測試好的code（at linux platform）End ===
else:
    import clr
    from clr import System
    from System import Reflection

    # clr.AddReference("PICkitS")
    # 如果clr的提供的c# function還不夠，可直接import c#的.net api
    # ref: https://docs.microsoft.com/zh-tw/dotnet/api/system?view=netstandard-1.6
    # clr.AddReference("System.Runtime")
    if __name__ == "__main__":
        Reflection.Assembly.LoadFile("{}\\PICkitS.dll".format(path))
    else:
        path = os.path.dirname(os.path.abspath(__file__))
        Reflection.Assembly.LoadFile("{}\\PICkitS.dll".format(path))
    from PICkitS import *
    from System import *
    from System import Array

_PKSA_LOG.debug("pythonnet version: {}".format(clr.__version__))


class ErrorCode(enum.Enum):
    PKSA_Exec_err = (0, "PICkit Serial Analyzer used dll control error, use [clear_comm_errors] to "
                        "clear PICkit' error.")

    def __init__(self, code, msg):
        self.code = code
        self.msg = msg


class PICkitS():
    def __init__(self):
        super(PICkitS, self).__init__()
        self.is_connect = False  # is connected to PKSA

    def Device_initialize_PICkitSerial(self) -> bool:
        """
        prototype: bool = Device.Initialize_PICkitSerial()

        Attempts to establish communication with PICkit Serial
        device and initialize communication threads used by
        class library.  If multiple devices are attached to
        host PC, function will only initialize first device
        it finds.

        prototype: bool Initialize_PICkitSerial(ushort USBIndex)

        Attempts to establish communication with PICkit Serial
        device and initialize communication threads used by
        class library.

        :return: bool - whether function is executed success
        """
        status = Device.Initialize_PICkitSerial()
        _PKSA_LOG.debug("device_initialize_PICkitSerial: {}".format(status))
        self.is_connect = status
        return status

    @staticmethod
    def how_many_PICkitSerials_are_Attached() -> int:
        """
        prototype: ushort How_Many_PICkitSerials_Are_Attached()
        Polls USB devices for PKSA productID and vendorID looks for a maximum of 30 devices
        :return: int - How many PKSA device
        """
        number = Device.How_Many_PICkitSerials_Are_Attached()
        _PKSA_LOG.debug("Find device number: {}".format(number))
        return number

    def Device_flash_LED1_for_2_seconds(self) -> bool:
        """
            可以讓PICkit亮燈 2 秒
        :return: bool, whether function is executed success
        """
        if not Device.How_Many_PICkitSerials_Are_Attached():
            return False
        Device.Flash_LED1_For_2_Seconds()
        time.sleep(2.1)
        return True

    def set_I2C_bit_rate(self, bit_rate: int) -> bool:
        """
        prototype: bool Set_I2C_Bit_Rate(double p_Bit_Rate)
        :param bit_rate: changes byte 23 to coded value using
                         l_calc_baud = (int)System.Math.Round((double)(CONST.FOSC *1000) / (double)(Baud) / 4.0) - 1;
        :return: bool, true if baud rate was successfully changed, false if not
        """
        s = I2CM.Set_I2C_Bit_Rate(bit_rate)
        _PKSA_LOG.debug("set_I2C_bit_rate: {}".format(s))
        return s

    def get_I2C_bit_rate(self) -> float:
        """
        prototype: double Get_I2C_Bit_Rate()
        Read Status Packet from PKSA and calculates baud rate from byte 19 of the status block
        :return: float, Bit rate of PKSA in kHz, returns zero if an error occured
        """
        return I2CM.Get_I2C_Bit_Rate()

    def PMBus_write(self, address: int, cmd: Array, length: int) -> (bool, str):
        """
        PMBus的write
        prototype:
            bool = Write(byte p_slave_addr,
                         byte p_start_data_addr,
                         byte p_num_bytes_to_write,
                         ref byte[] p_data_array,
                         ref string p_script_view)
        :param address int of slave address, if slave address is 0x58, the value is 0xB0
        :param cmd the array includes start address and write data
        :param length int of write data length, must equal to write data length
        :return: bool, whether function is executed success
        :return: string, reference to a string to which will be copied a formatted view of the command
        """
        if not self.is_connect:
            return False, "No connect."

        p_slave_addr = Byte(address)
        p_start_data_addr = Byte(cmd[0])
        p_num_bytes_to_write = Byte(length)
        p_data_array = Array[Byte](cmd[1:])
        p_script_view = String("")
        result = I2CM.Write(p_slave_addr, p_start_data_addr, p_num_bytes_to_write, p_data_array, p_script_view)
        _PKSA_LOG.debug("PMBus write result: {}".format(result))
        return result[0], "{}".format(result[2])

    def I2C_receive(self, slave_addr: int, read_data_bytes: int) -> (bool, typing.List[bytes], str):
        """
        I2C receive用
        prototype:
                bool = Receive(byte p_slave_addr,
                               byte p_num_bytes_to_read,
                               ref byte[] p_data_array,
                               ref string p_script_view)
        :return: bool, whether function executed success.
        :return: typing.List[bytes], reference to byte array that will store retrieved data
        :return: str, reference to a string to which will be copied a formatted view of the command
        """
        if not self.is_connect:
            return False, [], "No connect to PKSA."

        p_slave_addr = Byte(slave_addr)
        p_num_bytes_to_read = Byte(read_data_bytes)
        p_data_array = Array[Byte]([0] * read_data_bytes)
        p_script_view = String("")
        result = I2CM.Receive(p_slave_addr, p_num_bytes_to_read, p_data_array, p_script_view)
        if not result[0]:
            return result[0], [], "{}".format(result[2])
        return_data = []
        for i in result[1]:
            return_data.append(i)
        return result[0], return_data, "{}".format(result[2])

    def PMBus_read(self, slave_addr: int, start_data_addr: int, read_data_bytes: int) -> (bool, typing.List[bytes]):
        """
        PMBus的read
        prototype:
                bool = Read(byte p_slave_addr,
                            byte p_start_data_addr,
                            byte p_num_bytes_to_read,
                            ref byte[] p_data_array,
                            ref string p_script_view)
        :return: bool, whether function executed success.
                 typing.List[bytes], reference to byte array that will store retrieved data
        """
        if not self.is_connect:
            return False, ["No connect."]

        p_slave_addr = Byte(slave_addr)
        p_start_data_addr = Byte(start_data_addr)
        p_num_bytes_to_read = Byte(read_data_bytes)
        p_data_array = Array[Byte]([0] * read_data_bytes)
        p_script_view = String("")
        result = I2CM.Read(p_slave_addr, p_start_data_addr, p_num_bytes_to_read, p_data_array, p_script_view)
        _PKSA_LOG.debug("I2C Master Read function return: {}".format(result))
        if not result[0]:
            # return result[0], ["{}".format(result[2])]
            return result[0], []
        return_data = []
        for i in result[1]:
            return_data.append(i)
        _PKSA_LOG.debug("I2C Master read data: {}".format(return_data))
        return result[0], return_data

    def special_command_read(self, slave_addr: int,
                             command1: int,
                             command2: int,
                             read_data_bytes: int) -> (bool, list):
        """
        處理特別的command
        prototype:
             public static bool Read(byte p_slave_addr,
                            byte p_command1,
                            byte p_command2,
                            byte p_num_bytes_to_read,
                            ref byte[] p_data_array,
                            ref string p_script_view)
        :return: bool, list
        """
        if not self.is_connect:
            return False, ["No connect."]
        p_slave_addr = Byte(slave_addr)
        p_command1 = Byte(command1)
        p_command2 = Byte(command2)
        p_num_bytes_to_read = Byte(read_data_bytes)
        p_data_array = Array[Byte]([0] * read_data_bytes)
        p_script_view = String("")
        result = I2CM.Read(p_slave_addr, p_command1, p_command2,
                           p_num_bytes_to_read, p_data_array, p_script_view)
        _PKSA_LOG.debug("I2C Master Read function return: {}".format(result))
        if not result[0]:
            return result[0], ["{}".format(result[2])]
        return_data = []
        for i in result[1]:
            return_data.append(i)
        _PKSA_LOG.debug("I2C Master read data: {}".format(return_data))
        return result[0], return_data

    def Configure_PICkitSerial_for_I2CMaster(self):
        # status = I2CM.Configure_PICkitSerial_For_I2CMaster(False,
        #                                                    False,
        #                                                    False,
        #                                                    False,
        #                                                    True,
        #                                                    3.3)
        status = I2CM.Configure_PICkitSerial_For_I2CMaster()
        _PKSA_LOG.debug("configure_PICkitSerial_for_I2CMaster status: {}".format(status))
        return status

    def set_read_wait_time(self, p_time):
        I2CM.Set_Read_Wait_Time(p_time)

    def get_read_wait_time(self):
        return I2CM.Get_Read_Wait_Time()

    def get_PICKitS_DLL_version(self):
        """
            dll版本
            prototype:
                Get_PickitS_DLL_Version(ref string p_version)
        :return: (bool, string)
        """
        version = ""
        status, version = Device.Get_PickitS_DLL_Version(version)
        _PKSA_LOG.debug("version :{}".format(version))
        return status, version

    def get_script_timeout(self) -> int:
        """
            取得DLL 執行timeout
        :return: ms
        """
        return Device.Get_Script_Timeout()

    def set_script_timeout(self, p_time: int):
        """
            設定DLL 執行timeout
        :param p_time:
        :return:
        """
        Device.Set_Script_Timeout(p_time)

    def cleanup(self):
        Device.Terminate_Comm_Threads()
        Device.Cleanup()
        self.is_connect = False

    def clear_comm_errors(self):
        Device.Clear_Comm_Errors()

    def reestablish_comm_thread(self):
        Device.Terminate_Comm_Threads()
        Device.ReEstablish_Comm_Threads()
        return True

    def get_source_voltage(self) -> (bool, float, bool):
        """
            取得I2C bus設定的主動電壓
             prototype:
                Get_Source_Voltage(ref double p_voltage, ref bool p_PKSA_power)
        :return: 執行成功與否, 設定的電壓點, 是否設定主動電壓
        """
        p_voltage = 0.0
        p_pksa_power = False
        s, v, is_power = I2CM.Get_Source_Voltage(p_voltage, p_pksa_power)
        return s, v, is_power

    def set_source_voltage(self, v: float) -> bool:
        """
            設定I2C bus的主動電壓
            prototype:
                Set_Source_Voltage(double p_voltage)
        :param v: float, need to be in range 0.0 ~ 5.0
        :return:
        """
        if not (0 < v < 5):
            return False

        return I2CM.Set_Source_Voltage(v)



    def __del__(self):
        if self.is_connect:
            self.cleanup()


if __name__ == "__main__":
    PICkit = PICkitS()
    print(PICkit.Device_initialize_PICkitSerial())
    print(PICkit.Configure_PICkitSerial_for_I2CMaster())
    print(PICkit.Device_flash_LED1_for_2_seconds())

    # 讀測試 01
    # print(PICkit.PMBus_read(0xB0, 0x55, 2))
    # ============================================

    # 寫測試
    # Page 12V
    # PICkit.PMBus_write(0xB0, [0x00, 0x00], 1)
    # print(PICkit.PMBus_read(0xB0, 00, 1))
    #
    # # Page 5V
    # PICkit.PMBus_write(0xB0, [0x00, 0x10], 1)
    # print(PICkit.PMBus_read(0xB0, 00, 1))
    #
    # print(PICkit.PMBus_write(0xB0, [0x00, 0x00], 1))
    # ============================================

    # 特殊讀測試
    # print(PICkit.special_command_read(0xB0, 0x02, 0x00, 2))

    # 只讀測試
    # print(PICkit.I2C_receive(0xB1, 0x01))

    # get pksa's setting voltage
    # print(PICkit.get_source_voltage())
    # print(PICkit.set_source_voltage(3.3))
    # print(PICkit.get_source_voltage())
