import os
import time

import requests
import serial
import win32api
import win32print
from pyhtml2pdf import converter

port = os.getenv("COM")


class Data:
    def __init__(self, platform=None, product=None):
        self.platform = platform
        self.product = product


ser = serial.Serial()
try:
    ser = serial.Serial(
        # Serial Port to read the data from
        port=port,

        # Rate at which the information is shared to the communication channel
        baudrate=9600,

        # Applying Parity Checking (none in this case)
        parity=serial.PARITY_NONE,

        # Pattern of Bits to be read
        stopbits=serial.STOPBITS_ONE,

        # Total number of bits to be read
        bytesize=serial.EIGHTBITS,

        # Number of serial commands to accept before timing out
        timeout=0.001
    )
    print("Выполнено подключение.")
    data = Data()
    while True:
        try:
            barcode = int(ser.readline().decode().strip())
        except Exception as ex:
            continue
        print(barcode)
        if barcode:
            if barcode == 99999994:
                print('Печать...')
                try:
                    converter.convert(f"http://{os.getenv("IP_AND_PORT")}{os.getenv("PRINT")}", "result.pdf", power=2,
                                      print_options={"scale": float(os.getenv("TEXT_SIZE"))})
                    win32api.ShellExecute(
                        0,
                        "print",
                        'result.pdf',
                        '/d:"%s"' % win32print.GetDefaultPrinter(),
                        ".",
                        0
                    )
                    print('Файл распечатан.')
                except Exception:
                    print(f'Невозможно распечатать: http://{os.getenv("IP_AND_PORT")}{os.getenv("PRINT")}')
                continue
            if barcode <= 215 or barcode == 666:
                data = Data(platform=barcode, product=None)
                print(f'Пирамида: {data.platform}')
            elif data.platform is not None:
                data.product = barcode
                print(f'{data.platform}: {data.product}')
            try:
                response = requests.post(f"http://{os.getenv("IP_AND_PORT")}{os.getenv("data")}", json=data.__dict__)
            except Exception:
                print(f'Ошибка сервера: http://{os.getenv("IP_AND_PORT")}/')
                break
except serial.SerialException as se:
    print("error:", str(se))

except Exception as ex:
    print("error:", str(ex))

finally:
    try:
        ser.close()
        print("Сканер отключен.")
    except Exception as ex:
        print(str(ex))

time.sleep(10)
