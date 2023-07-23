from getpass import getpass
from datetime import datetime
import xlsxwriter
from netmiko import NetMikoAuthenticationException, ConnectHandler

username = input('Enter Username: ')
password = getpass()

metro_cet = {
    'device_type': 'ciena_saos',
    'username': username,
    'password': password,
}

dt = datetime.now()
t_stamp = dt.strftime("%d.%m.%Y_%H.%M.%S")
workbook = xlsxwriter.Workbook(f'<#directory>\\<Report_filename>_{t_stamp}.xlsx')
worksheet_ng = workbook.add_worksheet('<#excel_file_worksheet_name>')

centre_bold = workbook.add_format({      
    'align': 'center',
    'bold': 1,
    'valign': 'rjustify',
    'border': 2,
    'text_wrap': True
})

centre_bold15 = workbook.add_format({      
    'align': 'center',
    'bold': 1,
    'valign': 'rjustify',
    'border': 2,
    'text_wrap': True,
    'font' : 15
})

centre = workbook.add_format({      
    'align': 'center',
    'valign': 'rjustify',
    'text_wrap': True,
    'border': 2
})

class Tx_Lagos_Core():

    def __init__(self):
        self.worksheet_ng = worksheet_ng

    def set_column_width(self):
        self.worksheet_ng.set_column('A:E', 18)
        self.worksheet_ng.set_column('F:BZ', 40)

    def freeze_panes(self):
        self.worksheet_ng.freeze_panes(7, 5)

    def add_legend(self):
        self.worksheet_ng.merge_range('A1:E1', 'Legend', centre_bold15)
        legends = {
            'A2': 'Ports',
            'A3': '10G (10km - 20km)',
            'A4': '10G (40km - 80km)',
            'A5': '100G CFP-2',
            'A6': '100G QSFP',
            'B2': 'Tx Upper Threshold',
            'C2': 'Tx Lower Threshold',
            'D2': 'Rx Upper Threshold',
            'E2': 'Rx Lower Threshold',
            'B3': '+0.49dBm',
            'C3': '-8.19',
            'D3': '+0.49dBm',
            'E3': '-14.4dBm',
            'B4': '+5.0dBm',
            'C4': '0.00',
            'D4': '-7.00dBm',
            'E4': '-23.01dBm',
            'B5': 'N/A',
            'C5': 'N/A',
            'D5': '+6.0dBm',
            'E5': '-25.0dBm',
            'B6': '+3.00dBm',
            'C6': '-9.40',
            'D6': '-2dBm',
            'E6': '-20.0dBm'
        }

        for cell, value in legends.items():
            self.worksheet_ng.write(cell, value, centre_bold)

    def write_header(self):
        headers = [
            ('E8', 'ROUTE'),
            ('E9', 'POP'),
            ('E10', 'Device'),
            ('E11', 'Port Num'),
            ('E12', 'Port Type'),
            ('E13', 'Description'),
            ('E14', 'Port Op Status'),
            ('E15', 'Port Admin Status'),
            ('E16', 'Tx Power'),
            ('E17', 'Rx Power')
        ]

        for cell, header in headers:
            self.worksheet_ng.write(cell, header, centre_bold)

    def write_device_data(self, cell_range, data):
        for i, value in enumerate(data):
            cell = f'{chr(ord(cell_range[0]) + i)}{cell_range[1]}'
            self.worksheet_ng.write(cell, value, centre_bold)

    def connect_to_device(self, host):
        try:
            return ConnectHandler(**host)
        except NetMikoAuthenticationException:
            print(f'{host} could not be reached or connection timed out !!!')

    def write_device_info(self, cell_range, device_info):
        for i, info in enumerate(device_info, start=2):
            cell = f'{chr(ord(cell_range[0]) + i)}{cell_range[1]}'
            self.worksheet_ng.write(cell, info, centre)

    def write_device_sfp_info(self, cell_range, sfp_info):
        for i, (tx, rx) in enumerate(sfp_info, start=
