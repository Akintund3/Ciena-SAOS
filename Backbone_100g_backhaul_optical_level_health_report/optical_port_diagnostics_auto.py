#!/usr/bin/env python

# Libs

from getpass import getpass
from mimetypes import init
from msilib.schema import Class
from typing_extensions import Self
import netmiko
from netmiko import NetMikoAuthenticationException, ConnectHandler, NetMikoTimeoutException, ConnectionException
from datetime import datetime
import xlsxwriter


username = input('Enter Username: ')
password = getpass()


metro_cet = {
'device_type': 'ciena_saos',
'username': username,
'password': password,
}


dt = datetime.now()

t_stamp = dt.strftime("%d.%m.%Y_%H.%M.%S")

workbook = xlsxwriter.Workbook(f'C:\\Users\\akintunde.adeniran\\OneDrive - mainone.net\\Desktop\Auto BH performance\\Mainone_Ciena_CET_Optical_Performance_check_{t_stamp}.xlsx')


worksheet_ng = workbook.add_worksheet('NG Tx Core')

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

    def __init__(self) :
        pass


    #Broadening the cells width
    worksheet_ng.set_column('A:E', 18)
    worksheet_ng.set_column('F:BZ', 40)

    #freeze row 1-7, col A-E
    worksheet_ng.freeze_panes(7, 5)


    #Legends
    worksheet_ng.merge_range('A1:E1', 'Legend', centre_bold15)
    worksheet_ng.write('A2', 'Ports', centre_bold)
    worksheet_ng.write('A3', '10G (10km - 20km)', centre_bold)
    worksheet_ng.write('A4', '10G (40km - 80km)', centre_bold)
    worksheet_ng.write('A5', '100G CFP-2', centre_bold)
    worksheet_ng.write('A6', '100G QSFP', centre_bold)
    worksheet_ng.write('B2', 'Tx Upper Threshold', centre_bold)
    worksheet_ng.write('C2', 'Tx Lower Threshold', centre_bold)
    worksheet_ng.write('D2', 'Rx Upper Threshold', centre_bold)
    worksheet_ng.write('E2', 'Rx Lower Threshold', centre_bold)
    worksheet_ng.write('B3', '+0.49dBm', centre_bold)
    worksheet_ng.write('C3', '-8.19', centre_bold)
    worksheet_ng.write('D3', '+0.49dBm', centre_bold)
    worksheet_ng.write('E3', '-14.4dBm', centre_bold)
    worksheet_ng.write('B4', '+5.0dBm', centre_bold)
    worksheet_ng.write('C4', '0.00', centre_bold)
    worksheet_ng.write('D4', '-7.00dBm', centre_bold)
    worksheet_ng.write('E4', '-23.01dBm', centre_bold)
    worksheet_ng.write('B5', 'N/A', centre_bold)
    worksheet_ng.write('C5', 'N/A', centre_bold)
    worksheet_ng.write('D5', '+6.0dBm', centre_bold)
    worksheet_ng.write('E5', '-25.0dBm', centre_bold)
    worksheet_ng.write('B6', '+3.00dBm', centre_bold)
    worksheet_ng.write('C6', '-9.40', centre_bold)
    worksheet_ng.write('D6', '-2dBm', centre_bold)
    worksheet_ng.write('E6', '-20.0dBm', centre_bold)


    #Generic inputs (Route, POP, devices, port e.t.c....)
    worksheet_ng.write('E8', 'ROUTE', centre_bold)
    worksheet_ng.write('E9', 'POP', centre_bold)
    worksheet_ng.write('E10', 'Device', centre_bold)
    worksheet_ng.write('E11', 'Port Num', centre_bold)
    worksheet_ng.write('E12', 'Port Type', centre_bold)
    worksheet_ng.write('E13', 'Description', centre_bold)
    worksheet_ng.write('E14', 'Port Op Status', centre_bold)
    worksheet_ng.write('E15', 'Port Admin Status', centre_bold)
    worksheet_ng.write('E16', 'Tx Power', centre_bold)
    worksheet_ng.write('E17', 'Rx Power', centre_bold)





    def saka_cls(self):

        hs = ['Saka 5171_01', 'Ajah 5171_01', 'Saka 5171_02', 'Ajah 5171_02' ]

        ne_p = {'172.20.38.249' : '40', 
        '172.20.39.4' : '39', 
        '172.20.38.250' : '40',
        '172.20.39.5' : '39'
        }

        ne = [i for i in ne_p.keys()]
        p = [j for j in ne_p.values()]


        #Route and POP
        worksheet_ng.merge_range('F8:G8', 'SAKA - CLS (Pry: 30km)', centre_bold)
        worksheet_ng.write('F9', 'SAKA POP', centre_bold)
        worksheet_ng.write('G9', 'CLS', centre_bold)

        #Route and POP
        worksheet_ng.merge_range('H8:I8', 'SAKA - CLS (Secondary: 28.2km)', centre_bold)
        worksheet_ng.write('H9', 'SAKA POP', centre_bold)
        worksheet_ng.write('I9', 'CLS', centre_bold)
        

        #Devices
        worksheet_ng.write('F10', hs[0], centre_bold)
        worksheet_ng.write('G10', hs[1], centre_bold)
        worksheet_ng.write('H10', hs[2], centre_bold)
        worksheet_ng.write('I10', hs[3], centre_bold)

        #Ports
        worksheet_ng.write('F11', p[0], centre_bold)
        worksheet_ng.write('G11', p[1], centre_bold)
        worksheet_ng.write('H11', p[2], centre_bold)
        worksheet_ng.write('I11', p[3], centre_bold)


        #sak - cls route
        
        
        #ajah 5171_01, ajah 5171_02, saka 5171_01, saka 5171_02
        
        for i in ne:

                        
            sak1 = {**metro_cet, 'host' : ne[0]}
            cls1 = {**metro_cet, 'host' : ne[1]}
            sak2 = {**metro_cet, 'host' : ne[2]}
            cls2 = {**metro_cet, 'host' : ne[3]}


        try:

            net_ssh = ConnectHandler(**sak1)    
            xc = net_ssh.send_command(f'port xc sh po {p[0]} diag')
            po = net_ssh.send_command(f'port sh po {p[0]}')
            hs = net_ssh.send_command('system show host-name')



            sak1_des = ((po.splitlines()[4]).split('|')[2]).replace(' ', '')
            sak1_op = ((po.splitlines()[6]).split('|')[2]).replace(' ', '')
            sak1_admin = ((po.splitlines()[6]).split('|')[3]).replace(' ', '')
            sak1_type = ((po.splitlines()[5]).split('|')[2]).replace(' ', '')
            sak1_name = ((hs.splitlines()[2]).split('|')[2]).replace(' ', '')
            sak1_tx = ((xc.splitlines()[17]).split('|')[2]).replace(' ', '') + ' dBm'
            sak1_rx = ((xc.splitlines()[23]).split('|')[2]).replace(' ', '') + ' dBm'


            if 'no XCVR present' in xc:

                print(f'SFP not detect, Please check {sak1_name} Port {p[0]}')
        

                worksheet_ng.write('F12', sak1_type, centre)
                worksheet_ng.write('F13', sak1_des, centre)
                worksheet_ng.write('F14', sak1_op, centre)
                worksheet_ng.write('F15', sak1_admin, centre)
                worksheet_ng.write('F16', 'SFP not found', centre)
                worksheet_ng.write('F17', 'SFP not found', centre)


            else :      

                worksheet_ng.write('F12', sak1_type, centre)
                worksheet_ng.write('F13', sak1_des, centre)
                worksheet_ng.write('F14', sak1_op, centre)
                worksheet_ng.write('F15', sak1_admin, centre)
                worksheet_ng.write('F16', sak1_tx, centre)
                worksheet_ng.write('F17', sak1_rx, centre)


        except NetMikoAuthenticationException :

            print(f'{i} could not be reached or connection timed out !!!')


            worksheet_ng.write('F12', 'N/A', centre)
            worksheet_ng.write('F13', 'N/A', centre)
            worksheet_ng.write('F14', 'N/A', centre)
            worksheet_ng.write('F15', 'N/A', centre)
            worksheet_ng.write('F16', 'N/A', centre)
            worksheet_ng.write('F17', 'N/A', centre)


        try:

            net_ssh = ConnectHandler(**cls1)    
            xc = net_ssh.send_command(f'port xc sh po {p[1]} diag')
            po = net_ssh.send_command(f'port sh po {p[1]}')
            hs = net_ssh.send_command('system show host-name')


            cls1_des = ((po.splitlines()[4]).split('|')[2]).replace(' ', '')
            cls1_op = ((po.splitlines()[6]).split('|')[2]).replace(' ', '')
            cls1_admin = ((po.splitlines()[6]).split('|')[3]).replace(' ', '')
            cls1_type = ((po.splitlines()[5]).split('|')[2]).replace(' ', '')
            cls1_name = ((hs.splitlines()[2]).split('|')[2]).replace(' ', '')
            cls1_tx = ((xc.splitlines()[17]).split('|')[2]).replace(' ', '') + ' dBm'
            cls1_rx = ((xc.splitlines()[23]).split('|')[2]).replace(' ', '') + ' dBm'


            if 'no XCVR present' in xc:
                print(f'SFP not detect, Please check {cls1_name} Port {p[1]}')


                worksheet_ng.write('G12', cls1_type, centre)
                worksheet_ng.write('G13', cls1_des, centre)
                worksheet_ng.write('G14', cls1_op, centre)
                worksheet_ng.write('G15', cls1_admin, centre)
                worksheet_ng.write('G16', 'SFP not found', centre)
                worksheet_ng.write('G17', 'SFP not found', centre)



            else :

                worksheet_ng.write('G12', cls1_type, centre)
                worksheet_ng.write('G13', cls1_des, centre)
                worksheet_ng.write('G14', cls1_op, centre)
                worksheet_ng.write('G15', cls1_admin, centre)
                worksheet_ng.write('G16', cls1_tx, centre)
                worksheet_ng.write('G17', cls1_rx, centre)



        except NetMikoAuthenticationException :

            print(f'{i} could not be reached or connection timed out !!!')

            worksheet_ng.write('G12', 'N/A', centre)
            worksheet_ng.write('G13', 'N/A', centre)
            worksheet_ng.write('G14', 'N/A', centre)
            worksheet_ng.write('G15', 'N/A', centre)
            worksheet_ng.write('G16', 'N/A', centre)
            worksheet_ng.write('G17', 'N/A', centre)


        try:

            net_ssh = ConnectHandler(**sak2)    
            xc = net_ssh.send_command(f'port xc sh po {p[2]} diag')
            po = net_ssh.send_command(f'port sh po {p[2]}')
            hs = net_ssh.send_command('system show host-name')


            sak2_des = ((po.splitlines()[4]).split('|')[2]).replace(' ', '')
            sak2_op = ((po.splitlines()[6]).split('|')[2]).replace(' ', '')
            sak2_admin = ((po.splitlines()[6]).split('|')[3]).replace(' ', '')
            sak2_type = ((po.splitlines()[5]).split('|')[2]).replace(' ', '')
            sak2_name = ((hs.splitlines()[2]).split('|')[2]).replace(' ', '')
            sak2_tx = ((xc.splitlines()[17]).split('|')[2]).replace(' ', '') + ' dBm'
            sak2_rx = ((xc.splitlines()[23]).split('|')[2]).replace(' ', '') + ' dBm'


            if 'no XCVR present' in xc:

                print(f'SFP not detect, Please check {sak2_name} Port {p[2]}')


                worksheet_ng.write('H12', sak2_type, centre)
                worksheet_ng.write('H13', sak2_des, centre)
                worksheet_ng.write('H14', sak2_op, centre)
                worksheet_ng.write('H15', sak2_admin, centre)
                worksheet_ng.write('H16', 'SFP not found', centre)
                worksheet_ng.write('H17', 'SFP not found', centre)


            else :

                worksheet_ng.write('H12', sak2_type, centre)
                worksheet_ng.write('H13', sak2_des, centre)
                worksheet_ng.write('H14', sak2_op, centre)
                worksheet_ng.write('H15', sak2_admin, centre)
                worksheet_ng.write('H16', sak2_tx, centre)
                worksheet_ng.write('H17', sak2_rx, centre)


        except NetMikoAuthenticationException :

            print(f'{i} could not be reached or connection timed out !!!')

            worksheet_ng.write('H12', 'N/A', centre)
            worksheet_ng.write('H13', 'N/A', centre)
            worksheet_ng.write('H14', 'N/A', centre)
            worksheet_ng.write('H15', 'N/A', centre)
            worksheet_ng.write('H16', 'N/A', centre)
            worksheet_ng.write('H17', 'N/A', centre)


        try:

            net_ssh = ConnectHandler(**cls2)    
            xc = net_ssh.send_command(f'port xc sh po {p[3]} diag')
            po = net_ssh.send_command(f'port sh po {p[3]}')
            hs = net_ssh.send_command('system show host-name')


            cls2_des = ((po.splitlines()[4]).split('|')[2]).replace(' ', '')
            cls2_op = ((po.splitlines()[6]).split('|')[2]).replace(' ', '')
            cls2_admin = ((po.splitlines()[6]).split('|')[3]).replace(' ', '')
            cls2_type = ((po.splitlines()[5]).split('|')[2]).replace(' ', '')
            cls2_name = ((hs.splitlines()[2]).split('|')[2]).replace(' ', '')
            cls2_tx = ((xc.splitlines()[17]).split('|')[2]).replace(' ', '') + ' dBm'
            cls2_rx = ((xc.splitlines()[23]).split('|')[2]).replace(' ', '') + ' dBm'


            if 'no XCVR present' in xc:
                print(f'SFP not detect, Please check {cls2_name} Port {p[3]}')


                worksheet_ng.write('I12', cls2_type, centre)
                worksheet_ng.write('I13', cls2_des, centre)
                worksheet_ng.write('I14', cls2_op, centre)
                worksheet_ng.write('I15', cls2_admin, centre)
                worksheet_ng.write('I16', 'SFP not found', centre)
                worksheet_ng.write('I17', 'SFP not found', centre)


            else :

                worksheet_ng.write('I12', cls2_type, centre)
                worksheet_ng.write('I13', cls2_des, centre)
                worksheet_ng.write('I14', cls2_op, centre)
                worksheet_ng.write('I15', cls2_admin, centre)
                worksheet_ng.write('I16', cls2_tx, centre)
                worksheet_ng.write('I17', cls2_rx, centre)


        except NetMikoAuthenticationException :

            print(f'{i} could not be reached or connection timed out !!!')

            worksheet_ng.write('I12', 'N/A', centre)
            worksheet_ng.write('I13', 'N/A', centre)
            worksheet_ng.write('I14', 'N/A', centre)
            worksheet_ng.write('I15', 'N/A', centre)
            worksheet_ng.write('I16', 'N/A', centre)
            worksheet_ng.write('I17', 'N/A', centre)

        print()
        print('Saka - CLS route Completed!')









    #Ikeja - Saka route
    def ikeja_saka(self):


        hs = ['Ikeja 5171_01', 'Saka 5171_01']

        ne_p = {'172.20.38.252' : ['1/1', '2/1'],
        '172.20.38.249' : ['1/1', '2/1'], 
                }

        ne = [i for i in ne_p.keys()]
        p = [j for j in ne_p.values()]




        ##Ikeja - Saka route pry 
        worksheet_ng.merge_range('J8:K8', 'IKEJA - SAKA First Ring (34km)', centre_bold)
        worksheet_ng.write('J9', 'IKEJA POP', centre_bold)
        worksheet_ng.write('K9', 'SAKA POP', centre_bold)

        #Ikeja - Saka route secondary
        worksheet_ng.merge_range('L8:M8', 'IKEJA - SAKA Second Ring (41km)', centre_bold)
        worksheet_ng.write('L9', 'IKEJA POP', centre_bold)
        worksheet_ng.write('M9', 'SAKA POP', centre_bold)

        #Devices
        worksheet_ng.write('J10', hs[0], centre_bold)
        worksheet_ng.write('K10', hs[1], centre_bold)
        worksheet_ng.write('L10', hs[0], centre_bold)
        worksheet_ng.write('M10', hs[1], centre_bold)

        #Ports
        worksheet_ng.write('J11', p[0][0], centre_bold)
        worksheet_ng.write('K11', p[1][0], centre_bold)
        worksheet_ng.write('L11', p[0][1], centre_bold)
        worksheet_ng.write('M11', p[1][1], centre_bold)

        #ikeja 5171_01, saka 5171_01

        
        for i in ne:
                   
            ikj = {**metro_cet, 'host' : ne[0]}
            sak = {**metro_cet, 'host' : ne[1]}


        try:

            net_ssh = ConnectHandler(**ikj)    
            xc = net_ssh.send_command(f'port xc sh po {p[0][0]} diag')
            po = net_ssh.send_command(f'port sh po {p[0][0]}')
            hs = net_ssh.send_command('system show host-name')



            ikj_des = ((po.splitlines()[4]).split('|')[2]).replace(' ', '')
            ikj_op = ((po.splitlines()[6]).split('|')[2]).replace(' ', '')
            ikj_admin = ((po.splitlines()[6]).split('|')[3]).replace(' ', '')
            ikj_type = ((po.splitlines()[5]).split('|')[2]).replace(' ', '')
            ikj_name = ((hs.splitlines()[2]).split('|')[2]).replace(' ', '')
            ikj1_tx = ((xc.splitlines()[36]).split('|')[3]).replace(' ', '') + ' dBm'
            ikj2_tx = ((xc.splitlines()[51]).split('|')[3]).replace(' ', '') + ' dBm'
            ikj3_tx = ((xc.splitlines()[66]).split('|')[3]).replace(' ', '') + ' dBm'
            ikj4_tx = ((xc.splitlines()[81]).split('|')[3]).replace(' ', '') + ' dBm'
            ikj1_rx = ((xc.splitlines()[42]).split('|')[3]).replace(' ', '') + ' dBm'
            ikj2_rx = ((xc.splitlines()[57]).split('|')[3]).replace(' ', '') + ' dBm'
            ikj3_rx = ((xc.splitlines()[72]).split('|')[3]).replace(' ', '') + ' dBm'
            ikj4_rx = ((xc.splitlines()[87]).split('|')[3]).replace(' ', '') + ' dBm'
            ikj_tx = [ikj1_tx, ikj2_tx, ikj3_tx, ikj4_tx]
            ikj_rx = [ikj1_rx, ikj2_rx, ikj3_rx, ikj4_rx]


            if 'no XCVR present' in xc:

                print(f'SFP not detect, Please check {ikj_name} Port {p[0][0]}')
        

                worksheet_ng.write('J12', ikj_type, centre)
                worksheet_ng.write('J13', ikj_des, centre)
                worksheet_ng.write('J14', ikj_op, centre)
                worksheet_ng.write('J15', ikj_admin, centre)
                worksheet_ng.write('J16', 'SFP not found', centre)
                worksheet_ng.write('J17', 'SFP not found', centre)


            else :      

                worksheet_ng.write('J12', ikj_type, centre)
                worksheet_ng.write('J13', ikj_des, centre)
                worksheet_ng.write('J14', ikj_op, centre)
                worksheet_ng.write('J15', ikj_admin, centre)
                worksheet_ng.write('J16', '\n'.join(ikj_tx), centre)
                worksheet_ng.write('J17', '\n'.join(ikj_rx), centre)


        except NetMikoAuthenticationException :

            print(f'{i} could not be reached or connection timed out !!!')


            worksheet_ng.write('J12', 'N/A', centre)
            worksheet_ng.write('J13', 'N/A', centre)
            worksheet_ng.write('J14', 'N/A', centre)
            worksheet_ng.write('J15', 'N/A', centre)
            worksheet_ng.write('J16', 'N/A', centre)
            worksheet_ng.write('J17', 'N/A', centre)


        try:

            net_ssh = ConnectHandler(**sak)    
            xc = net_ssh.send_command(f'port xc sh po {p[1][0]} diag')
            po = net_ssh.send_command(f'port sh po {p[1][0]}')
            hs = net_ssh.send_command('system show host-name')


            sak_des = ((po.splitlines()[4]).split('|')[2]).replace(' ', '')
            sak_op = ((po.splitlines()[6]).split('|')[2]).replace(' ', '')
            sak_admin = ((po.splitlines()[6]).split('|')[3]).replace(' ', '')
            sak_type = ((po.splitlines()[5]).split('|')[2]).replace(' ', '')
            sak_name = ((hs.splitlines()[2]).split('|')[2]).replace(' ', '')
            sak1_tx = ((xc.splitlines()[36]).split('|')[3]).replace(' ', '') + ' dBm'
            sak2_tx = ((xc.splitlines()[51]).split('|')[3]).replace(' ', '') + ' dBm'
            sak3_tx = ((xc.splitlines()[66]).split('|')[3]).replace(' ', '') + ' dBm'
            sak4_tx = ((xc.splitlines()[81]).split('|')[3]).replace(' ', '') + ' dBm'
            sak1_rx = ((xc.splitlines()[42]).split('|')[3]).replace(' ', '') + ' dBm'
            sak2_rx = ((xc.splitlines()[57]).split('|')[3]).replace(' ', '') + ' dBm'
            sak3_rx = ((xc.splitlines()[72]).split('|')[3]).replace(' ', '') + ' dBm'
            sak4_rx = ((xc.splitlines()[87]).split('|')[3]).replace(' ', '') + ' dBm'
            sak_tx = [sak1_tx, sak2_tx, sak3_tx, sak4_tx]
            sak_rx = [sak1_rx, sak2_rx, sak3_rx, sak4_rx]




            if 'no XCVR present' in xc:
                print(f'SFP not detect, Please check {sak_name} Port {p[1][0]}')


                worksheet_ng.write('K12', sak_type, centre)
                worksheet_ng.write('K13', sak_des, centre)
                worksheet_ng.write('K14', sak_op, centre)
                worksheet_ng.write('K15', sak_admin, centre)
                worksheet_ng.write('K16', 'SFP not found', centre)
                worksheet_ng.write('K17', 'SFP not found', centre)



            else :

                worksheet_ng.write('K12', sak_type, centre)
                worksheet_ng.write('K13', sak_des, centre)
                worksheet_ng.write('K14', sak_op, centre)
                worksheet_ng.write('K15', sak_admin, centre)
                worksheet_ng.write('K16', '\n'.join(sak_tx), centre)
                worksheet_ng.write('K17', '\n'.join(sak_rx), centre)



        except NetMikoAuthenticationException :

            print(f'{i} could not be reached or connection timed out !!!')

            worksheet_ng.write('K12', 'N/A', centre)
            worksheet_ng.write('K13', 'N/A', centre)
            worksheet_ng.write('K14', 'N/A', centre)
            worksheet_ng.write('K15', 'N/A', centre)
            worksheet_ng.write('K16', 'N/A', centre)
            worksheet_ng.write('K17', 'N/A', centre)


        try:

            net_ssh = ConnectHandler(**ikj)    
            xc = net_ssh.send_command('port xc sh po 2/1 diag')
            po = net_ssh.send_command('port sh po 2/1')
            hs = net_ssh.send_command('system show host-name')



            ikj_des = ((po.splitlines()[4]).split('|')[2]).replace(' ', '')
            ikj_op = ((po.splitlines()[6]).split('|')[2]).replace(' ', '')
            ikj_admin = ((po.splitlines()[6]).split('|')[3]).replace(' ', '')
            ikj_type = ((po.splitlines()[5]).split('|')[2]).replace(' ', '')
            ikj_name = ((hs.splitlines()[2]).split('|')[2]).replace(' ', '')
            ikj1_tx = ((xc.splitlines()[36]).split('|')[3]).replace(' ', '') + ' dBm'
            ikj2_tx = ((xc.splitlines()[51]).split('|')[3]).replace(' ', '') + ' dBm'
            ikj3_tx = ((xc.splitlines()[66]).split('|')[3]).replace(' ', '') + ' dBm'
            ikj4_tx = ((xc.splitlines()[81]).split('|')[3]).replace(' ', '') + ' dBm'
            ikj1_rx = ((xc.splitlines()[42]).split('|')[3]).replace(' ', '') + ' dBm'
            ikj2_rx = ((xc.splitlines()[57]).split('|')[3]).replace(' ', '') + ' dBm'
            ikj3_rx = ((xc.splitlines()[72]).split('|')[3]).replace(' ', '') + ' dBm'
            ikj4_rx = ((xc.splitlines()[87]).split('|')[3]).replace(' ', '') + ' dBm'
            ikj_tx = [ikj1_tx, ikj2_tx, ikj3_tx, ikj4_tx]
            ikj_rx = [ikj1_rx, ikj2_rx, ikj3_rx, ikj4_rx]



            if 'no XCVR present' in xc:

                print(f'SFP not detect, Please check {ikj_name} Port {p[0][1]}')
        

                worksheet_ng.write('J12', ikj_type, centre)
                worksheet_ng.write('J13', ikj_des, centre)
                worksheet_ng.write('J14', ikj_op, centre)
                worksheet_ng.write('J15', ikj_admin, centre)
                worksheet_ng.write('J16', 'SFP not found', centre)
                worksheet_ng.write('J17', 'SFP not found', centre)


            else :      

                worksheet_ng.write('L12', ikj_type, centre)
                worksheet_ng.write('L13', ikj_des, centre)
                worksheet_ng.write('L14', ikj_op, centre)
                worksheet_ng.write('L15', ikj_admin, centre)
                worksheet_ng.write('L16', '\n'.join(ikj_tx), centre)
                worksheet_ng.write('L17', '\n'.join(ikj_rx), centre)


        except NetMikoAuthenticationException :

            print(f'{i} could not be reached or connection timed out !!!')


            worksheet_ng.write('L12', 'N/A', centre)
            worksheet_ng.write('L13', 'N/A', centre)
            worksheet_ng.write('L14', 'N/A', centre)
            worksheet_ng.write('L15', 'N/A', centre)
            worksheet_ng.write('L16', 'N/A', centre)
            worksheet_ng.write('L17', 'N/A', centre)


        try:

            net_ssh = ConnectHandler(**sak)    
            xc = net_ssh.send_command(f'port xc sh po {p[0][1]} diag')
            po = net_ssh.send_command(f'port sh po {p[0][1]}')
            hs = net_ssh.send_command('system show host-name')


            sak_des = ((po.splitlines()[4]).split('|')[2]).replace(' ', '')
            sak_op = ((po.splitlines()[6]).split('|')[2]).replace(' ', '')
            sak_admin = ((po.splitlines()[6]).split('|')[3]).replace(' ', '')
            sak_type = ((po.splitlines()[5]).split('|')[2]).replace(' ', '')
            sak_name = ((hs.splitlines()[2]).split('|')[2]).replace(' ', '')
            sak1_tx = ((xc.splitlines()[36]).split('|')[3]).replace(' ', '') + ' dBm'
            sak2_tx = ((xc.splitlines()[51]).split('|')[3]).replace(' ', '') + ' dBm'
            sak3_tx = ((xc.splitlines()[66]).split('|')[3]).replace(' ', '') + ' dBm'
            sak4_tx = ((xc.splitlines()[81]).split('|')[3]).replace(' ', '') + ' dBm'
            sak1_rx = ((xc.splitlines()[42]).split('|')[3]).replace(' ', '') + ' dBm'
            sak2_rx = ((xc.splitlines()[57]).split('|')[3]).replace(' ', '') + ' dBm'
            sak3_rx = ((xc.splitlines()[72]).split('|')[3]).replace(' ', '') + ' dBm'
            sak4_rx = ((xc.splitlines()[87]).split('|')[3]).replace(' ', '') + ' dBm'
            sak_tx = [sak1_tx, sak2_tx, sak3_tx, sak4_tx]
            sak_rx = [sak1_rx, sak2_rx, sak3_rx, sak4_rx]


            if 'no XCVR present' in xc:
                print(f'SFP not detect, Please check {sak_name} Port {p[1][1]}')


                worksheet_ng.write('M12', sak_type, centre)
                worksheet_ng.write('M13', sak_des, centre)
                worksheet_ng.write('M14', sak_op, centre)
                worksheet_ng.write('M15', sak_admin, centre)
                worksheet_ng.write('M16', 'SFP not found', centre)
                worksheet_ng.write('M17', 'SFP not found', centre)



            else :

                worksheet_ng.write('M12', sak_type, centre)
                worksheet_ng.write('M13', sak_des, centre)
                worksheet_ng.write('M14', sak_op, centre)
                worksheet_ng.write('M15', sak_admin, centre)
                worksheet_ng.write('M16', '\n'.join(sak_tx), centre)
                worksheet_ng.write('M17', '\n'.join(sak_rx), centre)



        except NetMikoAuthenticationException :

            print(f'{i} could not be reached or connection timed out !!!')

            worksheet_ng.write('M12', 'N/A', centre)
            worksheet_ng.write('M13', 'N/A', centre)
            worksheet_ng.write('M14', 'N/A', centre)
            worksheet_ng.write('M15', 'N/A', centre)
            worksheet_ng.write('M16', 'N/A', centre)
            worksheet_ng.write('M17', 'N/A', centre)

        print()
        print('Ikeja - Saka route Completed!')


Tx_Lagos_Core.saka_cls(Self)

Tx_Lagos_Core.ikeja_saka(Self)


print()
print('-'*80)
print('Task Completed!')
print('-'*80)


workbook.close()


