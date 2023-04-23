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

#Login credentials
username = input('Enter Username: ')
password = getpass()

#Creating dictionary representing the devices
metro_cet = {
'device_type': 'ciena_saos',
'username': username,
'password': password,
}


#defining the date/time stamp
dt = datetime.now()

#defining the date/time stamp format
t_stamp = dt.strftime("%d.%m.%Y_%H.%M.%S")


#create the excel spreed sheet
workbook = xlsxwriter.Workbook(f'<#directory>\\<Report_filename>_{t_stamp}.xlsx')

#create the workbook on the excel spreed sheet
worksheet_ng = workbook.add_worksheet('<#excel_file_worksheet_name>')


#defining the format of the cells in the excel sprred sheet
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


# class
class Tx_BB_Core():

    # init function
    def __init__(self) :
        pass


    #Broadening the cells width
    worksheet_ng.set_column('A:E', 18)
    worksheet_ng.set_column('F:BZ', 10g_p1)

    #freeze row 1-7, col A-E
    worksheet_ng.freeze_panes(7, 5)


    #Legends
    worksheet_ng.merge_range('A1:E1', 'Legend', centre_bold15)
    worksheet_ng.write('A2', 'Ports', centre_bold)
    worksheet_ng.write('A3', '10G (10km - 20km)', centre_bold)
    worksheet_ng.write('A4', '10G (10g_p1km - 80km)', centre_bold)
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
    worksheet_ng.write('C6', '-9.10g_p1', centre_bold)
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



    #function route_central_eastern

    def central_eastern(self):
        #nertwork elements   
   
        hs = ['ne1_01', 'ne2_01', 'ne1_02', 'ne2_02' ]
        
        
        #nertwork element mgmt ip and port dictionary

        ne_p = {'x11.x11.x11.x11' : '10g_p1', 
        'x21.x21.x21.x21' : '10g_p2', 
        'x12.x12.x12.x12' : '10g_p1',
        'x22.x22.x22.x22' : '10g_p2'
        }

        ne = [i for i in ne_p.keys()]
        p = [j for j in ne_p.values()]


        #Route and POP
        worksheet_ng.merge_range('F8:G8', 'Route_ab (Pry: 30km)', centre_bold)
        worksheet_ng.write('F9', 'central_pop', centre_bold)
        worksheet_ng.write('G9', 'Eastern_pop', centre_bold)

        #Route and POP
        worksheet_ng.merge_range('H8:I8', 'Route_ab (Secondary: 28.2km)', centre_bold)
        worksheet_ng.write('H9', 'central_pop', centre_bold)
        worksheet_ng.write('I9', 'Eastern_pop', centre_bold)
        

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


  
        
        
        #login for NEs
        
        for i in ne:

            #login to NEs            
            central1 = {**metro_cet, 'host' : ne[0]}
            eastern1 = {**metro_cet, 'host' : ne[1]}
            central2 = {**metro_cet, 'host' : ne[2]}
            eastern2 = {**metro_cet, 'host' : ne[3]}


        try:
             # CLI inputs on the NEs
            net_ssh = ConnectHandler(**central1)    
            xc = net_ssh.send_command(f'port xc sh po {p[0]} diag')
            po = net_ssh.send_command(f'port sh po {p[0]}')
            hs = net_ssh.send_command('system show host-name')


            # filtering CLI outputs on the NEs
            central1_des = ((po.splitlines()[4]).split('|')[2]).replace(' ', '')
            central1_op = ((po.splitlines()[6]).split('|')[2]).replace(' ', '')
            central1_admin = ((po.splitlines()[6]).split('|')[3]).replace(' ', '')
            central1_type = ((po.splitlines()[5]).split('|')[2]).replace(' ', '')
            central1_name = ((hs.splitlines()[2]).split('|')[2]).replace(' ', '')
            central1_tx = ((xc.splitlines()[17]).split('|')[2]).replace(' ', '') + ' dBm'
            central1_rx = ((xc.splitlines()[23]).split('|')[2]).replace(' ', '') + ' dBm'

            
            if 'no XCVR present' in xc:             #empty port
                
                print(f'SFP not detect, Please check {central1_name} Port {p[0]}')
        
                #writing the CLI outputs on the excel
                worksheet_ng.write('F12', central1_type, centre)
                worksheet_ng.write('F13', central1_des, centre)
                worksheet_ng.write('F14', central1_op, centre)
                worksheet_ng.write('F15', central1_admin, centre)
                worksheet_ng.write('F16', 'SFP not found', centre)
                worksheet_ng.write('F17', 'SFP not found', centre)


            else :      
                #writing the CLI outputs on the excel
                worksheet_ng.write('F12', central1_type, centre)
                worksheet_ng.write('F13', central1_des, centre)
                worksheet_ng.write('F14', central1_op, centre)
                worksheet_ng.write('F15', central1_admin, centre)
                worksheet_ng.write('F16', central1_tx, centre)
                worksheet_ng.write('F17', central1_rx, centre)


        except NetMikoAuthenticationException :             #connection timed out/error

            print(f'{i} could not be reached or connection timed out !!!')

            #writing the CLI outputs on the excel
            worksheet_ng.write('F12', 'N/A', centre)
            worksheet_ng.write('F13', 'N/A', centre)
            worksheet_ng.write('F14', 'N/A', centre)
            worksheet_ng.write('F15', 'N/A', centre)
            worksheet_ng.write('F16', 'N/A', centre)
            worksheet_ng.write('F17', 'N/A', centre)


        try:
            # CLI inputs on the NEs
            net_ssh = ConnectHandler(**eastern1)    
            xc = net_ssh.send_command(f'port xc sh po {p[1]} diag')
            po = net_ssh.send_command(f'port sh po {p[1]}')
            hs = net_ssh.send_command('system show host-name')

            # filtering CLI outputs on the NEs
            eastern1_des = ((po.splitlines()[4]).split('|')[2]).replace(' ', '')
            eastern1_op = ((po.splitlines()[6]).split('|')[2]).replace(' ', '')
            eastern1_admin = ((po.splitlines()[6]).split('|')[3]).replace(' ', '')
            eastern1_type = ((po.splitlines()[5]).split('|')[2]).replace(' ', '')
            eastern1_name = ((hs.splitlines()[2]).split('|')[2]).replace(' ', '')
            eastern1_tx = ((xc.splitlines()[17]).split('|')[2]).replace(' ', '') + ' dBm'
            eastern1_rx = ((xc.splitlines()[23]).split('|')[2]).replace(' ', '') + ' dBm'

            
            if 'no XCVR present' in xc:             #empty port
                print(f'SFP not detect, Please check {eastern1_name} Port {p[1]}')

                #writing the CLI outputs on the excel
                worksheet_ng.write('G12', eastern1_type, centre)
                worksheet_ng.write('G13', eastern1_des, centre)
                worksheet_ng.write('G14', eastern1_op, centre)
                worksheet_ng.write('G15', eastern1_admin, centre)
                worksheet_ng.write('G16', 'SFP not found', centre)
                worksheet_ng.write('G17', 'SFP not found', centre)



            else :
                #writing the CLI outputs on the excel
                worksheet_ng.write('G12', eastern1_type, centre)
                worksheet_ng.write('G13', eastern1_des, centre)
                worksheet_ng.write('G14', eastern1_op, centre)
                worksheet_ng.write('G15', eastern1_admin, centre)
                worksheet_ng.write('G16', eastern1_tx, centre)
                worksheet_ng.write('G17', eastern1_rx, centre)



        except NetMikoAuthenticationException :                 #connection timed out/error

            print(f'{i} could not be reached or connection timed out !!!')
            #writing the CLI outputs on the excel
            worksheet_ng.write('G12', 'N/A', centre)
            worksheet_ng.write('G13', 'N/A', centre)
            worksheet_ng.write('G14', 'N/A', centre)
            worksheet_ng.write('G15', 'N/A', centre)
            worksheet_ng.write('G16', 'N/A', centre)
            worksheet_ng.write('G17', 'N/A', centre)


        try:
            # CLI inputs on the NEs
            net_ssh = ConnectHandler(**central2)    
            xc = net_ssh.send_command(f'port xc sh po {p[2]} diag')
            po = net_ssh.send_command(f'port sh po {p[2]}')
            hs = net_ssh.send_command('system show host-name')

            # filtering CLI outputs on the NEs
            central2_des = ((po.splitlines()[4]).split('|')[2]).replace(' ', '')
            central2_op = ((po.splitlines()[6]).split('|')[2]).replace(' ', '')
            central2_admin = ((po.splitlines()[6]).split('|')[3]).replace(' ', '')
            central2_type = ((po.splitlines()[5]).split('|')[2]).replace(' ', '')
            central2_name = ((hs.splitlines()[2]).split('|')[2]).replace(' ', '')
            central2_tx = ((xc.splitlines()[17]).split('|')[2]).replace(' ', '') + ' dBm'
            central2_rx = ((xc.splitlines()[23]).split('|')[2]).replace(' ', '') + ' dBm'


            if 'no XCVR present' in xc:               #empty port         

                print(f'SFP not detect, Please check {central2_name} Port {p[2]}')

                #writing the CLI outputs on the excel
                worksheet_ng.write('H12', central2_type, centre)
                worksheet_ng.write('H13', central2_des, centre)
                worksheet_ng.write('H14', central2_op, centre)
                worksheet_ng.write('H15', central2_admin, centre)
                worksheet_ng.write('H16', 'SFP not found', centre)
                worksheet_ng.write('H17', 'SFP not found', centre)


            else :
                #writing the CLI outputs on the excel
                worksheet_ng.write('H12', central2_type, centre)
                worksheet_ng.write('H13', central2_des, centre)
                worksheet_ng.write('H14', central2_op, centre)
                worksheet_ng.write('H15', central2_admin, centre)
                worksheet_ng.write('H16', central2_tx, centre)
                worksheet_ng.write('H17', central2_rx, centre)


        except NetMikoAuthenticationException :                #empty port 

            print(f'{i} could not be reached or connection timed out !!!')
            #writing the CLI outputs on the excel
            worksheet_ng.write('H12', 'N/A', centre)
            worksheet_ng.write('H13', 'N/A', centre)
            worksheet_ng.write('H14', 'N/A', centre)
            worksheet_ng.write('H15', 'N/A', centre)
            worksheet_ng.write('H16', 'N/A', centre)
            worksheet_ng.write('H17', 'N/A', centre)


        try:
             # CLI inputs on the NEs
            net_ssh = ConnectHandler(**eastern2)    
            xc = net_ssh.send_command(f'port xc sh po {p[3]} diag')
            po = net_ssh.send_command(f'port sh po {p[3]}')
            hs = net_ssh.send_command('system show host-name')

             # filtering CLI outputs on the NEs
            eastern2_des = ((po.splitlines()[4]).split('|')[2]).replace(' ', '')
            eastern2_op = ((po.splitlines()[6]).split('|')[2]).replace(' ', '')
            eastern2_admin = ((po.splitlines()[6]).split('|')[3]).replace(' ', '')
            eastern2_type = ((po.splitlines()[5]).split('|')[2]).replace(' ', '')
            eastern2_name = ((hs.splitlines()[2]).split('|')[2]).replace(' ', '')
            eastern2_tx = ((xc.splitlines()[17]).split('|')[2]).replace(' ', '') + ' dBm'
            eastern2_rx = ((xc.splitlines()[23]).split('|')[2]).replace(' ', '') + ' dBm'


            if 'no XCVR present' in xc:                     #empty port 
                print(f'SFP not detect, Please check {eastern2_name} Port {p[3]}')

                #writing the CLI outputs on the excel
                worksheet_ng.write('I12', eastern2_type, centre)
                worksheet_ng.write('I13', eastern2_des, centre)
                worksheet_ng.write('I14', eastern2_op, centre)
                worksheet_ng.write('I15', eastern2_admin, centre)
                worksheet_ng.write('I16', 'SFP not found', centre)
                worksheet_ng.write('I17', 'SFP not found', centre)


            else :
                #writing the CLI outputs on the excel
                worksheet_ng.write('I12', eastern2_type, centre)
                worksheet_ng.write('I13', eastern2_des, centre)
                worksheet_ng.write('I14', eastern2_op, centre)
                worksheet_ng.write('I15', eastern2_admin, centre)
                worksheet_ng.write('I16', eastern2_tx, centre)
                worksheet_ng.write('I17', eastern2_rx, centre)


        except NetMikoAuthenticationException :                      #empty port 

            print(f'{i} could not be reached or connection timed out !!!')
            
            #writing the CLI outputs on the excel
            worksheet_ng.write('I12', 'N/A', centre)
            worksheet_ng.write('I13', 'N/A', centre)
            worksheet_ng.write('I14', 'N/A', centre)
            worksheet_ng.write('I15', 'N/A', centre)
            worksheet_ng.write('I16', 'N/A', centre)
            worksheet_ng.write('I17', 'N/A', centre)

        print()
        print('Central - Eastern_pop route Completed!')









    #Western - Central route
    def western_central(self):

        #nertwork elements
        hs = ['ne3_01', 'ne1_01']
        #nertwork element mgmt ip and port dictionary
        ne_p = {'x31.x31.x31.x31' : ['100g_p11', '100g_p21'],
        'x11.x11.x11.x11' : ['100g_p11', '100g_p21'], 
                }
        #NE Ip address, port list
        ne = [i for i in ne_p.keys()]
        p = [j for j in ne_p.values()]




        #western - Central route pry 
        worksheet_ng.merge_range('J8:K8', 'Western - Central First Ring (34km)', centre_bold)
        worksheet_ng.write('J9', 'Western POP', centre_bold)
        worksheet_ng.write('K9', 'central_pop', centre_bold)

        #western - Central route secondary
        worksheet_ng.merge_range('L8:M8', 'Western - Central Second Ring (41km)', centre_bold)
        worksheet_ng.write('L9', 'Western POP', centre_bold)
        worksheet_ng.write('M9', 'central_pop', centre_bold)

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

        
        #login to NEs
        
        for i in ne:
                   
            western = {**metro_cet, 'host' : ne[0]}
            central = {**metro_cet, 'host' : ne[1]}


        try:
            
            # CLI inputs on the NEs
            net_ssh = ConnectHandler(**western)    
            xc = net_ssh.send_command(f'port xc sh po {p[0][0]} diag')
            po = net_ssh.send_command(f'port sh po {p[0][0]}')
            hs = net_ssh.send_command('system show host-name')

            
            # filtering CLI outputs on the NEs

            western_des = ((po.splitlines()[4]).split('|')[2]).replace(' ', '')
            western_op = ((po.splitlines()[6]).split('|')[2]).replace(' ', '')
            western_admin = ((po.splitlines()[6]).split('|')[3]).replace(' ', '')
            western_type = ((po.splitlines()[5]).split('|')[2]).replace(' ', '')
            western_name = ((hs.splitlines()[2]).split('|')[2]).replace(' ', '')
            western1_tx = ((xc.splitlines()[36]).split('|')[3]).replace(' ', '') + ' dBm'
            western2_tx = ((xc.splitlines()[51]).split('|')[3]).replace(' ', '') + ' dBm'
            western3_tx = ((xc.splitlines()[66]).split('|')[3]).replace(' ', '') + ' dBm'
            western4_tx = ((xc.splitlines()[81]).split('|')[3]).replace(' ', '') + ' dBm'
            western1_rx = ((xc.splitlines()[42]).split('|')[3]).replace(' ', '') + ' dBm'
            western2_rx = ((xc.splitlines()[57]).split('|')[3]).replace(' ', '') + ' dBm'
            western3_rx = ((xc.splitlines()[72]).split('|')[3]).replace(' ', '') + ' dBm'
            western4_rx = ((xc.splitlines()[87]).split('|')[3]).replace(' ', '') + ' dBm'
            western_tx = [western1_tx, western2_tx, western3_tx, western4_tx]
            western_rx = [western1_rx, western2_rx, western3_rx, western4_rx]

            
            if 'no XCVR present' in xc:                         #empty port
                
                #empty port
                print(f'SFP not detect, Please check {western_name} Port {p[0][0]}')
        
                #writing the CLI outputs on the excel                
                worksheet_ng.write('J12', western_type, centre)
                worksheet_ng.write('J13', western_des, centre)
                worksheet_ng.write('J14', western_op, centre)
                worksheet_ng.write('J15', western_admin, centre)
                worksheet_ng.write('J16', 'SFP not found', centre)
                worksheet_ng.write('J17', 'SFP not found', centre)


            else :      
                
                #writing the CLI outputs on the excel
                worksheet_ng.write('J12', western_type, centre)
                worksheet_ng.write('J13', western_des, centre)
                worksheet_ng.write('J14', western_op, centre)
                worksheet_ng.write('J15', western_admin, centre)
                worksheet_ng.write('J16', '\n'.join(western_tx), centre)
                worksheet_ng.write('J17', '\n'.join(western_rx), centre)


        except NetMikoAuthenticationException :              #connection timed out/error
           
            print(f'{i} could not be reached or connection timed out !!!')

            #writing the CLI outputs on the excel
            worksheet_ng.write('J12', 'N/A', centre)
            worksheet_ng.write('J13', 'N/A', centre)
            worksheet_ng.write('J14', 'N/A', centre)
            worksheet_ng.write('J15', 'N/A', centre)
            worksheet_ng.write('J16', 'N/A', centre)
            worksheet_ng.write('J17', 'N/A', centre)


        try:
            # CLI inputs on the NEs
            net_ssh = ConnectHandler(**central)    
            xc = net_ssh.send_command(f'port xc sh po {p[1][0]} diag')
            po = net_ssh.send_command(f'port sh po {p[1][0]}')
            hs = net_ssh.send_command('system show host-name')

            # filtering CLI outputs on the NEs
            central_des = ((po.splitlines()[4]).split('|')[2]).replace(' ', '')
            central_op = ((po.splitlines()[6]).split('|')[2]).replace(' ', '')
            central_admin = ((po.splitlines()[6]).split('|')[3]).replace(' ', '')
            central_type = ((po.splitlines()[5]).split('|')[2]).replace(' ', '')
            central_name = ((hs.splitlines()[2]).split('|')[2]).replace(' ', '')
            central1_tx = ((xc.splitlines()[36]).split('|')[3]).replace(' ', '') + ' dBm'
            central2_tx = ((xc.splitlines()[51]).split('|')[3]).replace(' ', '') + ' dBm'
            central3_tx = ((xc.splitlines()[66]).split('|')[3]).replace(' ', '') + ' dBm'
            central4_tx = ((xc.splitlines()[81]).split('|')[3]).replace(' ', '') + ' dBm'
            central1_rx = ((xc.splitlines()[42]).split('|')[3]).replace(' ', '') + ' dBm'
            central2_rx = ((xc.splitlines()[57]).split('|')[3]).replace(' ', '') + ' dBm'
            central3_rx = ((xc.splitlines()[72]).split('|')[3]).replace(' ', '') + ' dBm'
            central4_rx = ((xc.splitlines()[87]).split('|')[3]).replace(' ', '') + ' dBm'
            central_tx = [central1_tx, central2_tx, central3_tx, central4_tx]
            central_rx = [central1_rx, central2_rx, central3_rx, central4_rx]



            #writing the CLI outputs on the excel
            if 'no XCVR present' in xc:                #empty port
  
                print(f'SFP not detect, Please check {central_name} Port {p[1][0]}')

                #writing the CLI outputs on the excel
                worksheet_ng.write('K12', central_type, centre)
                worksheet_ng.write('K13', central_des, centre)
                worksheet_ng.write('K14', central_op, centre)
                worksheet_ng.write('K15', central_admin, centre)
                worksheet_ng.write('K16', 'SFP not found', centre)
                worksheet_ng.write('K17', 'SFP not found', centre)



            else :
                #writing the CLI outputs on the excel
                worksheet_ng.write('K12', central_type, centre)
                worksheet_ng.write('K13', central_des, centre)
                worksheet_ng.write('K14', central_op, centre)
                worksheet_ng.write('K15', central_admin, centre)
                worksheet_ng.write('K16', '\n'.join(central_tx), centre)
                worksheet_ng.write('K17', '\n'.join(central_rx), centre)



        except NetMikoAuthenticationException :             #connection timed out/error
            
   
            print(f'{i} could not be reached or connection timed out !!!')
            
            #writing the CLI outputs on the excel
            worksheet_ng.write('K12', 'N/A', centre)
            worksheet_ng.write('K13', 'N/A', centre)
            worksheet_ng.write('K14', 'N/A', centre)
            worksheet_ng.write('K15', 'N/A', centre)
            worksheet_ng.write('K16', 'N/A', centre)
            worksheet_ng.write('K17', 'N/A', centre)


        try:
            # CLI inputs on the NEs
            net_ssh = ConnectHandler(**western)    
            xc = net_ssh.send_command('port xc sh po 100g_p21 diag')
            po = net_ssh.send_command('port sh po 100g_p21')
            hs = net_ssh.send_command('system show host-name')

            
            # filtering CLI outputs on the NEs

            western_des = ((po.splitlines()[4]).split('|')[2]).replace(' ', '')
            western_op = ((po.splitlines()[6]).split('|')[2]).replace(' ', '')
            western_admin = ((po.splitlines()[6]).split('|')[3]).replace(' ', '')
            western_type = ((po.splitlines()[5]).split('|')[2]).replace(' ', '')
            western_name = ((hs.splitlines()[2]).split('|')[2]).replace(' ', '')
            western1_tx = ((xc.splitlines()[36]).split('|')[3]).replace(' ', '') + ' dBm'
            western2_tx = ((xc.splitlines()[51]).split('|')[3]).replace(' ', '') + ' dBm'
            western3_tx = ((xc.splitlines()[66]).split('|')[3]).replace(' ', '') + ' dBm'
            western4_tx = ((xc.splitlines()[81]).split('|')[3]).replace(' ', '') + ' dBm'
            western1_rx = ((xc.splitlines()[42]).split('|')[3]).replace(' ', '') + ' dBm'
            western2_rx = ((xc.splitlines()[57]).split('|')[3]).replace(' ', '') + ' dBm'
            western3_rx = ((xc.splitlines()[72]).split('|')[3]).replace(' ', '') + ' dBm'
            western4_rx = ((xc.splitlines()[87]).split('|')[3]).replace(' ', '') + ' dBm'
            western_tx = [western1_tx, western2_tx, western3_tx, western4_tx]
            western_rx = [western1_rx, western2_rx, western3_rx, western4_rx]


           
            if 'no XCVR present' in xc:                 #empty port
   
                print(f'SFP not detect, Please check {western_name} Port {p[0][1]}')
        
                #writing the CLI outputs on the excel
                worksheet_ng.write('J12', western_type, centre)
                worksheet_ng.write('J13', western_des, centre)
                worksheet_ng.write('J14', western_op, centre)
                worksheet_ng.write('J15', western_admin, centre)
                worksheet_ng.write('J16', 'SFP not found', centre)
                worksheet_ng.write('J17', 'SFP not found', centre)


            else :      
                
                #writing the CLI outputs on the excel
                worksheet_ng.write('L12', western_type, centre)
                worksheet_ng.write('L13', western_des, centre)
                worksheet_ng.write('L14', western_op, centre)
                worksheet_ng.write('L15', western_admin, centre)
                worksheet_ng.write('L16', '\n'.join(western_tx), centre)
                worksheet_ng.write('L17', '\n'.join(western_rx), centre)


        except NetMikoAuthenticationException :               #connection timed out/error
            

            print(f'{i} could not be reached or connection timed out !!!')

            #writing the CLI outputs on the excel
            worksheet_ng.write('L12', 'N/A', centre)
            worksheet_ng.write('L13', 'N/A', centre)
            worksheet_ng.write('L14', 'N/A', centre)
            worksheet_ng.write('L15', 'N/A', centre)
            worksheet_ng.write('L16', 'N/A', centre)
            worksheet_ng.write('L17', 'N/A', centre)


        try:
            # CLI inputs on the NEs
            net_ssh = ConnectHandler(**central)    
            xc = net_ssh.send_command(f'port xc sh po {p[0][1]} diag')
            po = net_ssh.send_command(f'port sh po {p[0][1]}')
            hs = net_ssh.send_command('system show host-name')

            # filtering CLI outputs on the NEs
            central_des = ((po.splitlines()[4]).split('|')[2]).replace(' ', '')
            central_op = ((po.splitlines()[6]).split('|')[2]).replace(' ', '')
            central_admin = ((po.splitlines()[6]).split('|')[3]).replace(' ', '')
            central_type = ((po.splitlines()[5]).split('|')[2]).replace(' ', '')
            central_name = ((hs.splitlines()[2]).split('|')[2]).replace(' ', '')
            central1_tx = ((xc.splitlines()[36]).split('|')[3]).replace(' ', '') + ' dBm'
            central2_tx = ((xc.splitlines()[51]).split('|')[3]).replace(' ', '') + ' dBm'
            central3_tx = ((xc.splitlines()[66]).split('|')[3]).replace(' ', '') + ' dBm'
            central4_tx = ((xc.splitlines()[81]).split('|')[3]).replace(' ', '') + ' dBm'
            central1_rx = ((xc.splitlines()[42]).split('|')[3]).replace(' ', '') + ' dBm'
            central2_rx = ((xc.splitlines()[57]).split('|')[3]).replace(' ', '') + ' dBm'
            central3_rx = ((xc.splitlines()[72]).split('|')[3]).replace(' ', '') + ' dBm'
            central4_rx = ((xc.splitlines()[87]).split('|')[3]).replace(' ', '') + ' dBm'
            central_tx = [central1_tx, central2_tx, central3_tx, central4_tx]
            central_rx = [central1_rx, central2_rx, central3_rx, central4_rx]

            
            if 'no XCVR present' in xc:                  #empty port

                print(f'SFP not detect, Please check {central_name} Port {p[1][1]}')

                #writing the CLI outputs on the excel
                worksheet_ng.write('M12', central_type, centre)
                worksheet_ng.write('M13', central_des, centre)
                worksheet_ng.write('M14', central_op, centre)
                worksheet_ng.write('M15', central_admin, centre)
                worksheet_ng.write('M16', 'SFP not found', centre)
                worksheet_ng.write('M17', 'SFP not found', centre)



            else :
                #writing the CLI outputs on the excel
                worksheet_ng.write('M12', central_type, centre)
                worksheet_ng.write('M13', central_des, centre)
                worksheet_ng.write('M14', central_op, centre)
                worksheet_ng.write('M15', central_admin, centre)
                worksheet_ng.write('M16', '\n'.join(central_tx), centre)
                worksheet_ng.write('M17', '\n'.join(central_rx), centre)



        except NetMikoAuthenticationException :            #connection timed out/error
    
            print(f'{i} could not be reached or connection timed out !!!')
            #writing the CLI outputs on the excel
            worksheet_ng.write('M12', 'N/A', centre)
            worksheet_ng.write('M13', 'N/A', centre)
            worksheet_ng.write('M14', 'N/A', centre)
            worksheet_ng.write('M15', 'N/A', centre)
            worksheet_ng.write('M16', 'N/A', centre)
            worksheet_ng.write('M17', 'N/A', centre)

        print()
        print('Western - Central route Completed!')

#function call
Tx_BB_Core.central_eastern(Self)

Tx_BB_Core.western_central(Self)


print()
print('-'*80)
print('Task Completed!')
print('-'*80)

workbook.close()
