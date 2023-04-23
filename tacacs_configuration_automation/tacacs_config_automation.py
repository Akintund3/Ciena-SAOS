from getpass import getpass
from operator import ne
from netmiko import ConnectHandler, ConnectionException, NetMikoAuthenticationException, NetMikoTimeoutException


username = input('Enter Username: ')
password = getpass()


with open('tacacs_ping.txt', 'r') as file:
    t_ping = file.readlines()

with open('tacacs_cet.txt', 'r') as file:
    cet = file.readlines()

with open('tacacs_config_gen.txt', 'r') as file:
    t_gen = file.readlines()

with open('tacacs_saos6.txt', 'r') as file:
    s6 = file.readlines()

with open('tacacs_saos8.txt', 'r') as file:
    s8 = file.readlines()



for i in cet: 
    
    for k in range(len(cet)):

        metro_cet = {
            'device_type': 'ciena_saos',
            'host': i,
            'username': username,
            'password': password,
            }

        try:

            net_ssh = ConnectHandler(**metro_cet)    
            out_ping = net_ssh.send_config_set(t_ping)
            print()
            print('-'*75)
            print(f'Connectinng to device {i}')
            print()
            print('-'*75)
            print()


            if '3 packets transmitted, 3 packets received, 0% packet loss' in out_ping : 

                #generic tacacs config for saos6 and saos8
                
                net_ssh.send_config_set(t_gen)
                tacacs_sh = net_ssh.send_command('tacacs show')

                if '10.0.140.200' and '10.0.140.6' and '10.0.140.9' in tacacs_sh:

                    saos_v = net_ssh.send_command('software show')
                    
                    #tacacs config for saos8

                    if 'rel_saos5170_8' in saos_v:
                        net_ssh.send_config_set(s8)
                        tacacs_sh = net_ssh.send_command('tacacs show')
                        print(tacacs_sh)
                        
                    #tacacs config for saos6

                    else :
                        net_ssh.send_command('tacacs set preferred-source-ip ' + i)
                        net_ssh.send_config_set(s6)
                        tacacs_sh = net_ssh.send_command('tacacs show')
                        print(tacacs_sh)

                else :
                    print(f'A least one of the three tacacs severs not configured on {i}')

            else :

                print(f'{i} can not reach at least one of the three tacacs severs')


        except ConnectionError or NetMikoTimeoutException or NetMikoAuthenticationException or ConnectionException :
            print('Ne Timed out or authentication failure or connection error')

        
