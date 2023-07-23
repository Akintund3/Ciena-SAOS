from getpass import getpass
from netmiko import ConnectHandler, NetMikoAuthenticationException, NetMikoTimeoutException

def connect_to_device(device_type, host, username, password):
    try:
        net_ssh = ConnectHandler(
            device_type=device_type,
            host=host,
            username=username,
            password=password,
        )
        return net_ssh
    except (NetMikoAuthenticationException, NetMikoTimeoutException) as e:
        print(f"Failed to connect to {host}: {e}")
        return None

def check_tacacs_servers(net_ssh, servers):
    tacacs_sh = net_ssh.send_command('tacacs show')
    return all(server in tacacs_sh for server in servers)

if __name__ == "__main__":
    username = input('Enter Username: ')
    password = getpass()

    with open('tacacs_ping.txt', 'r') as file:
        t_ping = file.read()

    with open('tacacs_cet.txt', 'r') as file:
        cet = file.readlines()

    with open('tacacs_config_gen.txt', 'r') as file:
        t_gen = file.read()

    with open('tacacs_saos6.txt', 'r') as file:
        s6 = file.read()

    with open('tacacs_saos8.txt', 'r') as file:
        s8 = file.read()

    tacacs_servers = ['zz.zz.zz.zz', 'xx.xx.xx.xx', 'yy.yy.yy.yy']

    for i in cet:
        host = i.strip()
        metro_cet = {
            'device_type': 'ciena_saos',
            'host': host,
            'username': username,
            'password': password,
        }

        net_ssh = connect_to_device(**metro_cet)
        if net_ssh:
            print(f"\n{'-'*75}\nConnecting to device {host}\n{'-'*75}\n")
            out_ping = net_ssh.send_config_set(t_ping)

            if '3 packets transmitted, 3 packets received, 0% packet loss' in out_ping:
                net_ssh.send_config_set(t_gen)

                if check_tacacs_servers(net_ssh, tacacs_servers):
                    saos_v = net_ssh.send_command('software show')
                    if 'rel_saos5170_8' in saos_v:
                        net_ssh.send_config_set(s8)
                    else:
                        net_ssh.send_command(f'tacacs set preferred-source-ip {host}')
                        net_ssh.send_config_set(s6)

                    tacacs_sh = net_ssh.send_command('tacacs show')
                    print(tacacs_sh)
                else:
                    print(f"At least one of the three tacacs servers is not configured on {host}")
            else:
                print(f"{host} cannot reach at least one of the three tacacs servers")
        else:
            print(f"Failed to connect to {host}")
