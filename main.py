import pandas as pd
from netaddr import *

def dispach_address(ips):
    result = []
    if "-" in ips:
        # 192.168.1.1-10
        temp_ips = ips.split("-")
        temp_ips[0]=temp_ips[0].split(".")
        for item in range(int(temp_ips[0][3]),int(temp_ips[1])+1):
            result.append(f"{temp_ips[0][0]}.{temp_ips[0][1]}.{temp_ips[0][2]}.{item}")
    elif "/" in ips :
        net = IPNetwork('192.0.2.16/29')
        for item in net:
            result.append(str(item))

    else:
        result.append(ips)

    return result

def dispach_ports(ports):
    result = []
    for item in ports.split(","):
        if "-" in item:
            for port in range(int(item.split("-")[0]),int(item.split("-")[1])+1):
                result.append(port)
        else:
            result.append(item)
    
    return result


def action(src_ip,dst_ip,port,action,protocol,final_result_src_to_dst,final_result_dst_to_src):
        final_result_src_to_dst.append(f"{action} {protocol} host {src_ip} host {dst_ip} eq {port}")
        final_result_dst_to_src.append(f"{action} {protocol} host {dst_ip} eq {port} host {src_ip} ")



final_result_src_to_dst =[]
final_result_dst_to_src =[]

df = pd.read_excel (r'input.xlsx')

for index, row in df[0:].iterrows():
    for src_ip in dispach_address(row[0]):
        for dst_ip in dispach_address(row[1]):
            for port in dispach_ports(row[2]):
                action(src_ip,dst_ip,port,row[3],row[4],final_result_src_to_dst,final_result_dst_to_src)


with open("result.txt","w+", encoding = 'utf-8') as f:
    for item in final_result_src_to_dst:
        # write each item on a new line
        f.write("%s\n" % item)
    for item in final_result_dst_to_src:
        # write each item on a new line
        f.write("%s\n" % item)


