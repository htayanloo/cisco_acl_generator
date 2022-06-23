import pandas as pd
from netaddr import *

def dispach_address(ips):
    result = []
    for ip_item in ips.split(","):
        if "-" in ip_item:
            # 192.168.1.1-10
            temp_ips = ip_item.split("-")
            temp_ips[0]=temp_ips[0].split(".")
            for item in range(int(temp_ips[0][3]),int(temp_ips[1])+1):
                result.append(f"{temp_ips[0][0]}.{temp_ips[0][1]}.{temp_ips[0][2]}.{item}")
        elif "/" in ip_item :
            net = IPNetwork('192.0.2.16/29')
            for item in net:
                result.append(str(item))

        else:
            result.append(ip_item)

    return result



def action(target_ip,final_result):
        final_result.append(f"sh access-list | i list|{target_ip}_")




final_result =[]


df = pd.read_excel (r'input_2.xlsx')

for index, row in df[0:].iterrows():
    for target_ip in dispach_address(row[0]):
        action(target_ip,final_result)


with open("result_2.txt","w+", encoding = 'utf-8') as f:
    for item in final_result:
        # write each item on a new line
        f.write("%s\n" % item)


