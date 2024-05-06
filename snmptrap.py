from pysnmp.carrier.asynsock.dispatch import AsynsockDispatcher
from pysnmp.carrier.asynsock.dgram import udp, udp6
from pyasn1.codec.ber import decoder
from pysnmp.proto import api
from pysnmp.proto.rfc1905 import VarBind
import re
import datetime
import requests
import json
import openpyxl

# 读取EXCEL中的设备信息，用于与trap内的IP地址进行匹配
wb = openpyxl.load_workbook('wenjian.xlsx')
device_values = []
ws = wb.active
for value in ws.values:
    device_values.append(value)
wb.close()

# 发送消息至企业微信机器人


def send_text(webhook, content, mentioned_list=None, mentioned_mobile_list=None):
    header = {
        "Content-Type": "application/json",
        "Charset": "UTF-8"
    }
    data = {
        "msgtype": "text",
        "text": {
            "content": content, "mentioned_list": mentioned_list, "mentioned_mobile_list": mentioned_mobile_list
        }
    }
    data = json.dumps(data)
    info = requests.post(url=webhook, data=data, headers=header)

# 发送消息至企业微信机器人


def send_md(webhook, content, ip, state, erro1):
    header = {
        "Content-Type": "application/json",
        "Charset": "UTF-8"
    }
    data = {
        "msgtype": "markdown",
        "markdown": {
            "content": "# <font color=\"warning\">" + "{}".format(content) + "</font>" + '\n' +
            "> IP: {}".format(ip) + '\n' +
            "> 端口: {}".format(state) + '\n' +
            "> 告警名称: {}".format(erro1)
        }
    }
    data = json.dumps(data)
    # print(data)
    info = requests.post(url=webhook, data=data, headers=header)


webhook = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=xxxxxxxx"
# 在企业微信中创建机器人，获取机器人代码


def pick(varbind):
    pattern = "name=(.*)[\s\S]*value=(.*)"
    regx = re.compile(pattern)
    matchs = regx.findall(varbind)
    if matchs:
        return matchs[0][0], matchs[0][1]
    else:
        return False


def handle_trap(transportDispatcher, transportDomain, transportAddress, wholeMsg):
    # print(wholeMsg)
    while wholeMsg:
        try:
            msgVer = int(api.decodeMessageVersion(wholeMsg))
        except:
            msgVer = 'null'
        if msgVer in api.protoModules:
            pMod = api.protoModules[msgVer]
        else:
            print('Unsupported SNMP version %s' % msgVer)
            return
        reqMsg, wholeMsg = decoder.decode(wholeMsg, asn1Spec=pMod.Message(), )
        print('Notification message from %s:%s: ' %
              (transportDomain, transportAddress))
        reqPDU = pMod.apiMessage.getPDU(reqMsg)
        varBinds = pMod.apiPDU.getVarBindList(reqPDU)
        mib_oid_value = ''
        for row in varBinds:
            row: VarBind
            row = row.prettyPrint()
            k, v = pick(row)
            mib_oid_value += (k + ":" + v + "\n")
        ip = transportAddress[0]  # 提取IP地址
        alarm_state = ''
        alarm_check = re.findall(
            '3.1.12.\d+', mib_oid_value, re.MULTILINE)  # 获取告警MIB
        # print(alarm_check )
        if alarm_check:
            if int(alarm_check[0].replace('3.1.12:', '')) == 6:
                alarm_state = '告警消除'
            else:
                alarm_state = '发生告警'
        # print(alarm_state)

        sn_number = ''
        sn_number = re.findall('10.3.1.13:\.+', mib_oid_value, re.MULTILINE)
        if sn_number:
            sn_number = sn_number[0]
        device_name = 'null'
        device_ip = 'null'
        for i in device_values:
            if sn_number in i:
                device_name = i[0]
                device_ip = i[1]

        alarm = (re.findall('\[.+\]', mib_oid_value))
        if alarm:
            alarm = alarm[0]
        else:
            alarm = 'null'
        # print(alarm)
        # send_md(webhook, content='提示', ip=ip, state=alarm_state, erro1=alarm)
        shijian = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        send_text(webhook, content=('时间：{}\n网元：{}\nIP地址：{}\n消息类型：{}\n'
                                    '告警名称：{}'.format(shijian, device_name, device_ip, alarm_state, alarm)),
                  mentioned_mobile_list=None)
    # except Exception as e:
    #     print(e)

    return wholeMsg


transportDispatcher = AsynsockDispatcher()
transportDispatcher.registerRecvCbFun(handle_trap)
# UDP/IPv4
transportDispatcher.registerTransport(
    udp.domainName, udp.UdpSocketTransport().openServerMode(('0.0.0.0', 16667))
)
print('开始监听IPV4=0.0.0.0:16667')
# UDP/IPv6
transportDispatcher.registerTransport(
    udp6.domainName, udp6.Udp6SocketTransport().openServerMode(('::1', 16666))
)
print('开始监听IPV6=::1:16666')
transportDispatcher.jobStarted(1)
try:
    transportDispatcher.runDispatcher()
except:
    transportDispatcher.closeDispatcher()
    raise
