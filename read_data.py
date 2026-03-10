# -*- coding: utf-8 -*-
import base64
import json

from lxml import etree as ET


def read_file(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()

    a_data = {
        '项目信息': dict(root.attrib),
        '招标信息': dict(root.find("招标信息").attrib) if root.find("招标信息") is not None else {},
        '投标信息': dict(root.find("投标信息").attrib) if root.find("投标信息") is not None else {},
        '招标控制价信息': dict(root.find("招标控制价信息").attrib) if root.find("招标控制价信息") is not None else {},
        '单项工程': []
    }

    # 解密
    bid = a_data['投标信息']
    if "加密锁号" in bid:
        bid['加密锁号_解密'] = decrypt_data(bid['加密锁号'], 's')
    if "MAC地址" in bid:
        bid['MAC地址_解密'] = decrypt_data(bid['MAC地址'], 's')

    for danxiang in root.findall("单项工程"):
        dx = dict(danxiang.attrib)
        dx['单位工程'] = []
        a_data['单项工程'].append(dx)

        for danwei in danxiang.findall("单位工程"):
            dw = dict(danwei.attrib)
            dw['分部清单'] = []
            dw['措施清单'] = []
            dx['单位工程'].append(dw)

            # 分部清单
            for qd in danwei.findall(".//清单"):
                q = qd.attrib
                item = {
                    "编码": q.get('编码'),
                    "名称": q.get('名称'),
                    "项目特征": q.get('项目特征'),
                    "单位": q.get('单位'),
                    "数量": q.get('数量'),
                    "综合单价": q.get('综合单价'),
                    "综合合价": q.get('综合合价'),
                    "设备单价": 0.0,
                    "综合单价_含设备": 0.0
                }
                # if q.get("编码") == "040801010001":
                qdfx = qd.find("清单单价分析")
                for fx in qdfx.findall(".//单价分析费用项"):
                    fx_data = fx.attrib
                    if fx_data.get("名称") == "设备费":
                        shebei = fx_data['金额']
                        shuliang = q.get('数量')
                        danjia = round(float(shebei) / float(shuliang), 2)
                        # if float(shebei) != 0:
                        #     print(dx['名称'], dw['名称'], q['名称'], shebei, shuliang, danjia)
                        item['设备单价'] = danjia
                        item['综合单价_含设备'] = float(q.get('综合单价')) + danjia

                dw['分部清单'].append(item)

            # 措施项目（跳过总价措施）
            for cs in danwei.findall(".//措施项目计价表"):
                parent = cs.getparent()
                if parent is not None and parent.get("名称") == "其他总价措施项目":
                    continue
                c = cs.attrib
                dw['措施清单'].append({
                    "编码": c.get('编号'),
                    "名称": c.get('名称'),
                    "项目特征": c.get('项目特征'),
                    "单位": c.get('单位'),
                    "数量": c.get('数量'),
                    "综合单价": c.get('单价'),
                    "综合合价": c.get('合价'),
                })

            # 其他项目 - 优化为列表形式
            other = danwei.find("其他项目")
            if other is not None:
                dw['其他项目'] = {
                    '暂列金额明细': [dict(item.attrib) for item in other.findall(".//暂列金额明细")],
                    '暂估价材料明细': [dict(item.attrib) for item in other.findall(".//暂估价材料明细")],
                    '暂估价设备明细': [dict(item.attrib) for item in other.findall(".//暂估价设备明细")],
                    '专业工程暂估明细': [dict(item.attrib) for item in other.findall(".//专业工程暂估明细")],
                    '计日工项': [dict(item.attrib) for item in other.findall(".//计日工项")],
                    '总承包服务费项': [dict(item.attrib) for item in other.findall(".//总承包服务费项")],
                }
                # 移除空列表（可选）
                dw['其他项目'] = {k: v for k, v in dw['其他项目'].items() if v}

    return a_data


# 加密算法
def encrypt_data(data, key='s'):
    b64_encoded = base64.b64encode(data.encode()).decode()
    xor_result = ''.join(chr(ord(c) ^ ord(key)) for c in b64_encoded)
    final_encoded = base64.b64encode(xor_result.encode()).decode()
    return final_encoded


# 解密算法
def decrypt_data(encrypted_data, key='s'):
    if encrypted_data == "无":
        return "无"
    encrypted_list = encrypted_data.split(',')
    decrypted_list = []
    for enc_data in encrypted_list:
        xor_result = base64.b64decode(enc_data).decode()
        b64_encoded = ''.join(chr(ord(c) ^ ord(key)) for c in xor_result)
        original_data = base64.b64decode(b64_encoded).decode()
        decrypted_list.append(original_data)
    return ','.join(decrypted_list)


# 解密算法
def decrypt_data_until(encrypted_data, key: str):
    if encrypted_data == "无":
        return "无"

    encrypted_list = encrypted_data.split(',')
    decrypted_list = []

    for enc_data in encrypted_list:
        xor_result = base64.b64decode(enc_data).decode()

        b64_encoded = ''.join(chr(ord(c) ^ ord(key)) for c in xor_result)

        original_data = base64.b64decode(b64_encoded).decode()
        decrypted_list.append(original_data)

    return ','.join(decrypted_list)


if __name__ == '__main__':
    # file_path = r"./source/zhongyejiangong.XATB"
    # file_path = r"./source/ceshi.XAZB"
    file_path = r"./source/ceshi_zhaobiao.XAXJ"
    a_data = read_file(file_path)

    for dx in a_data['单项工程']:
        for dw in dx['单位工程']:
            print(dw['其他项目'].keys())

