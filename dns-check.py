#!/usr/bin/env python
# encoding: utf-8

import xlrd
import xlwt
import xlutils
import dns.resolver
import argparse

parser = argparse.ArgumentParser(description='args to this script')
parser.add_argument('-s', type=str, default='ns1.alidns.com')
parser.add_argument('-f', type=str)
parser.add_argument('-d', type=str)
args = parser.parse_args()

my_resolver = dns.resolver.Resolver()
my_resolver.nameservers = [args.s]

book = xlrd.open_workbook(args.f)
# print("表单数量： %d" % book.nsheets)
# print("##############3")
# print("表单名称: %s" % book.sheet_names)

sh1 = book.sheet_by_index(0)

#print("%s + %d + %d " % (sh1.name, sh1.nrows, sh1.ncols))

domain_zone = args.d


def getArember(domain):
    recode_value = []
    A = dns.resolver.query(domain, 'A')
    for i in A.response.answer:
        for j in i:
            recode_value.append(unicode(j.to_text().strip('"')))
    return recode_value


def getMXrember(domain):
    recode_value = []
    MX = dns.resolver.query(domain, 'MX')
    for i in MX.response.answer:
        for j in i:
            recode_value.append(
                unicode(j.exchange.to_text().strip('"').upper()))
    return recode_value


def getCnamerember(domain):
    recode_value = []
    cname = dns.resolver.query(domain, 'CNAME')
    for i in cname.response.answer:
        for j in i:
            recode_value.append(unicode(j.to_text().strip('"')))
    return recode_value


def getTXTrember(domain):
    recode_value = []
    txt = dns.resolver.query(domain, "TXT")
    for i in txt.response.answer:
        for j in i:
            recode_value.append(unicode(j.to_text().strip('"')))
    return recode_value


dns_dict = {}
# 从excel中生成去重过的 {域名#type:[解析值]...} 格式的字典
for R in range(1, sh1.nrows):
    if sh1.cell_value(R, 1) == '@':
        domain = domain_zone
        recode_name = domain+'#'+sh1.cell_value(R, 0)
    else:
        domain = sh1.cell_value(R, 1) + '.' + domain_zone
        recode_name = domain+'#'+sh1.cell_value(R, 0)

    if dns_dict.has_key(recode_name):
        dns_dict[recode_name].append(sh1.cell_value(R, 3))
    else:
        dns_dict[recode_name] = [sh1.cell_value(R, 3)]


# 从生成的域名字典中拿到域名去执行dns解析，并且比较解析结果和原本表中的解析值
for dnsname in dns_dict.keys():
    domain = dnsname.split('#')[0]
    dns_type = dnsname.split('#')[1]
    if dns_type == 'A':
        dnsresolve_value = getArember(domain)
    elif dns_type == 'CNAME':
        dnsresolve_value = getCnamerember(domain)
    elif dns_type == 'MX':
        dnsresolve_value = getMXrember(domain)
    elif dns_type == 'TXT':
        dnsresolve_value = getTXTrember(domain)
    else:
        print('Without the type of %s' % dns_type)
        continue
    dnsresolve_value = set(dnsresolve_value)
    dnsFromExecl_value = set(dns_dict[dnsname])
# 通过set的差集方法进行比较，注意要分别判断A差集B，和B差集A的结果。
    if len(dnsresolve_value.difference(dnsFromExecl_value)) == 0 and len(dnsFromExecl_value.difference(dnsresolve_value)) == 0:
        pass
    else:
        print('\033[1;31m %s \033[0m' % dnsname)
        print('Result from dns is %s' % dnsresolve_value)
        print('Result from excel is %s' % dnsFromExecl_value)
