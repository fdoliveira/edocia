import os
import xlwt
import xml.etree.ElementTree as ET
# import pandas as pd


def formatarDataHora(dataHora):
    ano = dataHora[:4]
    mes = dataHora[5:7]
    dia = dataHora[8:10]
    hor = dataHora[11:13]
    min = dataHora[14:16]
    seg = dataHora[17:19]

    return dia + '/' + mes + '/' + ano + ' ' + \
        hor + ':' + min + ':' + seg


def genSheet1(sheet, files, nfes):

    linha = 0
    for file in files:
        filename = xml_path + file
        tree = ET.parse(filename)
        root = tree.getroot()

        namespace = '{http://www.portalfiscal.inf.br/nfe}'

        if root.tag != namespace + 'nfeProc':
            print('Nao achou')
        else:
            NFe = root.findall(namespace + 'NFe')[0]
            infNFe = NFe.findall(namespace + 'infNFe')[0]
            ide = infNFe.findall(namespace + 'ide')[0]
            emit = infNFe.findall(namespace + 'emit')[0]
            pos = 0
            hasNext = True
            while hasNext:
                det = infNFe.findall(namespace + 'det')[pos]
                prod = det.findall(namespace + 'prod')[0]
                imp = det.findall(namespace + 'imposto')[0]
                icms = imp.findall(namespace + 'ICMS')[0]
                icms_sub = icms[0]

                # children
                CRT = emit.find(namespace + 'CRT').text
                ncm = prod.find(namespace + 'NCM').text
                cfop = prod.find(namespace + 'CFOP').text
                xprod = prod.find(namespace + 'xProd').text

                picms_cst = 0
                if icms_sub.find(namespace + 'pICMS') is not None:
                    picms_cst = float(icms_sub.find(namespace + 'pICMS').text)

                linha += 1

                sheet.write(linha, 0, CRT)
                sheet.write(linha, 1, ncm)
                sheet.write(linha, 2, cfop)
                sheet.write(linha, 3, xprod)
                sheet.write(linha, 4, picms_cst)
                sheet.write(linha, 5, formatarDataHora(
                    ide.find(namespace + 'dhEmi').text))

                print(CRT + '\t' + ncm + '\t' + cfop + '\t' + xprod + '\t' +
                      '\t' + str(picms_cst) + '\t' + formatarDataHora(
                        ide.find(namespace + 'dhEmi').text))

                pos = pos + 1
                try:
                    infNFe.findall(namespace + 'det')[pos]
                    hasNext = True
                except Exception:
                    hasNext = False


nfes = {}

path = os.path.dirname(os.path.realpath(__file__))
xml_path = path + '/xml/'
files = [x for x in os.listdir(xml_path) if x.endswith('nfe.xml')]
if len(files) == 0:
    files = [x for x in os.listdir(xml_path) if x.endswith('nfce.xml')]
linha = 0
workbook = xlwt.Workbook(encoding="utf-8")
sheet1 = workbook.add_sheet("nfe")


genSheet1(sheet1, files, nfes)


workbook.save("edoc.xls")
