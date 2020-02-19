import os
import xlwt
import xml.etree.ElementTree as ET


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
            dest = infNFe.findall(namespace + 'dest')[0]
            total = infNFe.findall(namespace + 'total')[0]
            ICMSTot = total.findall(namespace + 'ICMSTot')[0]

            cNF = ide.find(namespace + 'cNF').text
            vNF = float(ICMSTot.find(namespace + 'vNF').text)

            if dest.find(namespace + 'CPF') is not None:
                cpf_cnpj = dest.find(namespace + 'CPF').text
            else:
                cpf_cnpj = dest.find(namespace + 'CNPJ').text

            print(
                cNF + '\t' +
                ide.find(namespace + 'nNF').text + '\t' +
                cpf_cnpj + '\t' +
                infNFe.get('Id')[3:47] + '\t' +
                formatarDataHora(ide.find(namespace + 'dhEmi').text) + '\t' +
                ICMSTot.find(namespace + 'vNF').text.replace('.', ','))

            if cNF in nfes:
                nfes[cNF] = nfes[cNF] + 1
            else:
                nfes[cNF] = 1

            if cpf_cnpj in nfes_cpf_cnpj:
                nfes_cpf_cnpj[cpf_cnpj] = nfes_cpf_cnpj[cpf_cnpj] + 1
            else:
                nfes_cpf_cnpj[cpf_cnpj] = 1

            linha += 1

            sheet.write(linha, 0, cNF)
            sheet.write(linha, 1, ide.find(namespace + 'nNF').text)
            sheet.write(linha, 2, cpf_cnpj)
            sheet.write(linha, 3, infNFe.get('Id')[3:47])
            sheet.write(linha, 4, formatarDataHora(
                ide.find(namespace + 'dhEmi').text))
            sheet.write(linha, 5, vNF)


nfes = {}
nfes_cpf_cnpj = {}
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
