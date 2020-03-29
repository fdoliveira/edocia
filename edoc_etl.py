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
            emit = infNFe.findall(namespace + 'emit')[0]
            pos = 0
            hasNext = True
            while hasNext:
                det = infNFe.findall(namespace + 'det')[pos]
                prod = det.findall(namespace + 'prod')[0]
                imp = det.findall(namespace + 'imposto')[0]
                icms = imp.findall(namespace + 'ICMS')[0]
                icms_sub = icms[0]
                pis = 0
                pis = imp.findall(namespace + 'PIS')[0]
                pis_sub = pis[0]
                cofins = imp.findall(namespace + 'COFINS')[0]
                cofins_sub = cofins[0]
                endmit = emit.findall(namespace + 'enderEmit')[0]

                # children
                CRT = emit.find(namespace + 'CRT').text
                uf = ide.find(namespace + 'cUF').text
                ufend = endmit.find(namespace + 'UF').text
                nf = ide.find(namespace + 'cNF').text
                mod = ide.find(namespace + 'mod').text
                cnpj = emit.find(namespace + 'CNPJ').text
                nome = emit.find(namespace + 'xNome').text
                fant = emit.find(namespace + 'xFant').text
                ncm = prod.find(namespace + 'NCM').text
                cfop = prod.find(namespace + 'CFOP').text
                cprod = prod.find(namespace + 'cProd').text
                xprod = prod.find(namespace + 'xProd').text
                id = infNFe.get('Id')[3:47]

                picms = 0
                if icms_sub.find(namespace + 'pICMS') is not None:
                    picms = float(icms_sub.find(namespace + 'pICMS').text)

                cst_icms = 1
                if icms_sub.find(namespace + 'CST') is not None:
                    cst_icms = int(icms_sub.find(namespace + 'CST').text)
                
                cst_pis = 1
                if pis_sub.find(namespace + 'CST') is not None:
                    cst_pis = int(pis_sub.find(namespace + 'CST').text)

                cst_cofins = 1
                if cofins_sub.find(namespace + 'CST') is not None:
                    cst_cofins = int(cofins_sub.find(namespace + 'CST').text)
                
                csosn = 1
                if icms_sub.find(namespace + 'CSOSN') is not None:
                    csosn = int(icms_sub.find(namespace + 'CSOSN').text)

                ppis = 0
                if pis_sub.find(namespace + 'pPIS') is not None:
                    ppis = float(pis_sub.find(namespace + 'pPIS').text)

                pcofins = 0
                if cofins_sub.find(namespace + 'pCOFINS') is not None:
                    pcofins = float(cofins_sub.find(namespace + 'pCOFINS').text)

                fant = 0
                if emit.find(namespace + 'xFant') is not None:
                    fant = (emit.find(namespace + 'xFant').text)

                linha += 1

                sheet.write(linha, 0, CRT)
                sheet.write(linha, 1, uf)
                sheet.write (linha, 2, ufend)
                sheet.write(linha, 3, nf)
                sheet.write (linha, 4, mod)
                sheet.write (linha, 5, cnpj)
                sheet.write(linha, 6, nome)
                sheet.write(linha, 7, fant)
                sheet.write(linha, 8, ncm)
                sheet.write(linha, 9, cfop)
                sheet.write(linha, 10, cprod)
                sheet.write(linha, 11, xprod)
                sheet.write(linha, 12, picms)
                sheet.write(linha, 13, cst_icms)
                sheet.write(linha, 14, csosn)
                sheet.write(linha, 15, ppis)
                sheet.write (linha, 16, cst_pis)
                sheet.write(linha, 17, pcofins)
                sheet.write (linha, 18, cst_cofins)
                sheet.write(linha, 19, formatarDataHora(ide.find(namespace + 'dhEmi').text))
                sheet.write (linha, 20, id)

                print(CRT + '\t' + ncm + '\t' + id)

                pos = pos + 1
                try:
                    infNFe.findall(namespace + 'det')[pos]
                    hasNext = True
                except Exception:
                    hasNext = False


nfes = {}

path = os.path.dirname(os.path.realpath(__file__))
xml_path = path + '/xml/'
files = [x for x in os.listdir(xml_path) if x.endswith('.xml')]
if len(files) == 0:
    files = [x for x in os.listdir(xml_path) if x.endswith('nfe.xml')]
linha = 0
workbook = xlwt.Workbook(encoding="utf-8")
sheet1 = workbook.add_sheet("nfe")


genSheet1(sheet1, files, nfes)


workbook.save("edoc.xls")
