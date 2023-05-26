import pandas as pd
from xml.etree import ElementTree
name ='test_input.xlsx'

def serialize(CERDATEt):
    return(str(CERDATEt)[:10])



def read_file(x):
    sheet=pd.read_excel(x,header=None)
    s=pd.DataFrame(sheet)
    filename=s[1][2]
    CERTDATA=ElementTree.Element('CERTDATA')
    FILE=ElementTree.SubElement(CERTDATA,'FILENAME')
    ENVELOPE=ElementTree.SubElement(CERTDATA,'ENVELOPE')
    FILE.text=filename
    for strok in range(5,len(s[0])):
        ECERT=ElementTree.SubElement(ENVELOPE,'ECERT')
        CERTNO=ElementTree.SubElement(ECERT,'CERTNO')
        CERTNOt=s[0][strok]
        CERTNO.text=CERTNOt
        CERDATE=ElementTree.SubElement(ECERT,'CERDATE')
        CERDATEt=s[1][strok]
        CERDATEt=serialize(CERDATEt)
        CERDATE.text=CERDATEt
        STATUSt=s[2][strok]
        STATUS=ElementTree.SubElement(ECERT,'STATUS')
        STATUS.text=STATUSt
        IECt=s[3][strok]
        IEC=ElementTree.SubElement(ECERT,'IEC')
        IEC.text=str(IECt)
        EXPNAMEt='"'+s[4][strok]+'"'
        EXPNAME=ElementTree.SubElement(ECERT,'EXPNAME')
        EXPNAME.text=EXPNAMEt
        BILIDt=s[5][strok]
        BILID=ElementTree.SubElement(ECERT,'BILID')
        BILID.text=BILIDt
        SDATEt=s[6][strok]
        SDATE=ElementTree.SubElement(ECERT,'SDATE')
        SDATE.text=str(SDATEt)[:10]
        SCCt=s[7][strok]
        SCC=ElementTree.SubElement(ECERT,'SCC')
        SCC.text=SCCt  
        SVALUE=ElementTree.SubElement(ECERT,'SVALUE')
        SVALUEt=s[8][strok]
        SVALUE.text=str(SVALUEt)
    tree=ElementTree.ElementTree(CERTDATA)
    tree.write('result.xml')
read_file(name)
