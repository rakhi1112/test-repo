#imports XML Tree Python Library
import xml.etree.ElementTree as ET
import pandas as pd

#parses the XML file into a tree - the path can be automated - this path is just for testing
tree = ET.parse('_INPUT/RI_XML.xml')
#gets the initial root of the tree
root = tree.getroot()
#break down the XML tree into NFe, infNFe, and Signature sections
if root[0].tag=='{http://www.portalfiscal.inf.br/nfe}infNFe':
    infNFe = root[0]
    Signature = root[1]
else:
    NFe = root[0]
    infNFe = NFe[0]
    Signature = NFe[1]

#read in the mapping spreadsheet
mappingDF = pd.read_excel('_COLUMNS/RenameDict.xlsx')
mappingDict = dict(zip(mappingDF['XML_NAME'], mappingDF['COLUMN_NAME']))


#creates a common Dictionary for all data elements needed for all items
commonDict = {}
#ceates an item Dictionary for data elements specific to items
itemDict = {}

#loops through all the branches within the infNFe branch
for child in infNFe:
    #variables for future logic
    listName = 1
    attDict = {}
    branchType = 'common'
    
    #loops through all the data elements within the child branch
    for elem in child.iter():
        #if the branch is an item branch appen to the itemDict as a new dictinionary, otherwise append to commonDict
        if listName == 1:
            if not elem.attrib:
                branchTag = elem.tag
                #commonDict[dictTitle] = {}
            else:
                branchTag = elem.tag
                dictTitle = elem.tag, elem.attrib['nItem']
                itemDict[dictTitle] = {}
                branchType = 'item'
            listName = 0
        else:
            colStr = str(branchTag)+":"+str(elem.tag)
            colStr = colStr.replace('{http://www.portalfiscal.inf.br/nfe}','')
            attDict[colStr]=(elem.text)
    if branchType == 'common':
        commonDict.update(attDict)
    else:
        itemDict[dictTitle]=attDict
        
        
#appends commond dict to each dict within the item dict
for name,nestedDict in itemDict.items():
    nestedDict.update(commonDict)
    
#create a dataframe out of the itemDict
df = pd.DataFrame.from_dict(itemDict,orient='index')

#Add the data variables from the infNFe branch to the dataframe
df['infNFe:Id']=infNFe.attrib['Id']
df['infNFe:versao']=infNFe.attrib['versao']

#Add the data variables for vBC/CST for ICMS/IPI/PIS/COFINS (duplicate element name in same branch so they get overwritten)
#first drop the existing det:vBC and det:CST
df.drop(columns=['det:vBC','det:CST'])
#create a list to add to the columns
detICMSvBC=[]
detICMSCST=[]
detIPIvBC=[]
detIPICST=[]
detPISvBC=[]
detPISCST=[]
detCOFINSvBC=[]
detCOFINSCST=[]
itemCount=0
#finds the 'det' branches and loops through them to find the tax information for each item
for child in infNFe:
    #only continues to loop through 'det' branches
    if child.tag[-3:] == 'det':
        #counts total number of items
        itemCount+=1
        
        #initially we set the tax values to FALSE - as we are assuming there is no tax value for each item
        ICMSvBC = False
        ICMSCST = False
        IPIvBC = False
        IPICST = False
        PISvBC = False
        PISCST = False
        COFINSvBC = False
        COFINSCST = False
        
        #loop through tax portion of each item
        for taxPart in child[1]:
            tax = taxPart.tag
            
            #checks for vBC and CST values in each tax branch
            for elem in taxPart.iter():
                if ('ICMS' in tax) & ('vBC' in elem.tag):
                    detICMSvBC.append(elem.text)
                    ICMSvBC = True
                elif ('ICMS' in tax) & ('CST' in elem.tag):
                    detICMSCST.append(elem.text)
                    ICMSCST = True
                if ('IPI' in tax) & ('vBC' in elem.tag):
                    detIPIvBC.append(elem.text)
                    IPIvBC = True
                elif ('IPI' in tax) & ('CST' in elem.tag):
                    detIPICST.append(elem.text)
                    IPICST = True
                if ('PIS' in tax) & ('vBC' in elem.tag):
                    detPISvBC.append(elem.text)
                    PISvBC = True
                elif ('PIS' in tax) & ('CST' in elem.tag):
                    detPISCST.append(elem.text)
                    PISCST = True
                if ('COFINS' in tax) & ('vBC' in elem.tag):
                    detCOFINSvBC.append(elem.text)
                    COFINSvBC = True
                elif ('COFINS' in tax) & ('CST' in elem.tag):
                    detCOFINSCST.append(elem.text)
                    COFINSCST = True
        
        #appends NONE values to item tax list (so that we have the correct length of the list to append to the DataFrame)
        if not ICMSvBC:
            detICMSvBC.append(None)
        if not ICMSCST:
            detICMSCST.append(None)
        if not IPIvBC:
            detIPIvBC.append(None)
        if not IPICST:
            detIPICST.append(None)  
        if not PISvBC:
            detPISvBC.append(None)
        if not PISCST:
            detPISCST.append(None)
        if not COFINSvBC:
            detCOFINSvBC.append(None)
        if not COFINSCST:
            detCOFINSCST.append(None)

#append tax columns to the DataFrame
df['det:ICMS:vBC']=detICMSvBC
df['det:ICMS:CST']=detICMSCST
df['det:IPI:vBC']=detIPIvBC
df['det:IPI:CST']=detIPICST
df['det:PIS:vBC']=detPISvBC
df['det:PIS:CST']=detPISCST
df['det:COFINS:vBC']=detCOFINSvBC
df['det:COFINS:CST']=detCOFINSCST

#rename columns as per mapping dictionary
df = df.rename(columns=mappingDict)

#reorders and keeps only columns in the mappingDict
orderList = mappingDF['COLUMN_NAME'].tolist()
#df = df[orderList]
df = df.reindex(columns=orderList)

#create a csv out of the dataframe - the path can be automated - this path is just for testing
df.to_csv('_OUTPUT/Output_XML.csv',na_rep='')