# -*- coding: utf-8 -*-
import xml.sax
import os
import xlsxwriter

class nmapHander(xml.sax.ContentHandler):
 
    def __init__(self,filename):  
        self.output=filename
        self.ipinfo=''
        self.portinfo=''
        self.portstate=0
        self.servname=''
        self.state=''
        self.format = work.add_format({
            #'align': 'center',
            'valign': 'vcenter',
        })
        self.startl=0
        self.lines=1
        xml.sax.handler.ContentHandler.__init__(self)
        self.output.set_column('A:A',15)
        self.output.write('A1','IP')
        self.output.write('B1', 'PORT')
        self.output.write('C1', 'STATE')
        self.output.write('D1', 'SERVICE')

        
    def startDocument(self):
        self.lines = self.lines + 1
        print "Start handler documtnet for "+xmlFile

    def endDocument(self):
        if self.portstate <> 0:
            try:
                if self.startl <> self.lines - 1:
                    self.output.merge_range('A' + str(self.startl) + ':A' + str(self.lines - 1), self.ipinfo, self.format)
                else:
                    self.output.write('A' + str(self.startl), self.ipinfo)
            except:
                print 'Merged Error'
        print "End handler document for "+xmlFile+'\n'

    def startElement(self,name,attrs):
        if name=='address':
            if self.startl <> 0 and self.portstate <> 0:
                try:
                    self.portstate=0
                    if self.startl <> self.lines-1:
                        self.output.merge_range('A' + str(self.startl) + ':A' + str(self.lines-1), self.ipinfo, self.format)
                    else:
                        self.output.write('A' + str(self.startl), self.ipinfo)
                except:
                    print 'Merged Error'
            if attrs.__len__() > 0:
                attr=attrs.getNames()
                self.ipinfo=attrs.getValue(attr[1])
                self.startl=self.lines
            else:  
                print name,"节点不包含属性"

        elif name=='port':
            if attrs.__len__() > 0:
                self.portstate=1
                self.portinfo=attrs.getValue('portid')
                self.output.write('B'+str(self.lines),self.portinfo)
        elif name=='state':
            if attrs.__len__() > 0:
                self.state=attrs.getValue('state')
                self.output.write('C'+str(self.lines),self.state)

        elif name=='service':
            if attrs.__len__() > 0:
                self.servname=attrs.getValue('name')
                self.output.write('D'+str(self.lines),self.servname)
                self.lines += 1

        else: return

    def endElement(self, name):
        xml.sax.ContentHandler.endElement(self, name)


    def characters(self, content):
        xml.sax.ContentHandler.characters(self, content)

if __name__=='__main__':
    files=os.listdir(os.getcwd())
    for xmlFile in files:
        if xmlFile.find('.xml') <> -1:
            work=xlsxwriter.Workbook(xmlFile.strip('.xml')+'.xlsx')
            output=work.add_worksheet()
            parser=xml.sax.make_parser()
            parser.setFeature(xml.sax.handler.feature_namespaces, 0)
            hander=nmapHander(output)
            parser.setContentHandler(hander)
            parser.parse(xmlFile)


            work.close()



















