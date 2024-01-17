#!/usr/bin/env python
# -*- coding: utf-8 -*-

# Some parts of the code were inspired (well, copied mercilessly and re-adapted)
# from: https://balasankarc.in/tech/using-python-and-odfpy-to-create-open-document-texts.html

from tkinter import filedialog as fd

import pandas as pd

from odf import teletype, userfield, table, text
from odf.opendocument import OpenDocumentText
from odf.style import (Style, TextProperties, ParagraphProperties,
                       ListLevelProperties, TabStop, TabStops,
                       TableProperties, TableRowProperties,
                       TableColumnProperties, TableCellProperties)
from odf.text import (H, P, List, ListItem, ListStyle, ListLevelStyleNumber,
                      ListLevelStyleBullet)
from odf.table import Table, TableColumn, TableRow, TableCell
# ==================================================================
#
#
#
#
#
# ==================================================================
class CVmaker():

    def __init__(self, fileIn=None,
                textDoc = OpenDocumentText(),
                styles = None,
                configOds = None,
                theOds = None,
                levels = None,
                debug = False):

        # ------------------- Diccionaries.   These will be used in a future GUI.
        self.dict_setters = dict(fileIn = self.set_fileIn,
                textDoc = self.set_textDoc,
                styles = self.set_styles,
                configOds = self.set_configOds,
                theOds = self.set_theOds,
                tidyData = self.set_tidyData,
                levels = self.set_levels,
                tree = self.set_tree)

        self.dict_getters = dict(fileIn = self.get_fileIn,
                textDoc = self.get_textDoc,
                styles = self.get_styles,
                configOds = self.get_configOds,
                theOds = self.get_theOds,
                tidyData = self.get_tidyData,
                levels = self.get_levels,
                tree = self.get_tree)

        # Object attributes.
        self.fileIn = fileIn
        self.textDoc  = textDoc
        self.styles  = styles
        self.theOds  = theOds
        self.tidyData  = dict()
        self.levels  = levels
        self.tree = None
        self.debug = debug
        self.dic = dict()

        return None

    # --------------------------------------------- Setters and getters
    def get_dict_getters(self):
        return self.dict_getters

    def get_dict_setters(self):
        return self.dict_setters


    def set_fileIn(self, fileIn):
        self.fileIn = fileIn
        return None

    def get_fileIn(self):
        return self.fileIn


    def set_textDoc(self, textDoc):
        self.textDoc = textDoc
        return None

    def get_textDoc(self):
        return self.textDoc


    def set_styles(self, styles):
        self.styles = styles
        return None

    def get_styles(self):
        return self.styles


    def set_configOds(self, configOds):
        self.configOds = configOds
        return None

    def get_configOds(self):
        return self.configOds


    def set_theOds(self, theOds):
        self.theOds = theOds
        return None

    def get_theOds(self):
        return self.theOds


    def set_tidyData(self, tidyData):
        self.tidyData = tidyData
        return None

    def get_tidyData(self):
        return self.tidyData


    def set_levels(self, levels):
        self.levels = levels
        return None

    def get_levels(self):
        return self.levels


    def set_tree(self, tree):
        self.tree = tree
        return None

    def get_tree(self):
        return self.tree


    def set_debug(self, debug):
        self.debug = debug
        return None

    def get_debug(self):
        return self.debug
    # ----------------------
    #
    #
    # --------------------------------------------- Methods
    def readOdsFiles(self):

        wasOk = False

        if self.get_fileIn() != None:

            try:
                theOds = pd.read_excel(self.get_fileIn(), engine="odf")
                self.set_theOds(theOds)
                wasOk = True
            except:
                print ("Error reading CV data ods file.")

            if wasOk:
                try:
                    configIn = pd.read_excel("CV-config.ods", engine="odf")
                    config = configIn.transpose()
                    self.set_configOds(config)
                except:
                    print ("Error reading config ods file.")
                    wasOk = False
        else:
            print ("No file chosen to be read.")

        return wasOk
    # ----------------------
    #
    #
    # ----------------------
    def chooseFile(self):

        # choose file dialog.
        response = fd.askopenfilename()

        if response:
            if self.get_debug(): print ('\n\tClicked in ok')

            okStatus = True
            fileIn = response

            if self.get_debug(): print ('fileIn=',fileIn)

            self.set_fileIn(fileIn)
            return None

        else:
            if self.get_debug(): print ('\n\tClicked in cancel')

            return None
    # ----------------------
    #
    #
    # ----------------------
    def buildDictAndcheckForUniqueIds(self):

        theOdsIn = self.get_theOds()

        keys = theOdsIn.keys().tolist()
                # ~ print ("theOdsIn keys:", keys, "\n\n")

        self.set_theOds(theOdsIn.to_dict('index'))

        tidyData = dict()

        for k, v in self.theOds.items():

            if not(v['Id'] in tidyData.keys()):

                key = v['Id']

                del v['Id']

                auxD = dict()
                for ki in v.keys():
                    if ki != 'Id':
                        auxD[ki] = v[ki]

                tidyData[key] = auxD

            else:
                print ('\n\tWarning, Id: ', v['Id'], ' is repeated in CV data spreadsheet.')

        self.set_tidyData(tidyData)

        return None
    # ----------------------
    #
    #
    # ----------------------
    def buildTree(self):

        dataOut = dict()

        if len(self.get_tidyData()) == 0:
            return dataOut

        for key, v in self.get_tidyData().items():

            dad = v['Parent']

            if dad in dataOut.keys():
                dataOut[dad].append(key)

            # asign Id to parent children.
            if not(key in dataOut.keys()):
                dataOut[key] = []

        self.set_tree(dataOut)

        return None
    # ----------------------
    #
    #
    # ----------------------
    def defineStyles(self):

        styles = self.get_textDoc().styles

        # Title style.
        titleStyle = Style(name="Title", family="paragraph")
        titleStyle.addElement(ParagraphProperties(attributes={"textalign": "center"}))
        titleStyle.addElement(TextProperties(
            attributes={"fontsize": "18pt", "fontweight": "bold"}))

        # Subitle style.
        subtitleStyle = Style(name="Subtitle", family="paragraph")
        subtitleStyle.addElement(ParagraphProperties(attributes={"textalign": "center"}))
        subtitleStyle.addElement(TextProperties(
            attributes={"fontsize": "16pt", "fontweight": "bold"}))


        # Section style.
        sectionStyle = Style(name="Section", family="paragraph")
        sectionStyle.addElement(ParagraphProperties(attributes={"textalign": "start"}))
        sectionStyle.addElement(TextProperties(
            attributes={"fontsize": "14pt", "fontweight": "bold"}))


        # Subsection style.
        subsectionStyle = Style(name="Subsection", family="paragraph")
        subsectionStyle.addElement(ParagraphProperties(attributes={"textalign": "start"}))
        subsectionStyle.addElement(TextProperties(
            attributes={"fontsize": "12pt", "fontweight": "bold"}))


        # Bold text style.
        boldStyle = Style(name="Bold", family="text")
        boldStyle.addElement(TextProperties(
            attributes={"fontsize": "12pt", "fontweight": "bold"}))


        # Plain text style.
        plainText = Style(name = "Plain", family = "paragraph",
                         defaultoutlinelevel = "1", liststylename = "Plain")
        plainText.addElement(ParagraphProperties(
                attributes={"textalign": "justify",
                        "marginleft" : "1cm",
                        "marginright" : "0cm",
                        "margintop" : "0cm",
                        "marginbottom" : "0.25cm",
                        "keeptogether" : "always",
                        "textindent" : "0.0cm" }
                                ))
        plainText.addElement(TextProperties(
            attributes={"fontsize": "12pt"}))

        # For references
        reference = Style(name="Refernce", family="text")
        reference.addElement(TextProperties(
            attributes={"fontsize": "9pt", "fontweight":"regular"}))


        levels =   {0: titleStyle,
                    1: subtitleStyle,
                    2: sectionStyle,
                    3: subsectionStyle,
                    4: boldStyle,
                    5: plainText,
                    6: reference
                    }

        self.set_levels(levels)

        # Numbered list
        numberedListStyle = ListStyle(name="NumberedList")
        level = 1
        numberedlistproperty = ListLevelStyleNumber(
            level=str(level), numsuffix=".", startvalue=1)
        numberedlistproperty.setAttribute('numsuffix', ".")
        numberedlistproperty.addElement(ListLevelProperties(
            minlabelwidth="%fcm" % (level - .2)))
        numberedListStyle.addElement(numberedlistproperty)

        # For Bulleted list
        bulletedListStyle = ListStyle(name="BulletList")
        level = 1
        bulletlistproperty = ListLevelStyleBullet(level=str(level), bulletchar=u"•")
        bulletlistproperty.addElement(ListLevelProperties(
            minlabelwidth="%fcm" % level))
        bulletedListStyle.addElement(bulletlistproperty)


        # Justified style
        justifiedStyle = Style(name="justified", family="paragraph")
        justifiedStyle.addElement(ParagraphProperties(attributes={"textalign": "justify"}))


        # Register created styles to styleset
        styles.addElement(titleStyle)
        styles.addElement(subtitleStyle)
        styles.addElement(sectionStyle)
        styles.addElement(subsectionStyle)
        styles.addElement(boldStyle)
        styles.addElement(plainText)
        styles.addElement(numberedListStyle)
        styles.addElement(bulletedListStyle)
        styles.addElement(justifiedStyle)
        # ~ styles.addElement(tabbedParagraphStyle)

        self.set_styles(styles)

        return None
    # ----------------------
    #
    #
    # ----------------------
    def buildCV(self):

        self.chooseFile()
        wasOk = self.readOdsFiles()

        if wasOk:

            self.buildDictAndcheckForUniqueIds()
            self.buildTree()

            if len(self.get_tree()) == 0:
                return None

            # Define styles.
            self.defineStyles()

            startFrom = "Title"
            textdoc = self.addElements (startFrom, level = 0)

        else:
            pass

        return None
    # ----------------------
    #
    #
    # ----------------------
    def addElements(self, startFrom, level):

        textDoc = self.get_textDoc()
        styles = self.get_styles()
        tidyData = self.get_tidyData()
        tree = self.get_tree()
        levels = self.get_levels()

        elementName = tidyData[startFrom]

        if True:

            if level > 2:
                element = H(outlinelevel = 1, stylename = levels[level])
            else:
                element = P(stylename = levels[level])

            # chequear acá por idioma según config.
            text = elementName['English']

            teletype.addTextToElement(element, text)
            textDoc.text.addElement(element)

            children = self.getChildrenIds (startFrom)

            # ~ print ("children:", children)
            level += 1
            for child in children:

                textDoc = self.addElements (child, level)

        return None
    # ----------------------
    #
    #
    # ----------------------
    def getChildrenIds(self, parent):

        children = []

        if parent in self.tree.keys():
            children = self.get_tree()[parent]

        return children
    # ----------------------
    #
    #
    # ----------------------
    def getDataAsDict(self):   # This method is for future session management save/retrieval.

        getters = self.get_dict_getters()

        for theKey in getters.keys(): # returns a list of attributes keys or values.

            attrib = getters[theKey]() # returns the getter value or the getter object.

            if type(attrib) == type([1,2]):

                theList = []
                for item in attrib:

                    try:
                        obj = item.get_data_as_dict()
                        theList.append(obj)

                    except:

                        theList.append(item)

                self.dic[theKey] = theList

            else:

                try:
                    obj = attrib.get_data_as_dict()
                    self.dic[theKey] = obj

                except:
                    self.dic[theKey] = attrib

        return self.dic
    # ----------------------
    #
    #
    # ----------------------

# ========================== Main program.
if __name__ == "__main__":

    cvMaker = CVmaker()

    cvMaker.buildCV()

    textdoc = cvMaker.get_textDoc()

    if textdoc != None:
        textdoc.save(u"CV.odt")
        print ("\n\tSuccess.")


# eof.-