#/***********************************************************************
# * Licensed Materials - Property of IBM 
# *
# * IBM SPSS Products: Statistics Common
# *
# * (C) Copyright IBM Corp. 1989, 2015
# *
# * US Government Users Restricted Rights - Use, duplication or disclosure
# * restricted by GSA ADP Schedule Contract with IBM Corp. 
# ************************************************************************/

# Correlation matrix formatting

__author__ = "SPSS, JKP"
__version__ = "1.1.0"

# history
# 08-feb-2015 add option to bold significant correlations

# function RGB takes a list of three values and returns the RGB value
# function floatex decodes a numeric string value to its float value taking the cell format into account


import SpssClient   # for text constants
from modifytables import RGB
from extension import floatex  # strings to floats
from spssaux import getSpssVersion

ver=[int(v) for v in getSpssVersion().split(".")]
hidelok = ver[0] >= 19 or (ver[0] == 18 and ver[1] > 0 or (ver[1] == 0 and ver[2] >= 3))

# behavior settings
color = RGB((251, 248, 115))      # yellow
BSIZE=3   # number of statistics rows in table
style = SpssClient.SpssTextStyleTypes.SpssTSBold

def cleancorr(obj, i, j, numrows, numcols, section, more, custom):
    """Clean correlation matrix
    
    parameters:
    hideN   - hide count rows
    hideL   - hide statistics label
    lowertri - show only lower triangle
    hideinsig - threshold for hiding insignificant coefs - default .05
    emphlarge - threshold for emphasizing large correlations - default .5
    emphasis - synonym for emphlarge
    decimals - number of decimal places
    boldsig - bold significant correlations
"""
    
    if not (0 <custom.get("hideinsig", .05) <=1.):
        print "The significance threshold for hiding must be between 0 and 1"
        raise ValueError
    if not (0. <= custom.get("emphlarge", .5) <= 1.) or not (0 <= custom.get("emphasis", .5) <= 1.):
        print "The highlighting threshold must be between 0 and 1"
        raise ValueError
    if not (0. <= custom.get("boldsig", .05) <= 1):
        print "Significance threshold for bolding must be between 0 and 1"
    try:
        int(custom.get("decimals", 999))
    except:
        print "Decimal parameter must be an integer"
        raise ValueError
    
    ###debugging (move this code appropriately for repeated debugging)
    #try:
        #import wingdbstub
        #if wingdbstub.debugger != None:
            #import time
            #wingdbstub.debugger.StopDebug()
            #time.sleep(2)
            #wingdbstub.debugger.StartDebug()
        ##import thread
        ##wingdbstub.debugger.SetDebugThreads({thread.get_ident(): 1}, default_policy=0)
    #except:
        #pass
    ###SpssClient._heartBeat(False)
    
    if "emphlarge" in custom:
        emphlarge = custom["emphlarge"]
    elif "emphasis" in custom:
        emphlarge = custom["emphasis"]
    else:
        emphlarge = .5
    
    f = Clean(more.thetable,
              custom.get("hiden", True),
              custom.get("hidel", True),
              custom.get("lowertri", True),
              custom.get("hideinsig", .05),
              emphlarge,
              custom.get("decimals", None),
              custom.get("boldsig", 0.))
    
    while True:
        if  f.cleanblock() is False:
            return False
    
    
class Clean(object):
    def __init__(self, pt, hideN, hideL, lowertri, hideinsig, emphlarge, 
        decimals, boldsig):
        
        attributesFromDict(locals())
        if not self.decimals is None:
            self.decimals = int(self.decimals)
        
        self.datacells = pt.DataCellArray()
        self.dim = self.datacells.GetNumColumns()   # size of a block
        self.numrows = self.datacells.GetNumRows()
        
        self.rowl = pt.RowLabelArray()
        self.lastlbl = self.rowl.GetNumColumns() - 1  # innermost label
  
        self.block = 0
        
    def cleanblock(self):
        """clean one block of a table"""

        blockaddr = self.block * self.dim * BSIZE
        if blockaddr >= self.numrows:    # done

            if self.hideN:
                self.rowl.HideLabelsWithDataAt(2, self.lastlbl)
            if self.hideinsig < 1.:
                self.rowl.HideLabelsWithDataAt(1, self.lastlbl)
            if self.hideL and hidelok:
                self.rowl.SetRowLabelWidthAt(1,self.lastlbl, 0)
                ###self.rowl.SetValueAt(1, self.lastlbl, " ")
            return False        

        for row in range(self.dim):
            rowaddr = blockaddr + row * BSIZE
            if self.lowertri:  # blank cells past dialgonal
                for c in range(row + 1, self.dim):
                    self.datacells.SetValueAt(rowaddr, c, " ")
                    self.datacells.HideFootnotesAt(rowaddr, c)
                    if not self.hideN:
                        self.datacells.SetValueAt(rowaddr+2, c, " ")
                    if not (self.hideinsig < 1.):
                        self.datacells.SetValueAt(rowaddr+1, c, " ")

            # blank insignificant cells (if requested) and highlight large corrs (if requested)
            for c in range(self.dim):   # column loop
                if self.lowertri and c > row:
                    continue
                skip = False
                if self.hideinsig < 1. or self.boldsig < 1.:  # blank insig
                    if c != row:
                        try:
                            sig = floatex(self.datacells.GetValueAt(rowaddr+1, c))
                        except:
                            sig = 1.  # hide if not a number
                        if sig > self.hideinsig:
                            self.datacells.SetValueAt(rowaddr, c, " ")
                            self.datacells.HideFootnotesAt(rowaddr, c)  # remove all footnotes
                            skip = True
                        elif sig <= self.boldsig:
                            self.datacells.SetTextStyleAt(rowaddr, c, 
                                SpssClient.SpssTextStyleTypes.SpssTSBold)
                corr = None
                if self.emphlarge < 1. and (not skip) and c != row:  # highlight large corrs unless failed sig test
                    try:
                        corr = floatex(self.datacells.GetValueAt(rowaddr, c))
                    except:
                        corr = 10.
                    if abs(corr) >= self.emphlarge:
                        self.datacells.SetTextStyleAt(rowaddr, c, style)
                        self.datacells.SetBackgroundColorAt(rowaddr, c, color)
                if (not self.decimals is None) and (not skip):
                    self.datacells.SetHDecDigitsAt(rowaddr, c, self.decimals)
        
        self.block += 1
        
def attributesFromDict(d):
    """build self attributes from a dictionary d."""

    self = d.pop('self')
    for name, value in d.iteritems():
        setattr(self, name, value)
