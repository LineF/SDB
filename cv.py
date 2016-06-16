#!/usr/bin/python

import sys, argparse, io
import xml.dom.minidom

def cvt(ifile, ofile):
    DOMTree = xml.dom.minidom.parse(ifile)
    collection = DOMTree.documentElement

    if collection.hasAttribute("CheckSum"):
        print ("Root element : %s" % collection.getAttribute("CheckSum"))

    elems = collection.getElementsByTagName("")
    for elem in elems:
        print ("Tag")
        if elem.hasAttribute("title"):
            print ("Title: %s" % elem.getAttribute("title"))
        type = elem.getElementsByTagName('type')[0]
        print ("Type: %s" % type.childNodes[0].data)

    
def main(argv):
    parser = argparse.ArgumentParser(description='XML converter')
    parser.add_argument('-i', '--input', required=True)
    parser.add_argument('-o', '--output')
    parser.add_argument('-v', dest='verbose', action='store_true')
    args = parser.parse_args()

    ifile = args.input
    ofile = args.output
    if ofile == None :
        ofile = ifile + ".out"
        
    print ('Input file is "', ifile, '"', sep="")
    print ('Output file is "', ofile, '"', sep="")

    cvt(ifile, ofile)
    

if __name__ == "__main__":
   main(sys.argv[1:])
