#!/usr/bin/python

import sys, argparse, io

def cvt(ifile, ofile):
    file = open(ifile, "r", encoding="utf-8")
    print (file.read())
    file.close()

    
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
