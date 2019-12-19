"""
Read praat textgrid files connected with the JASMIN corpus

This version created by Erwin R. Komen 
Date: 19/dec/2019 
"""
import sys, getopt, os.path, importlib
import os, sys
import csv, json

# Application specific
import util                 # This allows using ErrHandle
from convert import transform_textgrids

errHandle = util.ErrHandle()

# ----------------------------------------------------------------------------------
# Name :    main
# Goal :    Main body of the function
# History:
# 19/dec/2018    ERK Created
# ----------------------------------------------------------------------------------
def main(prgName, argv):
    dirInput = ''   # input directory
    dirOutput = ''  # output directory
    bForce = False  # Force means: overwrite
    debug = None    # Debugging

    try:
        sSyntax = prgName + ' -i <input file> -o <output directory> [-f] [-d <level>]'
        # get all the arguments
        try:
            # Get arguments and options
            opts, args = getopt.getopt(argv, "hi:o:fd:", ["-help", "-inputdir=", "-outputdir=", "-force","-debug="])
        except getopt.GetoptError:
            print(sSyntax)
            sys.exit(2)
        # Walk all the arguments
        for opt, arg in opts:
            if opt in ("-h", "--help"):
                print(sSyntax)
                sys.exit(0)
            elif opt in ("-i", "--inputdir"):
                dirInput = arg
            elif opt in ("-o", "--outputdir"):
                dirOutput = arg
            elif opt in ("-d", "--debug"):
                try:
                    debug = int(arg)
                except:
                    debug = 10
            elif opt in ("-f", "--force"):
                bForce = True

        # Check if all arguments are there
        if (dirInput == '' or dirOutput == ""):
            errHandle.DoError(sSyntax)
            return False

        # Check if directories exists
        if not os.path.exists(dirInput):
            errHandle.DoError("Input directory does not exist", True)
        if not os.path.exists(dirOutput):
            errHandle.DoError("Output directory does not exist", True)

        # Continue with the program
        errHandle.Status('Input is "' + dirInput + '"')
        errHandle.Status('Output is "' + dirOutput + '"')

        # Call the function that does the job
        oArgs = {'input': dirInput,
                 'output': dirOutput,
                 'force': bForce,
                 'debug': debug}
        if (not transform_textgrids(oArgs, errHandle)) :
            errHandle.DoError("Could not complete")
            return False
    
            # All went fine  
        errHandle.Status("Ready")
        return True
    except:
        # act
        errHandle.DoError("main")
        return False




# ----------------------------------------------------------------------------------
# Goal :  If user calls this as main, then follow up on it
# ----------------------------------------------------------------------------------
if __name__ == "__main__":
    # Call the main function with two arguments: program name + remainder
    main(sys.argv[0], sys.argv[1:])
