#!/usr/bin/env python
# -*- coding: utf-8 -*-

__description__ = '[!] emlextractor.py: A tool to extract attachments from EML files.'
__author__ = 'Sam0rai'
__version__ = '1.0'
__date__ = '2017/08/10'


import sys, os
import optparse
import textwrap
import email
from emaildata.attachment import Attachment
import csv


def PrintManual():
    manual = '''
Manual:

emlextractor is a tool to extract attachments from EML files.
The input to the tool can be a single EML file, or a directory where EML files 
are located.
The extracted attachments are saved by default in the folder where the tool 
is located, or can be saved to a designated folder using the "-o" flag. 

[*] Use-case: extract attachment from a single file. 
    Usage: extractor.py -f <full path to file> -o <output directory>
    Example:
        extractor.py -f c:\test\eml\invoice.eml

[*] Use-case: extract attachments from EML files located in a folder.
    Usage: extractor.py -i <directory> -o <output directory>
    Example: 
        extractor.py -i c:\test\eml -o c:\test\output

[*] Use-case: export the mapping of attachments to their originating EML files
    Usage: extractor.py -i <directory> -o <directory> -c MyResults.csv
    Example:
        extractor.py -v -i c:\test\eml -o c:\test\output -c c:\test\MyResults.csv
'''
    for line in manual.split('\n'):
        print(textwrap.fill(line, 78))

# List which holds tuples of attachments names and the EML file which they 
# originated from.
fileMapping = list()

def extractAttachments(options):
    outputDir = ""
    if options.outputdir:
        try:
            outputDir = options.outputdir
            # If output directory does not exist - create it.
            if not os.path.isdir(outputDir):
                os.makedirs(outputDir)
        except Exception as e:
            print "[ERROR]:: " + str(e)
            sys.exit(1)
    
    if options.inputFile:
        f = options.inputFile
        if f.endswith("eml"):
            try:
                with open(f, 'r') as emlPath:
                    message = email.message_from_file(emlPath)
                    for content, filename, mimetype, message in Attachment.extract(message): # @UnusedVariable
                        filename = filename.replace("\n","")
                        if options.verbose:
                            print u"[*] Found attachment '{0}' in file {1}".format(filename, f)
                            
                        OutputPathFilename = os.path.join(outputDir, filename)
                        with open(OutputPathFilename, 'wb') as stream:
                            stream.write(bytearray(content))
                        if(outputDir == ""):
                            print u"Saving attachment to: {0}".format(os.path.join(os.getcwd(),filename))
                        # If message is not None then it is an instance of email.message.Message
                        if message:
                            print "The file {0} is a message with attachments.".format(filename)
                        fileMapping.append((filename, f))
            except IOError as ioe:
                print ("[Error] Could not find file '{0}'.".format(f))
                print ioe            
        
    if options.inputdir:
        try:
            emlDir = options.inputdir
            if not os.path.isdir(emlDir):
                msg = "{0} is not a valid directory.".format(emlDir)
                raise Exception(msg)
             
            for root, dirs, files in os.walk(emlDir):  # @UnusedVariable
                    for f in files:
                        path = os.path.join(root, f)
                        if f.endswith("eml"):
                            with open(path, 'r') as emlPath:
                                message = email.message_from_file(emlPath)
                                for content, filename, mimetype, message in Attachment.extract(message): # @UnusedVariable
                                    #filename = filename.encode('utf-8').replace("\n","")
                                    filename = filename.replace("\n","")
                                    if options.verbose:
                                        print u"[*] Found attachment '{0}' in file {1}".format(filename, f)
                                        
                                    OutputPathFilename = os.path.join(outputDir, filename)
                                    with open(OutputPathFilename, 'wb') as stream:
                                        stream.write(bytearray(content))
                                    if(outputDir == ""):
                                        print u"Saving attachment to: {0}".format(os.path.join(os.getcwd(),filename))
                                    # If message is not None then it is an instance of email.message.Message
                                    fileMapping.append((filename, f))
        except Exception as e:
            print "[ERROR] " + e.message
            sys.exit(1)


def exportToFile(options):
    OutputPathFilename = os.path.join(options.outputdir, options.csv)
    with open(OutputPathFilename, 'w') as csvfile:
        fieldnames = ['Attachment', 'EML_File']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        for (key,val) in fileMapping:
            print key +"," + val
            writer.writerow({'Attachment': key, 'EML_File': val})
        

def Main():
    oParser = optparse.OptionParser(usage='%prog [options] [EML source]\n' + __description__, version='%prog ' + __version__)
    oParser.add_option('-m', '--man', action='store_true', default=False, help='Print manual')
    oParser.add_option('-v', '--verbose', action='store_true', default=False, help='Verbose mode: print to stdout the mappings of attachments to originating EML files')
    oParser.add_option('-f', '--inputFile', type=str, default='', help='Single input EML file')
    oParser.add_option('-i', '--inputdir', type=str, default='', help='Input directory of EML files')
    oParser.add_option('-c', '--csv', type=str, default='', help='Name of CSV file to export mappings of attachments to originating EML files to.')
    oParser.add_option('-o', '--outputdir', type=str, default='', help='Directory to store extracted attachments in.')
    (options, args) = oParser.parse_args()

    if options.man:
            oParser.print_help()
            PrintManual()
            return
    
    if len(args) > 0:
        oParser.print_help()
        #print('')
        #print('  Source code put in the public domain by @sam0rai.')
        return
    else:
        try:
            extractAttachments(options)
        except Exception as e:
            print "[ERROR]" + e.message
            
        
    # Output mappings of attachments to originating EML files to a CSV file  
    if options.csv:
        exportToFile(options)
        
                            
if __name__ == '__main__':
    Main()