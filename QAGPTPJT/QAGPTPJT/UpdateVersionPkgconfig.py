import os
import re
import argparse # Step 1. import argparse

print('Start : Modify ViDi packages version to use Artifact dependencies in teamcity')
print('Get the current directory of python file: ' + os.path.abspath(__file__)) # Get the current directory of python file.

parser = argparse.ArgumentParser() # Step 2. Create parser
parser.add_argument("-v", "--version", default="7.0.0.00000") # Step 3. Register parameter to be got by parser.add_argment()
args = parser.parse_args() # Step 4. Analyze parameters
print('-v=' + args.version)
appliedVersion = args.version

pathPKG = 'packages.config' ## %Workspace%%OS.Path.Separator%QAGPTPJT\QAGPTPJT\packages.config

f = open(pathPKG, 'r')
lines = f.readlines()
allLines = []
##for line in tqdm(lines): ### line = line.strip()
for line in lines:
    allLines.append(line)            
f.close()

newAllLines = []
for content in allLines:
    if(content.find('ViDi')>0 and content.find('version=')>0):                    
        newAllLines.append(re.sub(content[content.find('version=')+9:content.find('version=')+20], appliedVersion, content))        
    else:
        newAllLines.append(content)

f = open(pathPKG, 'w')
for str in newAllLines:
    f.writelines(str)
f.close()

print('End : Complete update ViDi version in packages.config')
