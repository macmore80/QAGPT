import os
import re

# Step 1. import argparse
import argparse
# Step 2. Create parser
parser = argparse.ArgumentParser()
# Step 3. Register parameter to be got by parser.add_argment()
parser.add_argument("-v", "--version", default="7.0.0.00000")
# Step 4. Analyze parameters
args = parser.parse_args()
print('-v=' + args.version)
appliedVersion = args.version

print('Start : Modify ViDi packages version to use Artifact dependencies in teamcity')
print(os.path.abspath(__file__)) # Get the current directory of python file.
## %Workspace%%OS.Path.Separator%QAGPTPJT\QAGPTPJT\packages.config
#pathPKG = '%Workspace%%OS.Path.Separator%QAGPTPJT\QAGPTPJT\packages.config'

#pathPKG = 'H:\\Test_PythonVPDLVersion\\packages.config'
pathPKG = 'packages.config'
#appliedVersion = "2.2.2.22222"

pathUpdatePKG = 'updatepackages.config'

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

# f = open(pathPKG, 'w')
# for str in newAllLines:
#     f.writelines(str)
# f.close()
f = open(pathUpdatePKG, 'w')
for str in newAllLines:
    f.writelines(str)
f.close()

print('End : Complete update ViDi version in  packages.config')
