import os
import re
print('Start : Modify ViDi packages version to use Artifact dependencies in teamcity')
print(os.path.abspath(__file__)) # Get the current directory
## %Workspace%%OS.Path.Separator%QAGPTPJT\QAGPTPJT\packages.config
#pathPKG = '%Workspace%%OS.Path.Separator%QAGPTPJT\QAGPTPJT\packages.config'

#pathPKG = 'H:\\Test_PythonVPDLVersion\\packages.config'
pathPKG = 'packages.config'

appliedVersion = "2.2.2.22222"

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
print('End : Complete update ViDi version in  packages.config')
