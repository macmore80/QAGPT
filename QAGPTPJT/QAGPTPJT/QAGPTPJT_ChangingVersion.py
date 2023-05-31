import os
import re
import argparse # Step 1. import argparse
import numpy as np
import xml.etree.ElementTree as ET

print('Start : Modify ViDi packages version to use Artifact dependencies in teamcity')
print('Get the current directory of python file: ' + os.path.abspath(__file__)) # Get the current directory of python file.

parser = argparse.ArgumentParser() # Step 2. Create parser
parser.add_argument("-v", "--version", default="7.0.0.00000") # Step 3. Register parameter to be got by parser.add_argment()
#parser.add_argument("-v", "--version", default="7.0.1.27813") # Step 3. Register parameter to be got by parser.add_argment()
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

# print("0. Get Version info. \n")
# parser = argparse.ArgumentParser() # Step 2. Create parser
# parser.add_argument("-v", "--version", default="7.0.0.00000") # Step 3. Register parameter to be got by parser.add_argment()
# ##parser.add_argument("-v", "--version", default="7.0.1.27813") # Step 3. Register parameter to be got by parser.add_argment()
# args = parser.parse_args() # Step 4. Analyze parameters
# print('-v=' + args.version)
# appliedVersion = args.version

print("1. Start the modification of VPDL Version in csproj file. \n")
#testVersion = "7.0.0.00000"
testVersion = "7.0.1.27813" ## It is the build version of VPDL example when JK tested in the old PC.
csproj_filename = "QAGPTPJT.csproj"
csproj_path = os.path.join(os.getcwd(), csproj_filename)
namespace = "" 
namespace_prefix = ""

print("2. Open csproj file. \n")
with open(csproj_path, 'r', encoding='UTF8') as f:
    tree_ = ET.parse(f)
    root_ = tree_.getroot()
    strNamespace = root_.tag ## root_.tag: {http://schemas.microsoft.com/developer/msbuild/2003}Project    
    strNamespace = strNamespace.strip("Project") ## {http://schemas.microsoft.com/developer/msbuild/2003}
    namespace_prefix = strNamespace ## >> 파싱할 때는 {namespace} prefix를 꼭 태그명앞에 붙여 주어야 한다.
    strNamespace = strNamespace.strip("{") 
    strNamespace = strNamespace.strip("}") ## strip 특정 문자열 제거 : https://engineer-mole.tistory.com/238   
namespace = strNamespace ## set namespace : http://egloos.zum.com/sweeper/v/3045388
ET.register_namespace('', namespace) # Must register namespace in elementItemGroupTree before parsing xml file. because of adding ns0 in save file.
with open(csproj_path, 'r', encoding='UTF8') as file:
    tree = ET.parse(file)
    root = tree.getroot()

print("3. Replace old version with new version.\n")
# Find ItemGroup to use reference tag in Packages
# elementItemGroup = tree.findall(".//{http://schemas.microsoft.com/developer/msbuild/2003}ItemGroup") ## namespace 적용전  코드
elementItemGroup = tree.findall(namespace_prefix+"ItemGroup") ## namespace 적용 후  코드
for index, referenceTag in enumerate(elementItemGroup[0]):
    existsearchKey = referenceTag.attrib.get('Include')
    #if (existsearchKey.find('ViDi')==False):     ## 변수.find('검색키워드') : 사용하면 매칭시에 반환값 0(zero), The return value in case of mismatching is '-1'
    if (existsearchKey.find('ViDi')>=False):     ## 변수.find('검색키워드') return value is position index number        
        ## newsearchKey = existsearchKey.replace("Version=7.0.1.27813", "Version=7.0.0.00000")
        ##newsearchKey = existsearchKey.replace("Version=7.0.1.27813", ("Version="+appliedVersion))
        newsearchKey = existsearchKey.replace(("Version="+testVersion), ("Version="+appliedVersion))
        elementItemGroup[0][index].set("Include", newsearchKey)
    for idx, hintpathTag in enumerate(referenceTag):
        existhintpath = hintpathTag.text ## >> ..\packages\AWSSDK.Core.3.3.103.34\lib\net45\AWSSDK.Core.dll        
        if(existhintpath.find('ViDi') > 0):
            ## hintpathTag.text = existhintpath.replace("7.0.1.27813", "7.0.0.00000")
            ##hintpathTag.text = existhintpath.replace("7.0.1.27813", appliedVersion) ##testVersion
            hintpathTag.text = existhintpath.replace(testVersion, appliedVersion) ##testVersion

print("4. Save new version in csproj file. \n")
## Applied new version and Save csproj file
with open(csproj_path, mode='wb') as file:
    tree.write(file, encoding='utf-8', xml_declaration=True)

print("5. Complete the applied new version. \n")

