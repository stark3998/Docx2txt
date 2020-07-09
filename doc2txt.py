import importlib

def install(package):
    try:
        if "win" in sys.platform:
            subprocess.check_call([sys.executable, "-m", "pip", "--no-cache-dir", "install", package])
        else:
            subprocess.check_call(["sudo", "-H", "pip3", "install", "--upgrade", "--no-cache-dir", package])
            # subprocess.check_call(["sudo", "pip3", "--no-cache-dir", "install", package])
    except Exception as e:
        print("Exception occured in installation : ",e)

listimport = ['sys','subprocess','docx2txt','argparse','glob']

for x in listimport:
    try:
        globals()[x]=importlib.import_module(x)
        print("Imported Module ",x)
    except Exception as e:
        print(e)
        install(x)
        globals()[x]=importlib.import_module(x)

parser = argparse.ArgumentParser()

parser.add_argument("-convert", "--convert", help="Path of folders to convert")
args = parser.parse_args()

folders=[]
try:
    if args.convert:
        folders=args.convert.split(',')
    else:    
        print("No Folders Provided")
except:
    print("Exception Occured")

print("Converting Folders : ",folders)

for i in folders:
    files=glob.glob("{}/*.docx".format(i))
    print("{} Docx Files found in {} : {}".format(len(files),i,files))
    for i in files:
        file_name=i.replace("\\","/").split('/')[-1].split('.')[0]
        if "win" in sys.platform:
            file_path="\\".join(i.replace("/","\\").split("\\")[:-1])+"\\"
        else:
            file_path="/".join(i.replace("\\","/").split("/")[:-1])+"/"
        doc_text=docx2txt.process(i)
        with open(file_path+"\\"+file_name+".txt","w",encoding='utf-8') as f:
            f.write(doc_text)
        print("Converted {}.docx to {}.txt".format(file_name,file_name))