import os
import docx
import json

'''
{
    "id": "",
    "quote": "",
    "yearText": "",
    "monthText": "",
    "dayText": "",
    "timeText": "",
    "ampm": ""
  }
'''

 
def read_docx(path):
    try:
        doc = docx.Document(path)  # Creating word reader object.
        data = ""
        fullText = []
        # fullText.append("[")
        count = 0
        qLen = 1
        id = quote = yearText = monthText = dayText = timeText = ampm = head = ''
        nextLine = True
        sep = ','

        for para in doc.paragraphs:
            lineText = para.text
            if(nextLine == False):
                id = ''
            if('-' in lineText and ';' in lineText and '[' in lineText and ']' in lineText and ':' in lineText):
                nextLine = True
                head = lineText.split(' ')
                dateText = head[0].strip().replace(';', '')
                dateList = dateText.split('-')
                yearText = dateList[0]
                monthText = dateList[1]
                dayText = dateList[2]

                timeText = head[1].strip()
                ampm = head[2]

                id = head[3].strip().replace('[', '').replace(']', '').replace('dq', '')
                # print(id)
            elif(nextLine == True and lineText and id):
                quote = lineText.strip()
                # mask ' and "
                quote = quote.replace('"', '\\"')
                # print(body)
                nextLine = False
                qLen +=1

                output = '''
                    {
                        "id": "%s",
                        "quote": "%s",
                        "yearText": "%s",
                        "monthText": "%s",
                        "dayText": "%s",
                        "timeText": "%s",
                        "ampm": "%s"
                    }%s
                ''' % (id, quote, yearText, monthText, dayText, timeText, ampm, sep)
                fullText.append(output)
                # data = '\n'.join(fullText)
                data = ''.join(fullText)
            count +=1
        # print(data)
        print("The json file '%s' was created successfully with %d Quotes" % (path, qLen))
        return data, qLen
 
 
    except IOError:
        print('There was an error opening the file!')
        return

def read_file(path):
    with open(path, 'r') as f:
        return f.read()

def listDir(path, jsonName):
    os.chdir(path)
    fullText = []
    data = ''
    tLen = 0
    fullText.append("[")
    for file in os.listdir():
        if file.endswith('.docx'):
            # file_path = f"{path}\{file}"
            dataDocx, qLen = read_docx(file)
            fullText.append(dataDocx)
            tLen += qLen
    
    fullText.append("]")
    data = ''.join(fullText)


    with open(jsonName, 'w', encoding="utf-8") as f:
        f.write(data)
        print("The json file was created successfully with %d items" %(tLen))

if __name__ == '__main__':
    msg = '''
        ####
        # This Program reads from files in a given directory (default: files)
        # and creates a json file (default output.json) in same directory.
        # You may supply a relative path or full path
        # --- Created by Chukwuemeka Ndefo ---
        # --- 29th June, 2022 ---
        # --- Visit https://github.com/metaversedataman
        ####
    '''
    fileName = 'output'
    dirName = 'files'

    print(msg)
    print('=== Skip to allow the default values ===')
    getDir = input('Enter directory name of path:   ')
    if(getDir):
        dirName = getDir
    
    jsonName = input('Enter preferred Json file name:  ')
    if(jsonName):
        fileName = jsonName
    
    listDir(dirName, "%s.json" %(fileName))


