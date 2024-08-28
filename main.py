import pandas as pd
import os,glob
import re
stopWords = {}
posWords = {}
negWords = {}

def countSyllables(word):
    vowels = 'aeiouy'
    ans = 0
    for i in word:
        if i in vowels:
            ans+=1
    if word.endswith('es') or word.endswith('ed'):
        if ans > 1:
            ans-=1
    return ans
def fill():
    file_pattern = os.path.join('StopWords/', "*.txt")
    text_files = glob.glob(file_pattern)
    for file in text_files:
        with open(file,'r',encoding='utf-8') as file:
            content = file.readlines()
            content = [line.strip() for line in content]
            content = [line.split(" ")[0] for line in content]
            for word in content:
                word = word.lower()
                stopWords[word] = True
    
    file = 'MasterDictionary/negative-words.txt'
    with open(file,'r',encoding='utf-8') as file:
        content = file.readlines()
        content = [line.strip() for line in content]
        for word in content:
            negWords[word] = True
    
    
    file = 'MasterDictionary/positive-words.txt'
    with open(file,'r',encoding='utf-8') as file:
        content = file.readlines()
        content = [line.strip() for line in content]
        for word in content:
            posWords[word] = True

def clean(file_content):
    newFile=[]
    for line in file_content:
        line = line.strip()
        if line == "":
            continue
        else :
            words = line.split()
            for word in words:
                word = cleanWord(word)
            newFile.append(words)
    newFile.pop()
    return newFile

def cleanWord(word):
    cleaned_word = re.sub(r'^[^a-zA-Z]+|[^a-zA-Z]+$', '', word)
    return cleaned_word

def score(file_content):
    pos = 0
    neg = 0
    total = 0
    sentences = 0
    for line in file_content:
        sentences+=1
        for word in line:
            word = word.lower()
            if posWords.get(word) is not None:
                pos+=1
            if negWords.get(word) is not None:
                neg+=1
            if stopWords.get(word) is None:
                total+=1
    avg = total/sentences
    return pos,neg,total,avg

def moreScores(file_content):
    complexWords = 0 
    syllableCount = 0
    charcount = 0
    ppcount = 0
    pp = ['i','we','my','ours','us']
    for line in file_content:
        for word in line:
            if word.lower() in pp and word!='US':
                ppcount+=1 
            if stopWords.get(word) is None:
                charcount+=len(word)
            word = word.lower()
            
            syllable = countSyllables(word)
            syllableCount+=syllable
            if(syllable>2):
                complexWords+=1

    return complexWords,syllableCount,charcount,ppcount
def main(excel_file):
    
    df = pd.read_excel(excel_file)
    fill()
    
    
    for index, row in df.iterrows():
        file_name = row['URL_ID']
        file_name_with_extension = f"articles/{file_name}.txt"
        
        # Check if file exists
        if os.path.isfile(file_name_with_extension):
            with open(file_name_with_extension, 'r', encoding='utf-8') as file:
                file_content = file.readlines()
                file_content = clean(file_content)
                pos,neg,total,avg = score(file_content)
                complexWords,syllableCount,charcount,ppcount = moreScores(file_content)
                complexWordsPercent = complexWords/total
                polarity = (pos-neg)/((pos+neg)+0.000001)
                subjectivity = (pos+neg)/(total+0.000001)
                fog = 0.4*(avg+complexWordsPercent)

                df.at[index, 'POSITIVE SCORE'] =pos
                df.at[index, 'NEGATIVE SCORE'] =neg
                df.at[index, 'POLARITY SCORE'] =polarity
                df.at[index, 'SUBJECTIVITY SCORE'] = subjectivity
                df.at[index, 'PERCENTAGE OF COMPLEX WORDS'] = complexWordsPercent
                df.at[index, 'FOG INDEX'] = fog
                df.at[index, 'AVG SENTENCE LENGTH'] = avg
                df.at[index, 'AVG NUMBER OF WORDS PER SENTENCE'] = avg
                df.at[index, 'COMPLEX WORD COUNT'] = complexWords
                df.at[index, 'WORD COUNT'] = total
                df.at[index, 'SYLLABLE PER WORD'] = syllableCount/total
                df.at[index, 'PERSONAL PRONOUNS'] = ppcount
                df.at[index, 'AVG WORD LENGTH'] = charcount/total

                
        
        
        
        else:
            print(f"File not found: {file_name_with_extension}")
        
        
        df.to_excel('output.xlsx', index=False)

# Run the main function
if __name__ == "__main__":
    excel_file = 'Output.xlsx'  # Change this to your actual Excel file name
    main(excel_file)
