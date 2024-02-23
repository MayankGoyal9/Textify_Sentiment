import pandas as pd
import os
import requests
from bs4 import BeautifulSoup
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from nltk.tokenize import sent_tokenize
import textstat
from xlsxwriter import Workbook
from openpyxl import Workbook
import openpyxl

def extract(excelfile):
    try:
        df= pd.read_excel(excelfile)
        urls = df.values.tolist()
        return urls
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

def create(urls):
    try:
        for row in urls:
            url_id = str(row[0])
            url = row[1]
            filename = f"{url_id}.txt"
            with open(filename, 'w') as file:
                file.write('')
            print(f"Text file created: {filename}")
            browse(url, filename)
    except Exception as e:
        print(f"Error creating text files: {e}")

def browse(url, filename):
    try:
        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            text_data = soup.get_text()
            with open(filename, 'r+') as file:
                file_content = file.read()
                text_data = text_data.replace(url, '')
                file.seek(0, os.SEEK_END)
                file.write('\n\n')
                file.write(text_data)
            print(f"Text data updated in file: {filename}")
        else:
            print(f"Failed to collect data from URL: {url}")
    except Exception as e:
        print(f"Error browsing URL and updating text data: {e}")

if __name__ == "__main__":

    excelfile = "C:\\Users\\mayan\\Desktop\\zpython\\Input.xlsx"
    if os.path.exists(excelfile):
        urls= extract(excelfile)
        if urls:
            create(urls)
        else:
            print("No data found in the Excel file.")
    else:
        print(f"Excel file '{excelfile}' not found.")

    combined_text_file = "C:\\Users\\mayan\\PycharmProjects\\pythonProject_black\\combined_text_files.txt"
    try:
        with open(combined_text_file, 'r') as file:
            lines_set = {line.strip() for line in file}
        print("Set created successfully.")
    except FileNotFoundError:
        print(f"Error: File '{combined_text_file}' not found.")
    except Exception as e:
        print(f"Error reading file: {e}")

    positive_words_file = "C:\\Users\\mayan\\PycharmProjects\\pythonProject_black\\positive-words.txt"
    try:
        with open(positive_words_file, 'r') as file:
            positive_words_set = {line.strip() for line in file}
        print("Positive words set created successfully.")
    except FileNotFoundError:
        print(f"Error: File '{positive_words_file}' not found.")
    except Exception as e:
        print(f"Error reading positive words file: {e}")

    negative_words_file = "C:\\Users\\mayan\\PycharmProjects\\pythonProject_black\\negative-words.txt"
    try:
        with open(negative_words_file, 'r') as file:
            negative_words_set = {line.strip() for line in file}
        print("Negative words set created successfully.")
    except FileNotFoundError:
        print(f"Error: File '{negative_words_file}' not found.")
    except Exception as e:
        print(f"Error reading negative words file: {e}")

    files = os.listdir()
    text_files = [file for file in files if file.endswith('.txt')]

    for text_file in text_files:
        if text_file.find("blacka")==-1:
            continue
        print(f"Reading contents of {text_file}:")
        try:
            with open(text_file, 'r') as file:
                contents = file.read()
        except Exception as e:
            print(f"Error reading {text_file}: {e}")

        sentences= sent_tokenize(contents)
        word_tokens = word_tokenize(contents)
        personal_pronouns = ['i', 'you', 'he', 'she', 'it', 'we', 'they', 'me', 'him', 'her', 'us', 'them', 'myself',
                             'yourself', 'himself', 'herself', 'itself', 'ourselves', 'yourselves', 'themselves']
        pp_counter = 0
        for w in word_tokens:
            if w.lower() in personal_pronouns:
                pp_counter+=1

        filtered_sentence = [w for w in word_tokens if not w.lower() in lines_set]
        complex_count=0
        arr=[0]*13
        syl_sum=0
        ch_counter=0
        for w in filtered_sentence:
            if w in positive_words_set:
                arr[0]+=1
            if w in negative_words_set:
                arr[1]+=1
            syl_count=textstat.syllable_count(w)
            syl_sum+=syl_count
            if syl_count>2:
                complex_count+=1
            ch_counter+=len(w)
        arr[2]=(arr[0]-arr[1])/((arr[0]+arr[1])+0.000001)

        arr[3]=(arr[0]+arr[1])/((len(filtered_sentence))+ 0.000001)

        if(len(sentences)==0):
            arr[4]=len(filtered_sentence)
        else:
            arr[4]=len(filtered_sentence)/len(sentences)

        if(len(filtered_sentence)==0):
            arr[5] = complex_count
        else:
            arr[5]=complex_count/len(filtered_sentence)

        arr[6]=0.4 * (arr[4]+arr[5])

        if (len(sentences) == 0):
            arr[7]=len(filtered_sentence)
        else:
            arr[7]=len(filtered_sentence)/len(sentences)

        arr[8] =complex_count

        arr[9] =len(filtered_sentence)

        arr[10]=syl_sum

        arr[11] =pp_counter

        if (len(filtered_sentence) == 0):
            arr[12] =ch_counter
        else:
            arr[12] =ch_counter/len(filtered_sentence)

        name, extension= os.path.splitext(text_file)
        start_row = int(name[-4:])+1
        wb=openpyxl.load_workbook("C:\\Users\\mayan\\PycharmProjects\\pythonProject_black\\Output Data Structure.xlsx")
        sheet=wb.active
        start_column = 2
        for col in range(3,16):
            sheet.cell(row=start_row, column=col, value=arr[col-3])
        wb.save("C:\\Users\\mayan\\PycharmProjects\\pythonProject_black\\Output Data Structure.xlsx")




