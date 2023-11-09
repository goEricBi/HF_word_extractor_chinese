'''
@Aurthor: Eric
@Date: 11/8/2023 
@Purpose: Extractor Keyworks
@Assumption: Other 字段 will not be under considerations, so I use list to store the data for simplicity 

'''
import jieba
import openpyxl 
import re
import json
import os

def preprocess_sentence(stop_words,sentence):
    """
    Per paper, this preprocess drops unused samples 
    """

    # if we need emojis, comment below "pattern" uncomment the second "filtered_sentence"
    pattern = re.compile(r'[\u4e00-\u9fffA-Za-z]+')
    words = jieba.cut(sentence)

    # Filter out stop words and match only Chinese characters and English alphabets
    filtered_sentence = ''.join(word for word in words if word not in stop_words and (pattern.match(word) or word in '，'))
    #filtered_sentence = ''.join(word for word in words if word not in stop_words)
    return filtered_sentence

def extract_and_preprocess_column(file_path, column_index=8, stop_words=None):
    """
    Load the workbook and select the active sheet
    """
    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    # Extract and preprocess the specified column
    processed_column_data = []
    for row in sheet.iter_rows(min_col=column_index, max_col=column_index, values_only=True):
        cell_value = row[0]  # Since we're only looking at one column, we take the first value
        if cell_value is not None:  # This will skip empty cells
            # Preprocess the sentence
            processed_sentence = preprocess_sentence(stop_words, cell_value)
            processed_column_data.append(processed_sentence)
    
    return processed_column_data

# Bad syntax for simplicity
def main():

    # Input path
    # Check the root, I hide my path here 
    file1 = r'tongcheng_travel_review1.xlsx'
    file2 = r'tongcheng_travel_review2.xlsx'
    stop_words = set(['呢', '啦', '吧', '嘛', '啊', '哦', '的', '了', '在','太'])

    processed_column_data = extract_and_preprocess_column(file1,stop_words=stop_words) + extract_and_preprocess_column(file2,stop_words=stop_words)

    # Cut words 
    # Tune: maybe jieba.lcut_for_search will be better for searching purpose
    cut_words = [word for sentence in processed_column_data for word in jieba.lcut(sentence)]

    data = {}

    for char in cut_words:
        if len(char)<2:
            continue 
        if char in data: 
            data[char] += 1
        else:
            data[char]=1 

    data=sorted(data.items(),key=lambda x:x[1], reverse=True)

    # Optimization: 
    # Used lcut once in preprocess, which might be bad in terms of results...
    print(data)

    script_dir = os.path.dirname(__file__)
    filename = os.path.join(script_dir, 'result.json')

    # Writing JSON data
    with open(filename, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


if __name__ == "__main__":
    main()