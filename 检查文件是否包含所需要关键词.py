#使用场景："
#我现在需要检查同文件夹中TXT文件是否含有‘事業利益’,'コア営業利益','調整後営業利益','コア利益','Core利益','EBITDA','キャシュー利益','キャッシュ利益','EBIT','実業ベース','キャシュアーニングス','フリーキャッシュフロー','基礎営業キャッシュ','キャシュー利益','調整後EBITDA','EBITA','EVA','FIV','Non-GAAP'。 同样也是保存在EXCEL文件中。比如EXCEL的A列是文件名，B列以后是检查内容在文件中出现次数排列整理。
#代码如下：
import os
import pandas as pd

def check_keywords_in_file(file_path, keywords):
    with open(file_path, 'r', encoding='utf-8') as file:
        contents = file.read()
    keyword_count = {}
    for keyword in keywords:
        count = contents.count(keyword)
        keyword_count[keyword] = count
    return keyword_count

folder_path = r'E:\kenkyu\data\500TXT'
keywords = ['事業利益', 'コア営業利益', '調整後営業利益', 'コア利益', 'Core利益', 'EBITDA', 
            'キャシュー利益', 'キャッシュ利益', 'EBIT', '実業ベース', 'キャシュアーニングス', 
            'フリーキャッシュフロー', '基礎営業キャッシュ', 'キャシュー利益', '調整後EBITDA', 
            'EBITA', 'EVA', 'FIV', 'Non-GAAP']

results = []

for file_name in os.listdir(folder_path):
    if file_name.endswith('.txt'):
        file_path = os.path.join(folder_path, file_name)
        keyword_count = check_keywords_in_file(file_path, keywords)
        result_row = {'File Name': file_name}
        result_row.update(keyword_count)
        results.append(result_row)

df = pd.DataFrame(results)
df.to_excel('results.xlsx', index=False)
