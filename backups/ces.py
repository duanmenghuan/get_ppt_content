# import nltk # 导入nltk库
# from nltk.corpus import words # 导入英文词典
# text = "This" # 定义一个英文文本
# tokens = nltk.word_tokenize(text) # 使用nltk库对文本进行分词
# for token in tokens: # 遍历每个单词
#     if token.lower() in words.words(): # 判断单词是否在英文词典中
#         print(token, "is a word.")
#     else:
#         print(token, "is not a word.")

# import re # 导入re模块
# pattern = r"^[a-zA-Z]+$" # 定义一个正则表达式，表示由一个或多个字母组成的字符串
# string = "Hello" # 定义一个字符串
# match = re.match(pattern, string) # 使用re模块对字符串进行匹配
# if match: # 判断是否匹配成功
#     print(string, "is a word.")
# else:
#     print(string, "is not a word.")
text = "zhangsan"
from langdetect import detect
def is_english_string(text):
    language = detect(text)
    if language == 'en':
        return True
    else:
        print(f"The string is in {language} language, not an English string.")
        return False




print(is_english_string(text))