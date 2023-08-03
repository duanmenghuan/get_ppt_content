import json

# 从文件中读取JSON数据
file_path = "_test002.json"  # 替换为实际的JSON文件路径
with open(file_path, "r", encoding="utf-8") as file:
    data = json.load(file)



with open('template_text.txt', "w", encoding="utf-8") as output_file:
    for item in data:
        slide_index = item["slide_index"]
        for content_item in item["content"]:
            hint_text = content_item["hint_ext"]
            id =  content_item["id"]
            if len(hint_text) >0:
                output_text = f"Slide_{slide_index}_{id}_{hint_text}\n"
                output_file.write(output_text)



print('完成')



