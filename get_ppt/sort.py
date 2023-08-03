import json

def rearrange_content_type(content):
    cover_title = None
    directory_titles = []  # 使用列表存储所有目录标题
    top_title = None
    titles = []
    texts = []
    top_title_contents = None
    chapter_head = []

    for i ,item in enumerate(content):
        # if i == 1 or i == 2 :
        #     if item["type"] == " ":
        #        item["type"] =  '目录章节标题'
        # else:

            if item["type"] == "封面标题":
                cover_title = item
            elif item["type"] == "目录标题":
                directory_titles.append(item)  # 将所有目录标题添加到列表中
            elif item["type"] == "章节标题":
                top_title = item
            elif item["type"] == "标题":
                # 在处理 "标题" 时，先将多余的标题添加到新内容列表的最前面
                 titles.append(item)
            elif item["type"] == "正文":
                if item["hint_ext"] == "行业PPT模板http://www.1ppt.com/hangye/":
                    continue
                else:
                    texts.append(item)
            elif item["type"] == "副标题":
                texts.append(item)
            elif item["type"] == "目录章节标题":
                if item['hint_ext'].isdigit():
                    continue
                else:
                    chapter_head.append(item)
                    if len(chapter_head) >1:
                        continue
                    elif len(chapter_head) == 1:
                        texts.append(chapter_head[0])
            elif item['type']=="":
                continue




    # print(titles,len(titles))
    # print(texts,len(texts))
    new_content = []
    if cover_title:
        new_content.append(cover_title)
    if top_title:
        new_content.append(top_title)

    # 将所有目录标题添加到 new_content 列表中
    new_content.extend(directory_titles)

    # 在目录标题之间交替添加标题和正文
    max_len = max(len(titles), len(texts))
    title_text_length_diff  = len(titles) - len(texts)
    #print(title_text_length_diff)
    for i in range(max_len):
        if i < len(titles):
            new_content.append(titles[i])
        if i < len(texts):
            new_content.append(texts[i])

    if title_text_length_diff == 1:
        num = len(titles) + len(texts)
        if num < len(new_content):
            new_content.insert(1, new_content.pop(num))
        else:
            pass

    return new_content


with open(rf"F:\pptx2md\json\哆啦A梦.json", "r", encoding='utf-8') as file:
    data = json.load(file)

for slide in data:
    slide["content"] = rearrange_content_type(slide["content"])

# 将处理后的数据写入新的 JSON 文件
with open(rf"F:\pptx2md\json\哆啦A梦.json", "w", encoding="utf-8") as file:
    json.dump(data, file, indent=2, ensure_ascii=False)

print("JSON文件已成功写入。")
