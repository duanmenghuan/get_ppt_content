import json

def move_last_title_to_first(json_data):
    try:
        data = json.loads(json_data)

        if not data:
            raise ValueError("JSON数据是空")

        for num in range(len(data)):
            content = data[num]['content']
            last_title_index = None

            for i, item in enumerate(content):
                if item['type'] == '标题':
                    last_title_index = i


            if last_title_index is not None:
                content.insert(1, content.pop(last_title_index))

        return json.dumps(data, ensure_ascii=False, indent=2)

    except json.JSONDecodeError as e:
        raise ValueError(f"无效的JSON格式: {e}")

if __name__ == "__main__":
    # 读取 JSON 文件内容
    with open('test002.json', 'r', encoding='utf-8') as file:
        json_data = file.read()

    # 执行移动操作
    modified_json = move_last_title_to_first(json_data)

    # 保存修改后的数据回 JSON 文件
    with open('test002.json', 'w', encoding='utf-8') as file:
        file.write(modified_json)
