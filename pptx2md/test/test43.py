# def match_patterns_with_disorder(patterns):
#     titles = []
#     bodies = []
#     excess_titles = []
#     excess_bodies = []
#
#     #将标题和正文提取到单独的列表中
#     for item in patterns:
#         if item.startswith("标题"):
#             titles.append(item)
#         elif item.startswith("正文"):
#             bodies.append(item)
#
#     # 检查标题是否多余
#     while len(titles) > len(bodies) + 1:
#         excess_titles.append(titles.pop())
#
#     while len(bodies) > len(titles):
#         excess_bodies.append(bodies.pop())
#
#     # 结合匹配的标题和正文，同时保留其原始顺序
#     result = []
#     while titles and bodies:
#         result.append(titles.pop(0))  # 添加匹配的标题
#         result.append(bodies.pop(0))
#
#     # 添加剩余的正文
#     while bodies:
#         result.append(bodies.pop(0))
#
#     # 添加多余的标题
#     result = excess_titles + result
#
#     # 添加任何剩余的多余物体
#     result.extend(excess_bodies)
#
#     return result
#
# # 测试示例
# input_patterns = ["标题1", "正文1", "标题2", "正文2", "正文3", "标题3", "标题4", "标题5",'正文','正文','正文']
# output_patterns = match_patterns_with_disorder(input_patterns)
# print(output_patterns)

def get_left_positions(slide):
    left_positions = []
    # 遍历每一张幻灯片
    for slide in slide.slides:
        # 遍历幻灯片中的每个形状
        for shape in slide.shapes:
            # 判断形状是否是文本框
            if shape.has_text_frame:
                left_positions.append(shape.left)

    return left_positions


if __name__ == "__main__":
    # 幻灯片文件路径
    pptx_file = "your_presentation.pptx"

    # 获取所有文本框的left属性
    left_positions_list = get_left_positions(pptx_file)

    # 打印结果
    print(left_positions_list)
