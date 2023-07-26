from pptx import Presentation

def get_slide_layouts(prs):
    layouts = []
    for slide in prs.slides:
        layout = slide.slide_layout.name
        layouts.append(layout)
    return layouts

if __name__ == "__main__":
    ppt_file_path =  rf"F:\pptx2md\大气工作总结计划汇报PPT模板.pptx"
    presentation = Presentation(ppt_file_path)

    slide_layouts = get_slide_layouts(presentation)
    for idx, layout_name in enumerate(slide_layouts):
        print(f"Slide {idx + 1} layout: {layout_name}")



