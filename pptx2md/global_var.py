'''
这段代码定义了一个名为 global_var 的类和一个 g 的对象。
以下是代码的解释：
定义了一个 global_var 类，它继承自 object。
global_var 类的 __init__ 方法为空，即没有任何初始化逻辑。
创建了一个 g 对象，该对象是 global_var 类的实例，用于存储全局变量。
在 g 对象中定义了一系列全局变量，包括：
img_path：图像输出目录的路径。
max_img_width：图像的最大宽度。
use_custom_title：是否使用预定义的标题列表。
out_path：输出文件的路径和文件名。
text_block_threshold：文本块转换的字符阈值。
disable_image：是否禁用图像提取。
disable_color：是否禁用颜色的 HTML 标签。
disable_escaping：是否禁用特殊字符的转义。
titles：标题字典，用于存储标题和级别信息。
file_prefix：文件的前缀名。
last_title：最后一个标题的信息。
max_custom_title：最大自定义标题级别。
page：指定的页面数。
该代码段的目的是定义和初始化一些全局变量，用于在程序的不同部分共享和访问这些变量的值。通过创建 global_var 类和 g 对象，可以在代码的其他部分通过 g 对象来访问和修改这些全局变量的值。
'''
class global_var(object):

    def __init__(self):
        pass

    # utilities
    def path_name_ext(self, path, name, ext):
        return path + '/' + name + '.' + ext


g = global_var()

# configs
# image output dir
g.img_path = 'img'
# maximum image width
g.max_img_width = None
# weather use predefined TOC in titles.txt
g.use_custom_title = False
# output path & filename
g.out_path = 'out.md'
# text frame thar contain more characters than this will be transferred
g.text_block_threshold = 3
# disable image extraction
g.disable_image = False
# prevent adding html tags with colors
g.disable_color = False
# prevent escaping of characters
g.disable_escaping = False

# global variables
g.titles = {}
g.file_prefix = '1'

g.last_title = {}
g.max_custom_title = 1

g.page = None
