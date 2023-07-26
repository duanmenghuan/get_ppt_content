from rapidfuzz import fuzz #fuzzywuzzy库是Python中的模糊匹配库，它依据 Levenshtein Distance 算法 计算两个序列之间的差异。
from pptx2md.global_var import g
import re
import os
import urllib.parse

'''
这段代码定义了一个名为 outputter 的类，它是一个输出器的基类。
以下是代码的解释：
定义了一个 outputter 类，它继承自 object。
outputter 类的 __init__ 方法接受一个 file_path 参数，表示输出文件的路径。
在 __init__ 方法中，通过调用 os.makedirs 方法创建输出文件所在目录，并使用 os.path.dirname 和 os.path.abspath 方法获取输出文件路径的绝对路径。
使用 open 函数以写入模式（'w'）和 UTF-8 编码（'utf8'）打开输出文件，将文件对象赋值给 self.ofile。
定义了一系列方法，用于向输出文件写入不同类型的内容。这些方法包括：
put_title：写入标题。
put_list：写入列表。
put_para：写入段落。
put_image：写入图像。
put_table：写入表格。
get_accent：返回带有重音的文本。
get_strong：返回加粗的文本。
get_colored：返回带有颜色的文本。
get_hyperlink：返回超链接的文本。
get_escaped：返回转义的文本。
write：将指定的文本写入输出文件。
flush：刷新输出缓冲区。
close：关闭输出文件。
该代码段定义了一个基础的输出器类，提供了一些方法用于向输出文件写入不同类型的内容。这个类的实例可以被子类继承并扩展，以实现特定格式的输出逻辑。
'''
class outputter(object):

    def __init__(self, file_path):
        os.makedirs(os.path.dirname(os.path.abspath(file_path)), exist_ok=True)
        self.ofile = open(file_path, 'w', encoding='utf8')

    def put_title(self, text, level):
        pass

    def put_list(self, text, level):
        pass

    def put_para(self, text):
        pass

    def put_image(self, path, max_width):
        pass

    def put_table(self, table):
        pass

    def get_accent(self, text):
        pass

    def get_strong(self, text):
        pass

    def get_colored(self, text, rgb):
        pass

    def get_hyperlink(self, text, url):
        pass

    def get_escaped(self, text):
        pass

    def write(self, text):
        self.ofile.write(text)

    def flush(self):
        self.ofile.flush()

    def close(self):
        self.ofile.close()

'''
这段代码定义了一个名为 md_outputter 的子类，它继承自 outputter 类，并用于将输出写入 Markdown 格式的文件。
以下是代码的解释：
定义了一个 md_outputter 类，它继承自 outputter 类。
在 __init__ 方法中，首先调用父类的 __init__ 方法，以初始化继承的属性和方法。
使用正则表达式编译了两个用于转义的模式：esc_re1 和 esc_re2。
重写了父类的 put_title 方法，用于将标题写入 Markdown 文件。
去除标题两端的空白字符。
使用 fuzz.ratio 方法比较当前标题和上一个相同级别的标题的相似度，如果相似度低于阈值（96），则将标题写入文件，并更新 g.last_title 的对应级别的标题。
重写了父类的 put_list 方法，用于将列表写入 Markdown 文件。
在列表项前添加适当数量的空格和星号。
去除列表项两端的空白字符。
重写了父类的 put_para 方法，用于将段落写入 Markdown 文件。
直接将段落文本写入文件。
重写了父类的 put_image 方法，用于将图像写入 Markdown 文件。
如果未指定最大宽度，则使用 URL 编码后的路径作为图片的 Markdown 格式。
如果指定了最大宽度，则使用 <img> 标签将图像写入文件，同时设置最大宽度样式。
重写了父类的 put_table 方法，用于将表格写入 Markdown 文件。
定义了一个辅助函数 gen_table_row，用于生成表格行的 Markdown 格式。
写入表格的首行和分隔行。
写入表格的内容行。
重写了父类的一系列文本修饰方法，包括 get_accent、get_strong、get_colored 和 get_hyperlink，用于返回带有相应修饰的文本。
定义了一个辅助方法 esc_repl，用于替换转义字符串的匹配项。
重写了父类的 get_escaped 方法，用于对文本进行转义处理。
使用正则表达式替换将特殊字符进行转义。
该代码段定义了一个 Markdown 输出器类，实现了具体的输出逻辑，将内容以 Markdown 格式写入到文件中。子类重写了父类的方法，根据 Markdown 的语法规则生成相应的输出内容。
'''
class md_outputter(outputter):
    # write outputs to markdown
    def __init__(self, file_path):
        super().__init__(file_path)
        self.esc_re1 = re.compile(r'([\\\*`!_\{\}\[\]\(\)#\+-\.])')
        self.esc_re2 = re.compile(r'(<[^>]+>)')

    def put_title(self, text, level):
        text = text.strip()
        if not fuzz.ratio(text, g.last_title.get(level, ''), score_cutoff=96):
            self.ofile.write('#' * level + ' ' + text + '\n\n')
            g.last_title[level] = text

    def put_list(self, text, level):
        self.ofile.write('  ' * level + '* ' + text.strip() + '\n')

    def put_para(self, text):
        self.ofile.write(text + '\n\n')

    def put_image(self, path, max_width=None):
        if max_width is None:
            self.ofile.write(f'![]({urllib.parse.quote(path)})\n\n')
        else:
            self.ofile.write(f'<img src="{path}" style="max-width:{max_width}px;" />\n\n')

    def put_table(self, table):
        gen_table_row = lambda row: '| ' + ' | '.join([c.replace('\n', '<br />') for c in row]) + ' |'
        self.ofile.write(gen_table_row(table[0]) + '\n')
        self.ofile.write(gen_table_row([':-:' for _ in table[0]]) + '\n')
        self.ofile.write('\n'.join([gen_table_row(row) for row in table[1:]]) + '\n\n')

    def get_accent(self, text):
        return ' _' + text + '_ '

    def get_strong(self, text):
        return ' __' + text + '__ '

    def get_colored(self, text, rgb):
        return ' <span style="color:#%s">%s</span> ' % (str(rgb), text)

    def get_hyperlink(self, text, url):
        return '[' + text + '](' + url + ')'

    def esc_repl(self, match):
        return '\\' + match.group(0)

    def get_escaped(self, text):
        text = re.sub(self.esc_re1, self.esc_repl, text)
        text = re.sub(self.esc_re2, self.esc_repl, text)
        return text

'''
这段代码定义了一个名为 `wiki_outputter` 的子类，它继承自 `outputter` 类，并用于将输出写入为维基文本（wikitext）格式的文件。
以下是代码的解释：
1. 定义了一个 `wiki_outputter` 类，它继承自 `outputter` 类。
2. 在 `__init__` 方法中，首先调用父类的 `__init__` 方法，以初始化继承的属性和方法。
3. 使用正则表达式编译了一个用于转义的模式 `esc_re`。
4. 重写了父类的 `put_title` 方法，用于将标题写入维基文本文件。
   - 去除标题两端的空白字符。
   - 使用 `fuzz.ratio` 方法比较当前标题和上一个相同级别的标题的相似度，如果相似度低于阈值（96），则将标题写入文件，并更新 `g.last_title` 的对应级别的标题。
   - 使用适当数量的感叹号作为标题的级别标识。
5. 重写了父类的 `put_list` 方法，用于将列表写入维基文本文件。
   - 使用适当数量的星号作为列表项的级别标识。
   - 去除列表项两端的空白字符。
6. 重写了父类的 `put_para` 方法，用于将段落写入维基文本文件。
   - 直接将段落文本写入文件。
7. 重写了父类的 `put_image` 方法，用于将图像写入维基文本文件。
   - 如果未指定最大宽度，则使用 `<img>` 标签将图像写入文件。
   - 如果指定了最大宽度，则使用 `<img>` 标签，并设置图像的宽度。
8. 重写了父类的一系列文本修饰方法，包括 `get_accent`、`get_strong`、`get_colored` 和 `get_hyperlink`，用于返回带有相应修饰的文本。
9. 定义了一个辅助方法 `esc_repl`，用于替换转义字符串的匹配项。
10. 重写了父类的 `get_escaped` 方法，用于对文本进行转义处理。
   - 使用正则表达式替换将尖括号内的文本进行转义。
该代码段定义了一个维基文本输出器类，实现了具体的输出逻辑，将内容以维基文本格式写入到文件中。子类重写了父类的方法，根据维基文本的语法规则生成相应的输出内容。
'''
class wiki_outputter(outputter):
    # write outputs to wikitext
    def __init__(self, file_path):
        super().__init__(file_path)
        self.esc_re = re.compile(r'<([^>]+)>')

    def put_title(self, text, level):
        text = text.strip()
        if not fuzz.ratio(text, g.last_title.get(level, ''), score_cutoff=96):
            self.ofile.write('!' * level + ' ' + text + '\n\n')
            g.last_title[level] = text

    def put_list(self, text, level):
        self.ofile.write('*' * (level + 1) + ' ' + text.strip() + '\n')

    def put_para(self, text):
        self.ofile.write(text + '\n\n')

    def put_image(self, path, max_width):
        if max_width is None:
            self.ofile.write(f'<img src="{path}" />\n\n')
        else:
            self.ofile.write(f'<img src="{path}" width={max_width}px />\n\n')

    def get_accent(self, text):
        return ' __' + text + '__ '

    def get_strong(self, text):
        return ' \'\'' + text + '\'\' '

    def get_colored(self, text, rgb):
        return ' @@color:#%s; %s @@ ' % (str(rgb), text)

    def get_hyperlink(self, text, url):
        return '[[' + text + '|' + url + ']]'

    def esc_repl(self, match):
        return "''''" + match.group(0)

    def get_escaped(self, text):
        text = re.sub(self.esc_re, self.esc_repl, text)
        return text

'''
这段代码定义了一个名为 `madoko_outputter` 的子类，它继承自 `outputter` 类，并用于将输出写入为 Madoko Markdown 格式的文件。
以下是代码的解释：
1. 定义了一个 `madoko_outputter` 类，它继承自 `outputter` 类。
2. 在 `__init__` 方法中，首先调用父类的 `__init__` 方法，以初始化继承的属性和方法。
3. 使用正则表达式编译了两个用于转义的模式：`esc_re1` 和 `esc_re2`。
4. 重写了父类的 `put_title` 方法，用于将标题写入 Madoko Markdown 文件。
   - 去除标题两端的空白字符。
   - 使用 `fuzz.ratio` 方法比较当前标题和上一个相同级别的标题的相似度，如果相似度低于阈值（96），则将标题写入文件，并更新 `g.last_title` 的对应级别的标题。
   - 使用适当数量的井号作为标题的级别标识。
5. 重写了父类的 `put_list` 方法，用于将列表写入 Madoko Markdown 文件。
   - 使用适当数量的空格作为列表项的级别标识。
   - 去除列表项两端的空白字符。
6. 重写了父类的 `put_para` 方法，用于将段落写入 Madoko Markdown 文件。
   - 直接将段落文本写入文件。
7. 重写了父类的 `put_image` 方法，用于将图像写入 Madoko Markdown 文件。
   - 如果未指定最大宽度，则使用 `<img>` 标签将图像写入文件。
   - 如果指定了最大宽度，并且小于 500，则使用 `<img>` 标签，并设置图像的宽度。
   - 如果指定的最大宽度大于等于 500，则使用 Madoko Markdown 的图像标记，包含标题和图像路径，并设置图像的宽度。
8. 重写了父类的一系列文本修饰方法，包括 `get_accent`、`get_strong`、`get_colored` 和 `get_hyperlink`，用于返回带有相应修饰的文本。
9. 定义了一个辅助方法 `esc_repl`，用于替换转义字符串的匹配项。
10. 重写了父类的 `get_escaped` 方法，用于对文本进行转义处理。
   - 使用正则表达式替换将特殊字符进行转义。
该代码段定义了一个 Madoko Markdown 输出器类，实现了具体的输出逻辑，将内容以 Madoko Markdown 格式写入到文件中。子类重写了父类的方法，根据 Madoko Markdown 的语法规则生成相应的输出内容。
'''
class madoko_outputter(outputter):
    # write outputs to madoko markdown
    def __init__(self, file_path):
        super().__init__(file_path)
        self.ofile.write('[TOC]\n\n')
        self.esc_re1 = re.compile(r'([\\\*`!_\{\}\[\]\(\)#\+-\.])')
        self.esc_re2 = re.compile(r'(<[^>]+>)')

    def put_title(self, text, level):
        text = text.strip()
        if not fuzz.ratio(text, g.last_title.get(level, ''), score_cutoff=96):
            self.ofile.write('#' * level + ' ' + text + '\n\n')
            g.last_title[level] = text

    def put_list(self, text, level):
        self.ofile.write('  ' * level + '* ' + text.strip() + '\n')

    def put_para(self, text):
        self.ofile.write(text + '\n\n')

    def put_image(self, path, max_width):
        if max_width is None:
            self.ofile.write(f'<img src="{path}" />\n\n')
        elif max_width < 500:
            self.ofile.write(f'<img src="{path}" width={max_width}px />\n\n')
        else:
            self.ofile.write('~ Figure {caption: image caption}\n')
            self.ofile.write('![](%s){width:%spx;}\n' % (path, max_width))
            self.ofile.write('~\n\n')

    def get_accent(self, text):
        return ' _' + text + '_ '

    def get_strong(self, text):
        return ' __' + text + '__ '

    def get_colored(self, text, rgb):
        return ' <span style="color:#%s">%s</span> ' % (str(rgb), text)

    def get_hyperlink(self, text, url):
        return '[' + text + '](' + url + ')'

    def esc_repl(self, match):
        return '\\' + match.group(0)

    def get_escaped(self, text):
        text = re.sub(self.esc_re1, self.esc_repl, text)
        text = re.sub(self.esc_re2, self.esc_repl, text)
        return text
