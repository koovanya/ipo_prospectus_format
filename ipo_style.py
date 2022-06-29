#依据《中国证券监督管理委员会公告[2020]39号——公开发行证券的公司信息披露内容与格式准则第45号——科创板上市公司发行证券申请文件》
#原文件地址http://www.sse.com.cn/lawandrules/regulations/csrcannoun/c/5178163.pdf

from docx import Document
import docx.shared

#导入word
doc = Document('')

# 创建自定义段落样式(第一个参数为样式名, 第二个参数为样式类型, 1为段落样式, 2为字符样式, 3为表格样式)
mystyle = doc.styles.add_style('UserStyle1', 1)
# 设置字体尺寸
mystyle.font.size = docx.shared.Pt(40)
# 设置字体颜色
mystyle.font.color.rgb = docx.shared.RGBColor(0xff, 0xde, 0x00)
# 居中文本
mystyle.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
# 设置中文字体
mystyle.font.name = '微软雅黑'
mystyle._element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
