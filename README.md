# 中文简介
这个程序能从学术期刊类 PDF 文件中提取文本并保存为 DOCX 文件。程序能够识别并合并跨页的正文段落，同时将脚注单独提取，改为尾注，以便于后续的编辑和使用，譬如全文翻译。程序只针对文本类pdf，未有提取图片功能。程序也限定于以数字开始的脚注。另外，脚注的识别和正文段落的合并准确率仍非100%。碍于本人有限的能力，无法再更新本程序，期待能抛砖引玉，有大神拿去完善。  
调用方法在英文简介后面。

# English Description
This program can extract text from academic journal-style PDF files and save it as a DOCX file. It is capable of identifying and merging body paragraphs that span multiple pages, while separately extracting footnotes and converting them into endnotes for subsequent editing and use, such as full-text translation. The program is designed specifically for text-based PDFs and does not include image extraction functionality. It is also limited to footnotes that begin with numbers. Additionally, the accuracy of footnote recognition and paragraph merging is not yet 100%. Due to my limited abilities, I am unable to update this program further, but I hope this can serve as a starting point for others, and I look forward to seeing experts improve it.

# 用法 Usage
extract_pdf2docx.exe   [pdf_path] [-h] [--skip-header] [--skip-footer] [--mixed-footnotes] 

| 参数							 | 用途	|
|--------------------|-----------|
| pdf_path	         |	Path to the input PDF file. The output file defaults to a DOCX file with the same name as the PDF, located in the same directory. 输入文件的路径。输出文件默认为与它同目录下同名的docx文件。  |
|  -h, --help        |   show this help message and exit. 提示参数信息。  |
|  --skip-header     |   Skip the header line on each page. 跳过每页的页眉文字。|  
|  --skip-footer     | Skip the footer line on each page. 跳过每页的页脚文字。|  
|  --mixed-footnotes | Indicate if there are letter footnotes following number footnotes (The program does not yet handle the case where letter footnotes appear before number footnotes). 若文档有混合脚注，譬如数字脚注后面跟着字母脚注，则需开启这参数（程序未能处理字母脚注在数字脚注之前的情况）。|
