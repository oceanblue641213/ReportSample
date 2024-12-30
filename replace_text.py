def replace_text_in_word(doc, replacements):
    # 遍歷段落
    for paragraph in doc.paragraphs:
        for old, new in replacements.items():
            if old in paragraph.text:
                # 获取段落的样式
                paragraph_style = paragraph.style
                
                # 在段落的所有run中查找并替换
                for run in paragraph.runs:
                    if old in run.text:
                        # 替换文字并保留样式
                        run.text = run.text.replace(old, new)