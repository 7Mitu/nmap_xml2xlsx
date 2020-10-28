# nmap_xml2xlsx
将nmap -oX导出的xml文件解析为Excel表格

依赖：
- Python2.7
- xlsxwriter
``` cmd
pip install xlsxwriter
```

运行：
脚本会自动读取当前目录下的所有xml文件，并输出同名的xlsx文件。
``` shell
python ParsingXMLtoXLSX.py
```
