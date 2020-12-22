## 帆软自动word处理程序使用说明

1. 使用时与该程序同目录下会产生名为 `tempory.docx` 和`image.png`的临时文件，无需做任何处理
2. 文件分为：
   - finereport_docx_process.py - 用于进行文档的重组
   - auto_docx_process - 用于自动的清洗，可接受命令行参数
   - API_docx_process - 提供API接口

### 自动清洗程序 auto_docx_process

自动处理：

待处理文件目录地址：D:/fineReports/software

已处理文件目录地址：D:/fineReports/processed

在word文件夹下的文件命名要求：

输入为：文件名称+图文  文件名称+表格   输出为：文件名称.docx

示例：（输入）2.0通信图文.doc、2.0通信表格.doc   （输出）2.0通信.docx  

格式为rtf和doc的会在同目录下自动生成同名的docx文件

可以接受命令行参数：

需要按照 `图文文件` 、  `表格文件` 、 `输出文件` 的顺序输入绝对地址

### API调用程序 API_docx_process

接口地址：http://127.0.0.1:7777/process_doc

需要的参数：

- source：输入的图文文件和表格文件的绝对路径，表格文件名称中必须有**表格**字样，采用半角的 **`,`** 来分隔
- tar：输出的文件的绝对路径

注意：文件的格式必须是`docx`格式文件
