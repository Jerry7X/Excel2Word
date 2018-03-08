该工具用于将excel里面自定义变量的值填充到word对应的变量位置中。
主要是方便熟悉excel的人员自定义公式计算，然后导入word。提高工作效率。

testdata的密码是特定人员的名字拼音。

打包成exe文件
(使用这种方式是为了避开pyinstaller命令的bug，参见https://stackoverflow.com/questions/31808180/installing-pyinstaller-via-pip-leads-to-failed-to-create-process)：

python "C:\Program Files (x86)\Python35-32\Scripts\pyinstaller-script.py" -F -w test.py

测试的python版本：
Python 3.5.2

依赖的库:
openpxl
python-docx
