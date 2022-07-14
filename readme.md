##运行前请执行一下命令行安装所需依赖库：

  	pip install -r requirements.txt


#使用方法：

​	1.按照input文件夹.xlsx文件格式书写需要转化的xlsx文件并添入该文件夹

​	2.第一次运行会生成按 

​				”语言\组件名.json“ 

​		的文件结构

​	3.第二次运行则会额外生成 

​				语言\old\old组件名.json 以及 

​				语言\old\change\change组件名.json

​		分别代表旧的json文件以及新旧对比的json文件
