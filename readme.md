##运行前请执行一下命令行安装所需依赖库：

  	pip install -r requirements.txt


#使用方法：
​	1.按照input文件夹.xlsx文件格式书写需要转化的xlsx文件并添入该文件夹
	运行并输入文件名前缀筛选(直接回车则默认选中所有xlsx文件)

​	2.第一次运行会生成按 

​				”语言\组件名.json“ 

​		的文件结构

​	3.第二次运行则会额外生成(按照outPut内上次生成的json文件作为生成以及对比依据)

​				语言\old\old组件名.json 以及 

​				语言\old\change\change组件名.json

​	分别代表旧的json文件以及新旧对比的json文件

 注：auto-py-to-exe  
