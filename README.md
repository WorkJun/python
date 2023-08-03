# 环境配置
1.报错ssl的重新安装 pip install urllib3==1.26.8

2.少包的下载包 pip install xxx

# 打包为可执行文件
1.pyinstaller -F export_excel.py devConfig.py prdConfig.py -n 导出投料卡


# 项目目录
## 1.export
导出excel需求代码

#### 1.[导出投料卡报货](export%2Fexport_excel.py)
## 2.log
监听canalLog异常。启动命令：nohup python canal-log.py > /dev/null 2>&1 &
[canal-log.py](log%2Fcanal-log.py)

## 3.excel
#### 1.[对比excel差异](excel%2FdiffExcel.py)
#### 2.[解析message服务发送飞书消息并提取相应字段](excel%2FparsingExcel.py)

