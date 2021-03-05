# ProRename-VBA
ProRename批量重命名工具，采用VBA实现
## 功能：
1. 一次可添加多个文件夹；
2. 按正则表达式、文件修改时间、文件大小筛选目标文件；
3. 按原有序号、文件修改时间等排序目标文件（开发中）；
4. 按正则表达式对原文件名进行截去和替换；
5. 可添加字符串、日期格式、序号、随机字符串等到新的文件名。
6. 预览后执行重命名。

## 执行流程：

1. 扫描目标文件夹的所有文件
2. 根据筛选规则筛选目标文件
3. 根据排序规则排序目标文件
4. 根据截取和替换规则对目标文件名进行处理
5. 根据增加模式及表达式对目标文件名进行处理
6. 生成预览重命名列表


## 界面总览
![image](https://user-images.githubusercontent.com/34180899/109485385-978d6b00-7abc-11eb-88d6-7ef56fdba24b.png)

### 添加文件来源
![image](https://user-images.githubusercontent.com/34180899/109486744-3c5c7800-7abe-11eb-8e5c-8bb537d02c80.png)

### 添加筛选条件
![image](https://user-images.githubusercontent.com/34180899/109487304-ff44b580-7abe-11eb-8129-c60c54bb931d.png)

### 选择排序模式（TODO）
![image](https://user-images.githubusercontent.com/34180899/109487345-0e2b6800-7abf-11eb-9fbd-7685ac0ec171.png)

### 正则表达式确定截取和替换
![image](https://user-images.githubusercontent.com/34180899/109487403-20a5a180-7abf-11eb-9d5e-f9522ced1115.png)

### 添加要增加的字符模式
![image](https://user-images.githubusercontent.com/34180899/109487455-31eeae00-7abf-11eb-81d2-b48197ac8103.png)

### 拼接增加模式表达式
![image](https://user-images.githubusercontent.com/34180899/109487498-3dda7000-7abf-11eb-8b41-e486d8003b80.png)
