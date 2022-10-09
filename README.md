# 上应大 第二课堂 报表生成

## 功能

就学校那个页面，找一个活动目前加没加分贼麻烦，并且找出一个类目下的加分详细列表更加是难于上青天的事情。

因此用python3写了这个小jio本，用于生成第二课堂的加分概况及活动申请详单的excel表格。若需要筛选，也可以直接在表格中操作。

## 使用方法

### 安装

本仓库有两种安装方式可供选择。

#### exe直接启动

仓库的release中已提供了直接使用pyinstaller生成的exe文件，下载后双击文件可直接启动。  
文件过大及启动速度过慢均属正常现象~~（大概，因为我也没找到什么好方法解决）~~。

#### 使用python启动

clone仓库后，可直接使用命令行启动。  
```bash
# 克隆库，实在不行下载zip也行，再不会手动复制粘贴也行
git clone https://github.com/Amazefcc233/SIT-genOAReport.git
cd SIT-genOAReport/
# 先安装库
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple/
# 然后直接启动即可
python main.py
```

### 使用步骤

0. 准备好`Edge/Chrome/Firefox`浏览器任一，登录`EasyConnect`。
1. 启动本脚本，选择所需使用的浏览器
2. 在弹出的浏览器中登入OA。本脚本不会窃取账号密码信息，请放心使用。
3. 返回控制台，按下回车键开始获取数据并等待报表生成
4. 生成完成后，程序所在文件夹将生成`[学号]-[时间]-[随机字符串].xls`的excel文件，该文件即为所生成的报表。

## 特别鸣谢

- [`上应小风筝`](https://github.com/SIT-kite/kite-app)（部分代码借鉴了里面的实现方式=-=）
