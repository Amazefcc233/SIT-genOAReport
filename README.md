# 上应大 第二课堂 报表生成

## 功能

就学校那个页面，找一个活动目前加没加分贼麻烦，并且找出一个类目下的加分详细列表更加是难于上青天的事情。

因此用python3写了这个小jio本，用于生成第二课堂的加分概况及活动申请详单的excel表格。若需要筛选，也可以直接在表格中操作。

## 使用方法

### 安装

本仓库有两种安装方式可供选择。

#### exe直接启动

仓库的[release](https://github.com/Amazefcc233/SIT-genOAReport/releases/latest)中已提供了直接使用pyinstaller（`pyinstaller main.py -F`）生成的exe文件，下载后双击文件可直接启动。  
文件过大及启动速度过慢均属正常现象~~（大概，因为我也没找到什么好方法解决）~~。  
如果您怀疑文件有毒或会偷偷上传个人信息，欢迎您一刻都不要使用该文件。  

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

#### 调试模式

如果在生成报告时遇到问题，可以尝试使用调试模式。

在调试模式下，程序所在目录将保存`data.txt`文件。**该文件包含了第二课堂记录的原始详细记录。文件中已尝试将您的学号与姓名去除，但无法排除因bug导致去除失败的可能性。请务必留意您的信息安全。**

如需启动，在启动命令后加上`--save-file`参数即可。（exe直接启动请在创建快捷方式后添加参数或使用cmd启动并附加参数。）

如：

```bash
# python启动
python main.py --save-file
# exe启动，首先在文件夹内按住shift+右键空白处，选择“在此处打开命令行窗口”
main.exe --save-file
```

### 使用步骤

0. 准备好`Edge/Chrome/Firefox`浏览器任一，登录`EasyConnect`。
1. 启动本脚本，选择所需使用的浏览器。
2. 在弹出的浏览器中登入OA。本脚本不会窃取账号密码信息，请放心使用。（如果你有任何疑虑，欢迎随时查看源代码或随时删除下载的所有文件。）
3. 返回控制台，按下回车键开始获取数据并等待报表生成。
4. 生成完成后，程序所在文件夹将生成`[学号]-[时间]-[随机字符串].xls`的excel文件，该文件即为所生成的报表。

## 注意事项

1. 程序存在生成失败的可能性，具体原因暂未知（无法稳定复现）。如果生成失败，请尝试重试。若多次仍然失败，请尝试使用调试模式生成文件，并在隐去个人信息后携带data.txt提交issue。
2. Edge理论只支持新版edge（chromium内核），旧版edge未测试。Chrome/Firefox仅在开发初期进行过测试。

## 特别鸣谢

- [`上应小风筝`](https://github.com/SIT-kite/kite-app)（部分代码借鉴了里面的实现方式=-=）
