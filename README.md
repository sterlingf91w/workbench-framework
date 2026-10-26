# 销售工作台

这是工作台框架。里面没有公司名，也没有客户数据。

和框架放在同一层时，目录是：

- `客户数据原稿/`：拜访原件和整理稿
- `个人客户数据/客户数据.xlsx`：正在使用的表，旁边有备份、日报和复盘
- `工作台框架/`：只有这份代码

启动时框架发现旁边的个人表，就直接读写它，页面和这张表合成完整工作台。旁边没有个人表时，用 `template/空白跟进表.xlsx` 在框架内的 `data/客户数据.xlsx` 生成空表。`data/` 已被忽略。

## 运行

```bash
python3 -m pip install -r requirements.txt
python3 server.py --open
```

浏览器打开 http://127.0.0.1:8765 。也可以双击 `启动工作台.command`。

指定别的表：`python3 server.py --data /路径/客户数据.xlsx --open`。

## 检查

```bash
python3 -m unittest test_app.py
node test_ui.cjs
```

测试只用示例数据，不读取个人表。
