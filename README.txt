
# 日产值与产成率（Gradio 稳健版）

## 运行
```bat
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
python app_gradio.py
```
访问：http://127.0.0.1:7901

## 步骤
1. 上传 **主表 Excel**（多 sheet）；
2. 上传 **净重台账**（仅用“交鸡日期”“净重”两列，跨 sheet 聚合）；
3. 上传 **物料映射**（`映射` sheet：两列【关键词】【部位】）；
4. 选择“日期”（仅按天）；
5. 点击“运行模型”，下载 **《今日产值和产成率.xlsx》**。
