import os
import math
import pandas as pd
from flask import Flask, request, render_template_string, session, send_from_directory
import tempfile
import json

# 设置上传及结果保存目录
UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")
RESULT_FOLDER = os.path.join(os.getcwd(), "results")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["RESULT_FOLDER"] = RESULT_FOLDER
app.secret_key = "your_secret_key_here"  # 请替换为安全的密钥

# ==================== 业务逻辑函数 ====================
def process_files(file_caigoujia, file_qianzhi, file_gangkou, file_haiyun, file_duandao):
    """
    业务逻辑说明：
    1. 从“路径规划-采购价.xlsx”和“路径规划-前置运输.xlsx”中计算出各矿山到海港的“到海港价格”
       （支持矿山直达海港以及经铁路港中转）。
    2. 再结合“路径规划-港口基础信息.xlsx”、“路径规划-海运费.xlsx”和“路径规划-短倒费.xlsx”，计算
       各发货港/发货码头到卸货港/卸货码头组合的费用：
         - 非平仓模式：总费用 = 到海港价格 + 发货码头费 + 卸货码头费 + 海运费 + 附加费用
         - 平仓模式：总费用 = 平仓价 + 发货码头费 + 卸货码头费 + 海运费 + 附加费用
       注意：无论矿山直达还是中转，发货相关费用均根据匹配到的候选发货港记录填入（即“发货港”、“发货码头”、“卸运方式”等）。
    3. 输出时：
         - 对非平仓记录（“是否平仓价发货”为“否”），以【矿山, 卸货码头】为分组，仅保留总费用最低的一条；
         - 对平仓记录（“是否平仓价发货”为“是”），以【发货港, 卸货码头】为分组，仅保留总费用最低的一条。
    """
    # 读取各个 Excel 文件
    df_caigoujia = pd.read_excel(file_caigoujia)
    df_qianzhi   = pd.read_excel(file_qianzhi)
    df_gangkou   = pd.read_excel(file_gangkou)
    df_gangkou["最大吨位（万吨)"] = pd.to_numeric(df_gangkou["最大吨位（万吨)"], errors="coerce")
    df_haiyun    = pd.read_excel(file_haiyun)
    df_haiyun["海运船吨数"] = pd.to_numeric(df_haiyun["海运船吨数"], errors="coerce")
    df_duandao   = pd.read_excel(file_duandao)

    # ---------------- 输出列顺序 ----------------
    # 注意：根据要求，“卸货码头”在“卸货港”之后
    output_columns = [
        "采购价", "矿山",
        "运输方式1", "汽运价格", "铁路港", "运输方式2", "铁路价格",
        "海港", "到海港价格",
        "平仓价", "是否平仓价发货",
        "发货港", "发货码头", "卸运方式", "发货码头费",
        "卸货港", "卸货码头", "卸货码头费",
        "海运船吨数", "海运费",
        "附加终点", "附加费用",
        "总费用"
    ]

    def has_rail_in_route(*ways):
        return any(w == "铁路" for w in ways)

    # ---------------- PART A: 计算矿山→海港组合（得“到海港价格”） ----------------
    results_mine_to_sea = []
    for _, row_mine in df_caigoujia.iterrows():
        mine_name  = row_mine["矿山"]
        mine_price = row_mine["采购价"]
        # 筛选：起点为该矿山，且起点类型为“矿山”
        df_start = df_qianzhi[(df_qianzhi["起点"] == mine_name) & (df_qianzhi["起点类型"]=="矿山")]
        if df_start.empty:
            continue
        for _, r1 in df_start.iterrows():
            end_point = r1["终点"]
            end_type  = r1["终点类型"]
            way1      = r1["运输方式"]
            price1    = r1["运输价格"]
            qiyun_price = 0
            tielu_price = 0
            if way1 == "汽运":
                qiyun_price = price1
            elif way1 == "铁路":
                tielu_price = price1
            # 矿山直达海港
            if end_type == "海港":
                to_sea_price = mine_price + qiyun_price + tielu_price
                results_mine_to_sea.append({
                    "采购价": mine_price,
                    "矿山": mine_name,
                    "运输方式1": way1,
                    "汽运价格": qiyun_price,
                    "铁路港": "",
                    "运输方式2": "",
                    "铁路价格": tielu_price,
                    "海港": end_point,
                    "到海港价格": to_sea_price,
                    "平仓价": "",
                    "是否平仓价发货": "否",
                    "发货港": "",
                    "发货码头": "",
                    "卸运方式": "",
                    "发货码头费": 0,
                    "卸货港": "",
                    "卸货码头": "",
                    "卸货码头费": 0,
                    "海运船吨数": 0,
                    "海运费": 0,
                    "附加终点": "",
                    "附加费用": 0,
                    "总费用": to_sea_price
                })
            # 矿山经过铁路港到海港
            elif end_type == "铁路港":
                df_port2sea = df_qianzhi[
                    (df_qianzhi["起点"]==end_point) &
                    (df_qianzhi["起点类型"]=="铁路港") &
                    (df_qianzhi["终点类型"]=="海港")
                ]
                if df_port2sea.empty:
                    continue
                for _, r2 in df_port2sea.iterrows():
                    way2   = r2["运输方式"]
                    price2 = r2["运输价格"]
                    sea2   = r2["终点"]
                    seg_qiyun = qiyun_price
                    seg_tielu = tielu_price
                    if way2 == "汽运":
                        seg_qiyun += price2
                    elif way2 == "铁路":
                        seg_tielu += price2
                    to_sea_price = mine_price + seg_qiyun + seg_tielu
                    results_mine_to_sea.append({
                        "采购价": mine_price,
                        "矿山": mine_name,
                        "运输方式1": way1,
                        "汽运价格": seg_qiyun,
                        "铁路港": end_point,
                        "运输方式2": way2,
                        "铁路价格": seg_tielu,
                        "海港": sea2,
                        "到海港价格": to_sea_price,
                        "平仓价": "",
                        "是否平仓价发货": "否",
                        "发货港": "",
                        "发货码头": "",
                        "卸运方式": "",
                        "发货码头费": 0,
                        "卸货港": "",
                        "卸货码头": "",
                        "卸货码头费": 0,
                        "海运船吨数": 0,
                        "海运费": 0,
                        "附加终点": "",
                        "附加费用": 0,
                        "总费用": to_sea_price
                    })
    df_mine_to_sea = pd.DataFrame(results_mine_to_sea)

    # ---------------- PART B: 结合港口信息、海运费与短倒费 ----------------
    final_results = []
    # 将港口基础信息拆分为发货港部分与卸货港部分
    df_gangkou_fahuo  = df_gangkou[df_gangkou["港口类型"]=="发货港"]
    df_gangkou_xiehuo = df_gangkou[df_gangkou["港口类型"]=="卸货港"]

    # 根据运输情况：若使用铁路则只考虑火车堆场/火车直卸；否则只考虑汽运堆场/汽运直卸
    rail_unload_ways = {"火车堆场", "火车直卸"}
    road_unload_ways = {"汽运堆场", "汽运直卸"}

    for _, row_ms in df_mine_to_sea.iterrows():
        mine_price   = row_ms["采购价"]
        mine_name    = row_ms["矿山"]
        way1         = row_ms["运输方式1"]
        qiyun_price  = row_ms["汽运价格"]
        rail_port    = row_ms["铁路港"]
        way2         = row_ms["运输方式2"]
        tielu_price  = row_ms["铁路价格"]
        seaport      = row_ms["海港"]
        price_to_sea = row_ms["到海港价格"]

        used_rail = (way1=="铁路" or way2=="铁路")
        # 从发货港信息中，匹配港口名称等于 seaport 的候选记录
        df_fahuo_candidates = df_gangkou_fahuo[df_gangkou_fahuo["港口名称"]==seaport]
        if df_fahuo_candidates.empty:
            continue

        for _, row_fahuo in df_fahuo_candidates.iterrows():
            fahuo_matou = row_fahuo["码头名称"]
            fahuo_tonnage = row_fahuo.get("最大吨位（万吨)", math.inf)
            fahuo_pcj = row_fahuo.get("平仓价", 0)
            fahuo_fee = row_fahuo.get("码头费", 0)
            fahuo_unload_way = row_fahuo.get("卸运方式", None)
            if pd.notna(fahuo_unload_way) and fahuo_unload_way.strip() != "":
                if used_rail:
                    if fahuo_unload_way not in rail_unload_ways:
                        continue
                else:
                    if fahuo_unload_way not in road_unload_ways:
                        continue

            # 无论如何都将发货港信息填入输出，计算发货码头费
            df_haiyun_candidates = df_haiyun[
                (df_haiyun["发货港"]==seaport) &
                ((df_haiyun["发货码头"]==fahuo_matou) | (df_haiyun["发货码头"].isna()))
            ]
            if df_haiyun_candidates.empty:
                continue

            for _, row_hai in df_haiyun_candidates.iterrows():
                xiehuo_gang = row_hai["卸货港"]
                xiehuo_matou = row_hai["卸货码头"]
                ship_tonnage = row_hai["海运船吨数"]
                haiyun_fee = row_hai["海运费"]
                if pd.isna(ship_tonnage):
                    continue

                df_xiehuo_candidates = df_gangkou_xiehuo[
                    (df_gangkou_xiehuo["港口名称"]==xiehuo_gang) &
                    (df_gangkou_xiehuo["码头名称"]==xiehuo_matou)
                ]
                if df_xiehuo_candidates.empty:
                    continue

                for _, row_xh in df_xiehuo_candidates.iterrows():
                    xiehuo_fee = row_xh.get("码头费", 0)
                    xiehuo_tonnage = row_xh.get("最大吨位（万吨)", math.inf)
                    if pd.isna(fahuo_tonnage) or pd.isna(xiehuo_tonnage):
                        continue
                    if ship_tonnage > fahuo_tonnage or ship_tonnage > xiehuo_tonnage:
                        continue

                    # -------------- 处理“短倒费.xlsx” --------------
                    df_duan = df_duandao[df_duandao["卸货码头"]==xiehuo_matou]
                    if df_duan.empty:
                        add_target = ""
                        add_price  = 0
                        # 模式 A（非平仓）：总费用 = 到海港价格 + 发货码头费 + 卸货码头费 + 海运费 + 附加费用
                        totalA = price_to_sea + fahuo_fee + xiehuo_fee + haiyun_fee + add_price
                        rowA = {
                            "采购价": mine_price,
                            "矿山": mine_name,
                            "运输方式1": way1,
                            "汽运价格": qiyun_price,
                            "铁路港": rail_port,
                            "运输方式2": way2,
                            "铁路价格": tielu_price,
                            "海港": seaport,
                            "到海港价格": price_to_sea,
                            "平仓价": "",
                            "是否平仓价发货": "否",
                            "发货港": seaport,
                            "发货码头": fahuo_matou,
                            "卸运方式": fahuo_unload_way if pd.notna(fahuo_unload_way) else "",
                            "发货码头费": fahuo_fee,
                            "卸货港": xiehuo_gang,
                            "卸货码头": xiehuo_matou,
                            "卸货码头费": xiehuo_fee,
                            "海运船吨数": ship_tonnage,
                            "海运费": haiyun_fee,
                            "附加终点": add_target,
                            "附加费用": add_price,
                            "总费用": totalA
                        }
                        final_results.append(rowA)
                        if pd.notna(fahuo_pcj) and fahuo_pcj != 0:
                            totalB = fahuo_pcj + fahuo_fee + xiehuo_fee + haiyun_fee + add_price
                            rowB = {
                                "采购价": "",
                                "矿山": "",
                                "运输方式1": "",
                                "汽运价格": 0,
                                "铁路港": "",
                                "运输方式2": "",
                                "铁路价格": 0,
                                "海港": "",
                                "到海港价格": "",
                                "平仓价": fahuo_pcj,
                                "是否平仓价发货": "是",
                                "发货港": seaport,
                                "发货码头": fahuo_matou,
                                "卸运方式": fahuo_unload_way if pd.notna(fahuo_unload_way) else "",
                                "发货码头费": fahuo_fee,
                                "卸货港": xiehuo_gang,
                                "卸货码头": xiehuo_matou,
                                "卸货码头费": xiehuo_fee,
                                "海运船吨数": ship_tonnage,
                                "海运费": haiyun_fee,
                                "附加终点": add_target,
                                "附加费用": add_price,
                                "总费用": totalB
                            }
                            final_results.append(rowB)
                    else:
                        for _, drow in df_duan.iterrows():
                            add_target = drow["附加终端"] if "附加终端" in drow else drow["附加终点"]
                            add_price  = drow["附加价格"]
                            totalA = price_to_sea + fahuo_fee + xiehuo_fee + haiyun_fee + add_price
                            rowA = {
                                "采购价": mine_price,
                                "矿山": mine_name,
                                "运输方式1": way1,
                                "汽运价格": qiyun_price,
                                "铁路港": rail_port,
                                "运输方式2": way2,
                                "铁路价格": tielu_price,
                                "海港": seaport,
                                "到海港价格": price_to_sea,
                                "平仓价": "",
                                "是否平仓价发货": "否",
                                "发货港": seaport,
                                "发货码头": fahuo_matou,
                                "卸运方式": fahuo_unload_way if pd.notna(fahuo_unload_way) else "",
                                "发货码头费": fahuo_fee,
                                "卸货港": xiehuo_gang,
                                "卸货码头": xiehuo_matou,
                                "卸货码头费": xiehuo_fee,
                                "海运船吨数": ship_tonnage,
                                "海运费": haiyun_fee,
                                "附加终点": add_target,
                                "附加费用": add_price,
                                "总费用": totalA
                            }
                            final_results.append(rowA)
                            if pd.notna(fahuo_pcj) and fahuo_pcj != 0:
                                totalB = fahuo_pcj + fahuo_fee + xiehuo_fee + haiyun_fee + add_price
                                rowB = {
                                    "采购价": "",
                                    "矿山": "",
                                    "运输方式1": "",
                                    "汽运价格": 0,
                                    "铁路港": "",
                                    "运输方式2": "",
                                    "铁路价格": 0,
                                    "海港": "",
                                    "到海港价格": "",
                                    "平仓价": fahuo_pcj,
                                    "是否平仓价发货": "是",
                                    "发货港": seaport,
                                    "发货码头": fahuo_matou,
                                    "卸运方式": fahuo_unload_way if pd.notna(fahuo_unload_way) else "",
                                    "发货码头费": fahuo_fee,
                                    "卸货港": xiehuo_gang,
                                    "卸货码头": xiehuo_matou,
                                    "卸货码头费": xiehuo_fee,
                                    "海运船吨数": ship_tonnage,
                                    "海运费": haiyun_fee,
                                    "附加终点": add_target,
                                    "附加费用": add_price,
                                    "总费用": totalB
                                }
                                final_results.append(rowB)

    df_final = pd.DataFrame(final_results, columns=output_columns)

    # ---------------- PART C：分组仅保留最优记录 ----------------
    # 非平仓记录：以【矿山, 卸货码头】为分组依据
    df_non = df_final[df_final["是否平仓价发货"]=="否"]
    if not df_non.empty:
        df_non_opt = df_non.sort_values("总费用").groupby(["矿山", "卸货码头"], as_index=False).first()
    else:
        df_non_opt = pd.DataFrame(columns=df_final.columns)
    # 平仓记录：以【发货港, 卸货码头】为分组依据
    df_ping = df_final[df_final["是否平仓价发货"]=="是"]
    if not df_ping.empty:
        df_ping_opt = df_ping.sort_values("总费用").groupby(["发货港", "卸货码头"], as_index=False).first()
    else:
        df_ping_opt = pd.DataFrame(columns=df_final.columns)
    df_opt = pd.concat([df_non_opt, df_ping_opt], ignore_index=True)
    df_opt = df_opt.sort_values("总费用")
    # 重新按照 output_columns 顺序排序，确保“卸货码头”位于“卸货港”之后
    df_opt = df_opt[output_columns]

    return df_opt

# ==================== Web 部分 ====================
@app.route("/", methods=["GET"])
def index():
    page = '''
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="utf-8">
        <title>路径规划在线运算</title>
        <script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
      </head>
      <body>
        <h1>上传Excel文件以运算路径规划</h1>
        <div>
          <h3>路径规划-采购价.xlsx</h3>
          <input type="file" id="file_caigoujia" name="file_caigoujia">
          <button onclick="uploadFile('file_caigoujia')">点击上传</button>
          <div id="preview_file_caigoujia"></div>
        </div>
        <hr>
        <div>
          <h3>路径规划-前置运输.xlsx</h3>
          <input type="file" id="file_qianzhi" name="file_qianzhi">
          <button onclick="uploadFile('file_qianzhi')">点击上传</button>
          <div id="preview_file_qianzhi"></div>
        </div>
        <hr>
        <div>
          <h3>路径规划-港口基础信息.xlsx</h3>
          <input type="file" id="file_gangkou" name="file_gangkou">
          <button onclick="uploadFile('file_gangkou')">点击上传</button>
          <div id="preview_file_gangkou"></div>
        </div>
        <hr>
        <div>
          <h3>路径规划-海运费.xlsx</h3>
          <input type="file" id="file_haiyun" name="file_haiyun">
          <button onclick="uploadFile('file_haiyun')">点击上传</button>
          <div id="preview_file_haiyun"></div>
        </div>
        <hr>
        <div>
          <h3>路径规划-短倒费.xlsx</h3>
          <input type="file" id="file_duandao" name="file_duandao">
          <button onclick="uploadFile('file_duandao')">点击上传</button>
          <div id="preview_file_duandao"></div>
        </div>
        <hr>
        <button onclick="processFiles()">运算</button>
        <div id="process_result"></div>
        <script>
          function uploadFile(fieldId) {
            var input = document.getElementById(fieldId);
            if(input.files.length == 0) {
              alert("请选择文件后再点击上传");
              return;
            }
            var file = input.files[0];
            var formData = new FormData();
            formData.append("file", file);
            formData.append("fieldId", fieldId);
            $.ajax({
              url: "/upload_file",
              type: "POST",
              data: formData,
              processData: false,
              contentType: false,
              success: function(data) {
                $("#preview_" + fieldId).html(data.preview);
                localStorage.setItem(fieldId, data.filepath);
              },
              error: function(err) {
                alert("上传失败: " + err.responseText);
              }
            });
          }
          function processFiles() {
            var payload = {
              file_caigoujia: localStorage.getItem("file_caigoujia"),
              file_qianzhi: localStorage.getItem("file_qianzhi"),
              file_gangkou: localStorage.getItem("file_gangkou"),
              file_haiyun: localStorage.getItem("file_haiyun"),
              file_duandao: localStorage.getItem("file_duandao")
            };
            for (var key in payload) {
              if(!payload[key]) {
                alert("请确保所有文件均已上传");
                return;
              }
            }
            $.ajax({
              url: "/process",
              type: "POST",
              contentType: "application/json",
              data: JSON.stringify(payload),
              success: function(data) {
                $("#process_result").html(data);
              },
              error: function(err) {
                alert("运算失败: " + err.responseText);
              }
            });
          }
        </script>
      </body>
    </html>
    '''
    return render_template_string(page)

@app.route("/upload_file", methods=["POST"])
def upload_file():
    file = request.files.get("file")
    fieldId = request.form.get("fieldId")
    if not file or not fieldId:
        return {"error": "缺少文件或字段参数"}, 400
    filename = fieldId + "_" + file.filename
    filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    file.save(filepath)
    session[fieldId] = filepath
    try:
        df = pd.read_excel(filepath)
        preview_html = df.to_html(classes="table table-bordered", index=False, escape=False)
    except Exception as e:
        preview_html = "<p>读取Excel文件出错: " + str(e) + "</p>"
    return {"preview": preview_html, "filepath": filepath}

@app.route("/process", methods=["POST"])
def process_endpoint():
    data = request.get_json()
    file_caigoujia = data.get("file_caigoujia")
    file_qianzhi   = data.get("file_qianzhi")
    file_gangkou   = data.get("file_gangkou")
    file_haiyun    = data.get("file_haiyun")
    file_duandao   = data.get("file_duandao")
    if not all([file_caigoujia, file_qianzhi, file_gangkou, file_haiyun, file_duandao]):
        return "请确保所有文件均已上传", 400
    try:
        df_result = process_files(file_caigoujia, file_qianzhi, file_gangkou, file_haiyun, file_duandao)
    except Exception as e:
        return "处理文件时发生错误: " + str(e), 500
    result_filepath = os.path.join(app.config["RESULT_FOLDER"], "路径规划-输出.xlsx")
    df_result.to_excel(result_filepath, index=False)
    session["result_file"] = result_filepath
    result_html = '''
      <h2>运算完成！</h2>
      <button onclick="window.location.href='/download_result'">下载结果</button>
      <button onclick="window.location.href='/view_result'">在线查看</button>
    '''
    return result_html

@app.route("/download_result", methods=["GET"])
def download_result():
    result_filepath = session.get("result_file")
    if not result_filepath or not os.path.exists(result_filepath):
        return "结果文件不存在", 404
    return send_from_directory(app.config["RESULT_FOLDER"], os.path.basename(result_filepath), as_attachment=True)

@app.route("/view_result", methods=["GET"])
def view_result():
    result_filepath = session.get("result_file")
    if not result_filepath or not os.path.exists(result_filepath):
        return "结果文件不存在", 404
    try:
        df = pd.read_excel(result_filepath)
        result_table = df.to_html(classes="table table-striped", index=False, escape=False, na_rep="")
    except Exception as e:
        result_table = "<p>读取结果文件出错: " + str(e) + "</p>"
    html_content = '''
    <html>
      <head>
        <meta charset="utf-8">
        <title>在线查看结果</title>
        <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
      </head>
      <body>
        <div class="container mt-4">
          <h1>运算结果</h1>
          ''' + result_table + '''
          <br>
          <a href="/">返回首页</a>
        </div>
      </body>
    </html>
    '''
    return render_template_string(html_content)

if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
