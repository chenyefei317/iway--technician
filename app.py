import datetime
import io
import os
import re
import zipfile
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, RGBColor, Pt
import pandas as pd
import qrcode
from PIL import Image
import requests
import streamlit as st
from streamlit_drawable_canvas import st_canvas

# ================= 1. 页面配置与初始化 =================
st.set_page_config(
    page_title="员工职业危害告知书和转岗培训记录表签收平台",
    layout="centered",
    initial_sidebar_state="expanded",
)

# 注入华文宋体全局样式与隐藏默认元素
hide_streamlit_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    html, body, [class*="css"] {
        font-family: "华文宋体", SimSun, serif;
    }
    </style>
    """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)


# 智能模糊查找 .docx 文件通用函数
def find_docx_file(folder, keyword1, keyword2):
  if not os.path.exists(folder):
    return None
  for filename in os.listdir(folder):
    if (
        keyword1 in filename
        and keyword2 in filename
        and filename.lower().endswith(".docx")
    ):
      return os.path.join(folder, filename)
  return None


# ================= 2. 百度网盘自动上传函数 =================
def upload_to_baidu_netdisk(file_bytes, remote_filename):
  """通过百度网盘开放平台 API 自动上传文件到云端网盘

  需要在百度网盘开放平台申请应用获取 Access Token，
  并可将 Token 存放在 Streamlit 的 st.secrets 中。
  """
  try:
    # 优先从 Streamlit Secrets 获取 Token，若未配置则跳过并提示
    access_token = st.secrets.get("BAIDU_ACCESS_TOKEN", "")
    if not access_token:
      return (
          False,
          "未配置百度网盘 Access Token（可在后台 Secrets 中设置），文件已成功生成并保存在本地/ZIP下载中。",
      )

    # 百度网盘上传接口地址（目标路径设定在 /apps/慧瑞EHS合规档案/ 目录下）
    target_path = f"/apps/慧瑞EHS合规档案/{remote_filename}"
    upload_url = f"https://pan.baidu.com/rest/2.0/xpan/file?method=upload&access_token={access_token}&path={target_path}&uploadid=&file=1"

    files = {"file": (remote_filename, file_bytes)}
    response = requests.post(upload_url, files=files)
    result = response.json()

    if "errno" in result and result["errno"] == 0:
      return True, f"成功自动同步至百度网盘目录：{target_path}"
    else:
      err_msg = result.get("error_msg", "未知错误")
      return False, f"百度网盘上传失败: {err_msg}"
  except Exception as e:
    return False, f"上传云盘异常: {str(e)}"


# ================= 3. 侧边栏：Logo、微信分享与模板下载 =================
with st.sidebar:
  try:
    st.image("logo.png", width=160)
  except Exception:
    st.image(
        "https://upload.wikimedia.org/wikipedia/commons/thumb/c/c5/Ikea_logo.svg/800px-Ikea_logo.svg.png",
        width=160,
    )

  st.markdown("### 📱 微信扫码与分享")
  st.write("已自动关联您的云端网址，二维码将实时更新供手机扫码填报。")

  app_url = st.text_input(
      "应用公网链接 (URL):",
      value="https://iway--technician.streamlit.app",
      help=(
          "请在此处粘贴部署到 Streamlit Cloud 后的真实网址，二维码会随之改变"
      ),
  )

  if app_url:
    qr = qrcode.make(app_url)
    img_buffer = io.BytesIO()
    qr.save(img_buffer, format="PNG")
    st.image(
        Image.open(img_buffer), caption="微信扫码快速填报通道", width=160
    )
    st.info(
        "💡 **提示**：将上方链接复制并发送至微信工作群，员工即可手机端签收。"
    )

  st.markdown("---")
  st.markdown("### 📥 常用制度模板快捷下载")

  training_folder_dl = "员工转岗安全与职业健康培训记录表"
  training_path_dl = os.path.join(
      training_folder_dl, "员工转岗安全与职业健康培训记录表.docx"
  )
  if os.path.exists(training_path_dl):
    with open(training_path_dl, "rb") as ft:
      st.download_button(
          label="📄 转岗培训记录表模板",
          data=ft.read(),
          file_name="员工转岗安全与职业健康培训记录表.docx",
      )

  st.write("**职业危害告知书下载：**")
  side_company = st.selectbox(
      "选择公司：",
      ["安徽恒林", "大连宜家", "东莞时兴", "福建龙竹", "福建双翼"],
      key="side_comp",
  )
  side_stage = st.selectbox(
      "选择阶段：", ["上岗前", "在岗期间"], key="side_stage"
  )

  side_hazard_path = find_docx_file(
      "职业危害告知书", side_company, side_stage
  )
  if side_hazard_path and os.path.exists(side_hazard_path):
    with open(side_hazard_path, "rb") as fsh:
      st.download_button(
          label=f"📥 下载选中的告知书",
          data=fsh.read(),
          file_name=f"职业危害告知书 - {side_company}（{side_stage}）.docx",
          key="side_dl_hazard_btn",
      )
  else:
    st.warning("⚠️ 暂未找到该模板")

# ================= 4. 主界面逻辑（Logo 变大并在标题左侧） =================
col_logo, col_title = st.columns([1, 5])
with col_logo:
  try:
    st.image("logo.png", width=110)  # 调大 Logo 宽度至 110px
  except Exception:
    st.image(
        "https://upload.wikimedia.org/wikipedia/commons/thumb/c/c5/Ikea_logo.svg/800px-Ikea_logo.svg.png",
        width=110,
    )
with col_title:
  st.markdown("## 员工职业危害告知书和转岗培训记录表签收平台")

st.markdown(
    "请仔细阅读下方各项内容，勾选确认并在底部完成手写签收。系统将自动把您的亲笔签名嵌入对应的 Word 正式档案中。"
)

# 基础信息录入
st.subheader("1. 员工基本信息")
col1, col2 = st.columns(2)
with col1:
  emp_name = st.text_input("员工姓名 (必填)：")
with col2:
  emp_id = st.text_input(
      "身份证号 (必填，须满18位)：",
      help="请输入标准的 18 位中国居民身份证号码",
  )

st.write("---")
st.markdown("### 📂 待签收项目清单")

# --- 项目一：职业危害告知书（精准匹配 .docx 模板） ---
st.subheader("⚠️ 项目一：职业危害告知书")

col_c, col_s = st.columns(2)
with col_c:
  company_choice = st.selectbox(
      "选择公司/站点：",
      ["安徽恒林", "大连宜家", "东莞时兴", "福建龙竹", "福建双翼"],
  )
with col_s:
  stage_choice = st.selectbox("选择告知阶段：", ["上岗前", "在岗期间"])

hazard_version = f"职业危害告知书 - {company_choice}（{stage_choice}）"
hazard_folder = "职业危害告知书"

hazard_path = find_docx_file(hazard_folder, company_choice, stage_choice)

try:
  if hazard_path and os.path.exists(hazard_path):
    with open(hazard_path, "rb") as f:
      hazard_docx_data = f.read()
  else:
    raise FileNotFoundError
except FileNotFoundError:
  doc_temp = Document()
  doc_temp.add_heading(hazard_version, level=1)
  doc_temp.add_paragraph(
      f"【系统提示】在 '{hazard_folder}' 文件夹中未找到匹配的 '.docx'"
      " 文件，请确认已上传至 GitHub。"
  )
  temp_io = io.BytesIO()
  doc_temp.save(temp_io)
  hazard_docx_data = temp_io.getvalue()
  hazard_path = None

st.info(f"您当前查阅的是：【{hazard_version}】。")
st.download_button(
    label=f"📥 下载《{hazard_version}.docx》",
    data=hazard_docx_data,
    file_name=f"{hazard_version}.docx",
    mime=(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    ),
)
c_hazard = st.checkbox(
    f"【须确认】本人已阅读并充分了解《{hazard_version}》的相关职业危害与防护要求，承诺在工作中严格落实。"
)

# 附加查阅：全套《职业危害告知书》模板快捷下载折叠区
with st.expander("📚 附加查阅：全套《职业危害告知书》模板快捷下载专区"):
  st.write(
      "如需查阅或下载其他公司/阶段的职业危害告知书，可直接点击下方按钮："
  )
  if os.path.exists(hazard_folder):
    all_hazard_files = [
        f for f in os.listdir(hazard_folder) if f.lower().endswith(".docx")
    ]
    for hf in all_hazard_files:
      full_hf_path = os.path.join(hazard_folder, hf)
      with open(full_hf_path, "rb") as fh:
        st.download_button(
            label=f"📥 下载：{hf}",
            data=fh.read(),
            file_name=hf,
            key=f"dl_all_{hf}",
        )
  else:
    st.write("暂无其他告知书文件。")

st.write("---")

# --- 项目二：员工转岗安全与职业健康培训记录表 ---
st.subheader("🎓 项目二：员工转岗安全与职业健康培训记录表")
training_folder = "员工转岗安全与职业健康培训记录表"


def find_training_file(folder):
  if not os.path.exists(folder):
    return None
  for filename in os.listdir(folder):
    if "转岗安全与职业健康培训记录表" in filename and filename.endswith(
        ".docx"
    ):
      return os.path.join(folder, filename)
  return None


training_path = find_training_file(training_folder)

try:
  if training_path and os.path.exists(training_path):
    with open(training_path, "rb") as f:
      training_docx_data = f.read()
  else:
    raise FileNotFoundError
except FileNotFoundError:
  doc_temp2 = Document()
  doc_temp2.add_heading("员工转岗安全与职业健康培训记录表", level=1)
  doc_temp2.add_paragraph(
      "【系统提示】未在 '员工转岗安全与职业健康培训记录表' 文件夹中找到对应的"
      " '.docx' 文件。"
  )
  temp_io2 = io.BytesIO()
  doc_temp2.save(temp_io2)
  training_docx_data = temp_io2.getvalue()
  training_path = None

st.markdown("请查阅以下标准培训记录表文件：")
st.download_button(
    label="📥 下载《员工转岗安全与职业健康培训记录表.docx》",
    data=training_docx_data,
    file_name="员工转岗安全与职业健康培训记录表.docx",
    mime=(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    ),
)
c_training = st.checkbox(
    "【须确认】本人已完成《员工转岗安全与职业健康培训记录表》所含全部课程的学习，熟知岗位危险源与操作规程。"
)

# 手写签名板块
st.write("---")
st.subheader("✍️ 3. 员工手写签名与提交")
st.markdown("**请在下方手写板内签名：**")
canvas_result = st_canvas(
    stroke_width=4,
    stroke_color="#000000",
    background_color="#F8F9FA",
    height=260,
    width=650,
    drawing_mode="freedraw",
    key="canvas",
    return_image_data=True,
)

sign_date = st.date_input("签收日期：", datetime.date.today())

# ================= 5. 提交校验与生成带签名的 Word 档案 =================
if st.button(
    "📁 确认无误，一键签收并生成带签名的 Word 档案", use_container_width=True
):
  is_canvas_empty = canvas_result.image_data is None or (
      canvas_result.json_data is not None
      and len(canvas_result.json_data.get("objects", [])) == 0
  )

  # 身份证 18 位强制校验正则
  id_pattern = re.compile(r"^\d{17}[\dXx]$")

  if not emp_name.strip() or not emp_id.strip():
    st.error("❌ 拦截：请完整填写【员工姓名】与【身份证号】！")
  elif not id_pattern.match(emp_id.strip()):
    st.error(
        "❌ 拦截：身份证号必须为严格的 **18 位**数字（末尾可为大写 X）！"
    )
  elif not c_hazard:
    st.error("❌ 拦截：您必须勾选确认已阅读《职业危害告知书》！")
  elif not c_training:
    st.error(
        "❌ 拦截：您必须勾选确认已完成《员工转岗安全与职业健康培训记录表》！"
    )
  elif is_canvas_empty:
    st.warning("⚠️ 拦截：请在上方画板完成手写签名后再提交。")
  else:
    st.success(
        "✅ 签收成功！系统已成功加载 Word 模板并在文末追加了您的手写签名。"
    )

    # 提取手写签名图片
    signature_img = Image.fromarray(
        canvas_result.image_data.astype("uint8"), "RGBA"
    )
    sig_io = io.BytesIO()
    signature_img.save(sig_io, format="PNG")
    sig_io.seek(0)


    # 安全加载模板并追加签名的核心函数（紧凑排版，确保同一页显示，强制华文宋体）
    def append_signature_to_docx(template_path, default_title):
      if template_path and os.path.exists(template_path):
        try:
          doc = Document(template_path)
        except Exception:
          doc = Document()
          doc.add_heading(default_title, level=1)
          doc.add_paragraph("（提示：模板文件读取异常，此为生成的标准确认单）")
      else:
        doc = Document()
        doc.add_heading(default_title, level=1)
        doc.add_paragraph("（提示：未找到对应的 .docx 模板文件）")

      # 统一设置华文宋体
      for p in doc.paragraphs:
        for r in p.runs:
          r.font.name = "华文宋体"
          r.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")

      # 紧凑排版：减小上下边距，确保和正文留在同一页
      p_line = doc.add_paragraph(
          "--------------------------------------------------"
      )
      p_line.paragraph_format.space_before = Pt(2)
      p_line.paragraph_format.space_after = Pt(2)

      p_confirm = doc.add_paragraph()
      p_confirm.paragraph_format.space_before = Pt(0)
      p_confirm.paragraph_format.space_after = Pt(2)
      run_c = p_confirm.add_run(
          f"【员工签收确认】 姓名：{emp_name} | 身份证号：{emp_id} | 日期：{sign_date}\n"
          f"本人已仔细阅读并充分理解上述内容，承诺在工作中严格遵守各项安全防范及操作规程。"
      )
      run_c.font.name = "华文宋体"
      run_c.font.size = Pt(10.5)
      run_c.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")

      p_sig_label = doc.add_paragraph()
      p_sig_label.paragraph_format.space_before = Pt(0)
      p_sig_label.paragraph_format.space_after = Pt(2)
      run_s = p_sig_label.add_run("员工本人手写亲笔签名：")
      run_s.font.name = "华文宋体"
      run_s.font.size = Pt(10.5)
      run_s.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")

      doc.add_picture(sig_io, width=Inches(1.8))
      sig_io.seek(0)

      buffer = io.BytesIO()
      doc.save(buffer)
      buffer.seek(0)
      return buffer


    signed_hazard_buffer = append_signature_to_docx(
        hazard_path, f"{hazard_version} 签收单"
    )
    signed_training_buffer = append_signature_to_docx(
        training_path, "员工转岗安全与职业健康培训记录表 签收单"
    )

    # 尝试自动同步到百度网盘
    hazard_filename_cloud = f"{hazard_version}_{emp_name}_{emp_id[-4:]}_已签字.docx"
    training_filename_cloud = (
        f"员工转岗培训记录表_{emp_name}_{emp_id[-4:]}_已签字.docx"
    )

    success_h, msg_h = upload_to_baidu_netdisk(
        signed_hazard_buffer.getvalue(), hazard_filename_cloud
    )
    success_t, msg_t = upload_to_baidu_netdisk(
        signed_training_buffer.getvalue(), training_filename_cloud
    )

    if success_h or success_t:
      st.info(f"☁️ 云盘同步状态：\n- {msg_h}\n- {msg_t}")
    else:
      st.warning(f"☁️ 云盘同步提示：{msg_h}")

    # ZIP 打包下载流
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
      zip_file.writestr(
          f"{hazard_version}_{emp_name}_已签字.docx",
          signed_hazard_buffer.getvalue(),
      )
      zip_file.writestr(
          f"员工转岗安全与职业健康培训记录表_{emp_name}_已签字.docx",
          signed_training_buffer.getvalue(),
      )
      img_byte_arr = io.BytesIO()
      signature_img.save(img_byte_arr, format="PNG")
      zip_file.writestr(
          f"手写签名原图_{emp_name}_{sign_date}.png", img_byte_arr.getvalue()
      )

    zip_buffer.seek(0)

    st.markdown("---")
    st.success(
        "🎉 您的专属带签名 Word 合规档案已打包完毕，点击下方按钮即可下载！"
    )

    col_d1, col_d2, col_d3 = st.columns(3)
    with col_d1:
      st.download_button(
          label="📄 下载带签名的告知书 (.docx)",
          data=signed_hazard_buffer,
          file_name=f"{hazard_version}_{emp_name}_已签字.docx",
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
      )
    with col_d2:
      st.download_button(
          label="📄 下载带签名的培训表 (.docx)",
          data=signed_training_buffer,
          file_name=f"员工转岗培训记录表_{emp_name}_已签字.docx",
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
      )
    with col_d3:
      st.download_button(
          label="📥 一键打包下载全部 (.ZIP)",
          data=zip_buffer,
          file_name=f"安全合规档案_{emp_name}_{sign_date}.zip",
          mime="application/zip",
      )

    st.balloons()

# ================= 6. 底部版权与开发者声明 =================
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: gray; font-size: 14px;'>"
    "内部使用，严禁商业用途 | 开发者：陈野菲 Yefei"
    "</div>",
    unsafe_allow_html=True,
)
