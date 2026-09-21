import datetime
import io
import os
import zipfile
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Inches, RGBColor, Pt
import pandas as pd
import qrcode
from PIL import Image
import streamlit as st
from streamlit_drawable_canvas import st_canvas

# ================= 1. 页面配置与初始化 =================
st.set_page_config(
    page_title="员工安全与职业健康签收平台",
    layout="centered",
    initial_sidebar_state="expanded",
)

# 隐藏默认元素
hide_streamlit_style = """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
    """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# ================= 2. 侧边栏：扫码通道 =================
with st.sidebar:
  st.header("📱 手机端/网页端扫码填报")
  st.write(
      "请使用手机扫描下方二维码，直接在手机端完成 Word 材料查阅与手写签收。"
  )

  app_url = "https://your-transfer-training.streamlit.app/"  # 部署后替换为您的云端链接

  qr = qrcode.QRCode(
      version=1,
      error_correction=qrcode.constants.ERROR_CORRECT_M,
      box_size=6,
      border=2,
  )
  qr.add_data(app_url)
  qr.make(fit=True)
  qr_img = qr.make_image(fill_color="black", back_color="white")

  buf = io.BytesIO()
  qr_img.save(buf, format="PNG")
  st.image(buf.getvalue(), caption="手机扫码快速签收通道")

# ================= 3. 主界面逻辑 =================
st.title("👨‍🔧 员工安全与职业健康签收平台")
st.markdown(
    "请查阅下方 Word 版本的【职业危害告知书】与【员工转岗安全与职业健康培训记录表】，勾选确认并在底部完成手写签收。系统将自动把您的签名嵌入原 Word 模板中。"
)

# 基础信息录入
st.subheader("1. 员工基本信息")
col1, col2 = st.columns(2)
with col1:
  emp_name = emp_name = st.text_input("员工姓名 (必填)：")
with col2:
  emp_id = st.text_input("工号/身份证号 (必填)：")

st.write("---")
st.markdown("### 📂 待签收项目清单（Word文档版）")

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


# 智能查找 .docx 文件
def find_docx_file(folder, keyword1, keyword2):
  if not os.path.exists(folder):
    return None
  for filename in os.listdir(folder):
    if (
        keyword1 in filename
        and keyword2 in filename
        and filename.endswith(".docx")
    ):
      return os.path.join(folder, filename)
  return None


hazard_path = find_docx_file(hazard_folder, company_choice, stage_choice)

try:
  if hazard_path and os.path.exists(hazard_path):
    with open(hazard_path, "rb") as f:
      hazard_docx_data = f.read()
    file_ready_1 = True
  else:
    raise FileNotFoundError
except FileNotFoundError:
  doc_temp = Document()
  doc_temp.add_heading(hazard_version, level=1)
  doc_temp.add_paragraph(
      f"【系统提示】未在 GitHub 的 '{hazard_folder}' 文件夹中找到对应的"
      " '.docx' 格式文件！\n请注意：必须上传 .docx 格式（不能是 .doc"
      " 格式），否则无法读取原文内容。"
  )
  temp_io = io.BytesIO()
  doc_temp.save(temp_io)
  hazard_docx_data = temp_io.getvalue()
  file_ready_1 = False

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
    file_ready_2 = True
  else:
    raise FileNotFoundError
except FileNotFoundError:
  doc_temp2 = Document()
  doc_temp2.add_heading("员工转岗安全与职业健康培训记录表", level=1)
  doc_temp2.add_paragraph(
      "【系统提示】未在 '员工转岗安全与职业健康培训记录表' 文件夹中找到对应的"
      " '.docx' 文件。\n请确认已将 .docx 格式的模板文件上传至 GitHub。"
  )
  temp_io2 = io.BytesIO()
  doc_temp2.save(temp_io2)
  training_docx_data = temp_io2.getvalue()
  file_ready_2 = False

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
st.markdown(
    "**请在下方空白处手写签名（签名与基本信息将自动嵌入并“盖章”到下载的"
    " Word 模板正文最下方）：**"
)
canvas_result = st_canvas(
    stroke_width=3,
    stroke_color="#000000",
    background_color="#F0F2F6",
    height=150,
    width=400,
    drawing_mode="freedraw",
    key="canvas",
    return_image_data=True,
)

sign_date = st.date_input("签收日期：", datetime.date.today())

# ================= 4. 提交校验与生成带签名的 Word 归档 =================
if st.button(
    "📁 确认无误，一键签收并生成带签名的 Word 档案", use_container_width=True
):
  is_canvas_empty = canvas_result.image_data is None or (
      canvas_result.json_data is not None
      and len(canvas_result.json_data.get("objects", [])) == 0
  )

  if not emp_name.strip() or not emp_id.strip():
    st.error("❌ 拦截：请完整填写【员工姓名】与【工号/身份证号】！")
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
        "✅ 签收成功！系统已成功加载您的 Word 原文模板，并在文末追加了您的手写签名。"
    )

    # 提取手写签名图片并转存为内存二进制流
    signature_img = Image.fromarray(
        canvas_result.image_data.astype("uint8"), "RGBA"
    )
    sig_io = io.BytesIO()
    signature_img.save(sig_io, format="PNG")
    sig_io.seek(0)


    # 核心函数：完美加载真实 Word 模板并在末尾追加签名
    def append_signature_to_docx(template_path, default_title):
      if template_path and os.path.exists(template_path):
        doc = Document(template_path)  # 完美加载您上传的 Word 模板原内容！
      else:
        doc = Document()
        doc.add_heading(default_title, level=1)
        doc.add_paragraph(
            "（提示：未找到对应 .docx 模板文件，此为系统生成的标准确认单）"
        )

      # 在文末追加签收确认信息与手写签名图片
      doc.add_paragraph("\n")
      doc.add_paragraph(
          "--------------------------------------------------"
      )

      p_confirm = doc.add_paragraph()
      run_c = p_confirm.add_run(
          f"【员工签收确认】\n员工姓名：{emp_name}    工号/身份证：{emp_id}    "
          f"签收日期：{sign_date}\n本人已仔细阅读并充分理解上述内容，承诺在工作中严格遵守各项安全防范及操作规程。"
      )
      run_c.font.name = "华文宋体"
      run_c.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")

      p_sig_label = doc.add_paragraph()
      run_s = p_sig_label.add_run("员工本人手写亲笔签名：")
      run_s.font.name = "华文宋体"
      run_s.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")

      # 将签名图片直接插入 Word 文档末尾
      doc.add_picture(sig_io, width=Inches(2.2))
      sig_io.seek(0)  # 重置指针

      buffer = io.BytesIO()
      doc.save(buffer)
      buffer.seek(0)
      return buffer


    # 1. 生成带签名的职业危害告知书 Word
    signed_hazard_buffer = append_signature_to_docx(
        hazard_path, f"{hazard_version} 签收单"
    )

    # 2. 生成带签名的转岗培训记录表 Word
    signed_training_buffer = append_signature_to_docx(
        training_path, "员工转岗安全与职业健康培训记录表 签收单"
    )

    # 3. ZIP 打包下载流
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
          file_name=f"员工安全与职业健康全套档案_{emp_name}_{sign_date}.zip",
          mime="application/zip",
      )

    st.balloons()
