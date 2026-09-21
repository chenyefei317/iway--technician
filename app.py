import datetime
import io
import os
import zipfile
from docx import Document
from docx.oxml.ns import qn
from docx.shared import RGBColor, Pt
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
  st.write("请使用手机扫描下方二维码，直接在手机端完成材料查阅与签收。")

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
    "请仔细阅读下方【职业危害告知书】与【员工转岗安全与职业健康培训记录表】各项内容，勾选确认并在底部完成手写签收。"
)

# 基础信息录入
st.subheader("1. 员工基本信息")
col1, col2 = st.columns(2)
with col1:
  emp_name = st.text_input("员工姓名 (必填)：")
with col2:
  emp_id = st.text_input("工号/身份证号 (必填)：")

st.write("---")
st.markdown("### 📂 待签收项目清单")

# --- 项目一：职业危害告知书（从“职业危害告知书”文件夹读取真实文件） ---
st.subheader("⚠️ 项目一：职业危害告知书")
hazard_options = [
    "职业危害告知书 - 大连宜家",
    "职业危害告知书 - 福建双翼",
    "职业危害告知书 - 福建龙竹",
    "职业危害告知书 - 安徽恒林",
    "职业危害告知书 - 东莞时兴",
]
hazard_version = st.selectbox(
    "请选择【职业危害告知书】对应版本/站点：", hazard_options
)

# 动态拼接文件夹路径与文件名
hazard_folder = "职业危害告知书"
hazard_filename = f"{hazard_version}.pdf"
hazard_path = os.path.join(hazard_folder, hazard_filename)

# 尝试读取真实的 PDF 文件
try:
  with open(hazard_path, "rb") as f:
    hazard_pdf_data = f.read()
  file_ready_1 = True
except FileNotFoundError:
  hazard_pdf_data = (
      f"提示：未在 '{hazard_folder}' 文件夹中找到 '{hazard_filename}'"
      " 文件，请确认已上传至 GitHub。"
  ).encode("utf-8")
  file_ready_1 = False

st.info(f"您当前查阅的是：【{hazard_version}】。")
st.download_button(
    label=f"📥 下载《{hazard_version}.pdf》",
    data=hazard_pdf_data,
    file_name=hazard_filename,
    mime="application/pdf",
)
c_hazard = st.checkbox(
    f"【须确认】本人已阅读并充分了解《{hazard_version}》的相关职业危害与防护要求，承诺在工作中严格落实。"
)

st.write("---")

# --- 项目二：员工转岗安全与职业健康培训记录表（从对应文件夹读取真实文件） ---
st.subheader("🎓 项目二：员工转岗安全与职业健康培训记录表")
training_folder = "员工转岗安全与职业健康培训记录表"
training_filename = "员工转岗安全与职业健康培训记录表.pdf"
training_path = os.path.join(training_folder, training_filename)

try:
  with open(training_path, "rb") as f:
    training_pdf_data = f.read()
  file_ready_2 = True
except FileNotFoundError:
  training_pdf_data = (
      f"提示：未在 '{training_folder}' 文件夹中找到 '{training_filename}'"
      " 文件，请确认已上传至 GitHub。"
  ).encode("utf-8")
  file_ready_2 = False

st.markdown("请查阅以下唯一的标准培训记录表文件：")
st.download_button(
    label="📥 下载《员工转岗安全与职业健康培训记录表.pdf》",
    data=training_pdf_data,
    file_name=training_filename,
    mime="application/pdf",
)
c_training = st.checkbox(
    "【须确认】本人已完成《员工转岗安全与职业健康培训记录表》所含全部课程的学习，熟知岗位危险源与操作规程。"
)

# 手写签名板块
st.write("---")
st.subheader("✍️ 3. 员工手写签名与提交")
st.markdown(
    "**请在下方空白处手写签名（签名将自动嵌入生成的 Word 告知书与培训表中）：**"
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

# ================= 4. 提交校验与生成导出（带自动嵌入签名逻辑） =================
if st.button(
    "📁 确认无误，一键签收并生成带签名的归档文件", use_container_width=True
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
        "✅ 签收成功！系统已将您的手写签名成功嵌入 Word 文档，请打包下载。"
    )

    # 提取手写签名图片并转存为内存二进制流
    signature_img = Image.fromarray(
        canvas_result.image_data.astype("uint8"), "RGBA"
    )
    sig_io = io.BytesIO()
    signature_img.save(sig_io, format="PNG")
    sig_io.seek(0)


    # 辅助函数：生成带签名的 Word 文档
    def create_signed_word(doc_title, content_text):
      doc = Document()
      style = doc.styles["Normal"]
      style.font.name = "华文宋体"
      style.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")

      h1 = doc.add_heading(level=1)
      run_h1 = h1.add_run(doc_title)
      run_h1.bold = True
      run_h1.font.name = "华文宋体"
      run_h1.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")

      p_info = doc.add_paragraph()
      run_info = p_info.add_run(
          f"员工姓名：{emp_name}    工号/身份证：{emp_id}    签收日期：{sign_date}"
      )
      run_info.font.name = "华文宋体"
      run_info.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")

      doc.add_paragraph("--------------------------------------------------")

      p_content = doc.add_paragraph()
      run_c = p_content.add_run(content_text)
      run_c.font.name = "华文宋体"
      run_c.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")

      doc.add_paragraph("\n")

      p_sig = doc.add_paragraph()
      run_s = p_sig.add_run("员工本人手写签名确认：\n")
      run_s.font.name = "华文宋体"
      run_s.font.element.rPr.rFonts.set(qn("w:eastAsia"), "华文宋体")

      doc.add_picture(sig_io, width=Inches(2.2))
      sig_io.seek(0)

      buffer = io.BytesIO()
      doc.save(buffer)
      buffer.seek(0)
      return buffer


    # 生成带签名的职业危害告知书 Word
    hazard_text = (
        f"本人（{emp_name}，工号：{emp_id}）已仔细阅读并充分了解【{hazard_version}】的相关职业危害与防护要求。"
        "明确知晓工作场所中存在的有害因素及其对健康的潜在影响，承诺在日常作业中严格遵守各项 EHS 安全操作规程，"
        "并按规定正确佩戴和使用个人劳动防护用品（PPE）。"
    )
    doc_hazard_buffer = create_signed_word(
        f"{hazard_version} 签收确认单", hazard_text
    )

    # 生成带签名的转岗培训记录表 Word
    training_text = (
        f"本人（{emp_name}，工号：{emp_id}）已正式完成《员工转岗安全与职业健康培训记录表》所含的全部课程学习。"
        "已全面熟知新岗位/新区域的重大危险源、应急救援措施、消防安全及环保职业健康管理要求，"
        "经考核合格，具备上岗操作资格。"
    )
    doc_training_buffer = create_signed_word(
        "员工转岗安全与职业健康培训记录表（签收确认）", training_text
    )

    # ZIP 打包下载流
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
      zip_file.writestr(
          f"{hazard_version}_{emp_name}_已签字.docx",
          doc_hazard_buffer.getvalue(),
      )
      zip_file.writestr(
          f"员工转岗安全与职业健康培训记录表_{emp_name}_已签字.docx",
          doc_training_buffer.getvalue(),
      )
      img_byte_arr = io.BytesIO()
      signature_img.save(img_byte_arr, format="PNG")
      zip_file.writestr(
          f"手写签名原图_{emp_name}_{sign_date}.png", img_byte_arr.getvalue()
      )

    zip_buffer.seek(0)

    st.markdown("---")
    st.success("🎉 您的专属带签名合规档案已打包完毕，点击下方按钮即可下载！")

    col_d1, col_d2, col_d3 = st.columns(3)
    with col_d1:
      st.download_button(
          label="📄 下载带签名的告知书",
          data=doc_hazard_buffer,
          file_name=f"{hazard_version}_{emp_name}_已签字.docx",
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
      )
    with col_d2:
      st.download_button(
          label="📄 下载带签名的培训表",
          data=doc_training_buffer,
          file_name=f"员工转岗培训记录表_{emp_name}_已签字.docx",
          mime=(
              "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          ),
      )
    with col_d3:
      st.download_button(
          label="📥 一键打包下载全部 (ZIP)",
          data=zip_buffer,
          file_name=f"员工安全与职业健康全套档案_{emp_name}_{sign_date}.zip",
          mime="application/zip",
      )

    st.balloons()
