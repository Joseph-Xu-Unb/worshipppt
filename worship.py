import json
import os
import logging
import tkinter as tk
from tkinter import filedialog, messagebox
from pptx import Presentation
from pptx.util import Pt

# ================= 核心逻辑函数 =================

# 配置日志
logging.basicConfig(
    filename="worship_ppt.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)


def delete_all_slides(prs):
    """
    深度清理：不仅删除幻灯片 ID，还同步清理关系映射，防止提示修复
    """
    # 1. 清理幻灯片 ID 列表
    sldIdLst = prs.slides._sldIdLst
    for sldId in list(sldIdLst):
        sldIdLst.remove(sldId)

    # 2. 清理 Presentation 零件中的所有幻灯片关系
    # 这是防止“提示修复”的关键步骤
    for rel in list(prs.part.rels.values()):
        if "slide" in rel.target_ref:
            prs.part.drop_rel(rel.rId)


def split_lines(lines, n=4):
    """将数组按每页 n 行切分"""
    return [lines[i : i + n] for i in range(0, len(lines), n)]


def safe_fill(slide, text, idx_list):
    """
    安全填充占位符，适配不同的母版 Index
    """
    for ph in slide.placeholders:
        if ph.placeholder_format.idx in idx_list:
            ph.text = text
            return True
    return False


def add_paged_content(prs, layout_name, title, lines):
    """
    分页生成内容，继承母版字体（36号）
    """
    layout = next(
        (
            l
            for l in prs.slide_layouts
            if l.name.lower().strip() == layout_name.lower().strip()
        ),
        None,
    )
    if not layout:
        layout = prs.slide_layouts[0]

    chunks = split_lines(lines, 4)
    for i, chunk in enumerate(chunks):
        slide = prs.slides.add_slide(layout)
        full_title = f"{title}"
        # 填充标题 (通常为 0 或 11)

        if slide.shapes.title:
            slide.shapes.title.text = full_title
        else:
            safe_fill(slide, full_title, [0, 11])

        # 填充正文 (通常为 1 或 10)
        content_text = "\n".join(chunk)
        safe_fill(slide, content_text, [1, 10])

        # 填充页码 (Index 12)
        safe_fill(slide, str(len(prs.slides)), [12])


# ================= 自动化流程 =================


def generate_worship_ppt(template_path, data, output_path):
    logging.info(f"Start generating PPT. Template: {template_path}")
    warnings = []
    prs = Presentation(template_path)

    # 1. 彻底清空旧页面
    # delete_all_slides(prs)

    # 2. 封面页 (Layout: cover)
    # Check if a cover slide already exists before adding
    if not any(slide.slide_layout.name == "cover" for slide in prs.slides):
        cover_layout = next(l for l in prs.slide_layouts if l.name == "cover")
        prs.slides.add_slide(cover_layout)

    # Add the selected date to the cover slide
    if slide.shapes.title:
        slide.shapes.title.text += f"\n{selected_date}"  # Append the date to the title
    else:
        safe_fill(slide, selected_date, [0, 11])  # Use placeholder if no title exists

    # 3. 增加预备页和标语页
    for name in ["prepare", "Slogan"]:
        layout = next(
            (
                l
                for l in prs.slide_layouts
                if l.name.lower().strip() == name.lower().strip()
            ),
            None,
        )
        if layout:
            prs.slides.add_slide(layout)
        else:
            msg = f"Layout '{name}' not found."
            logging.warning(msg)
            warnings.append(msg)

    # 4. 构建敬拜次序列表 (带全局序号和 Hymn 序号)
    raw_items = [f"宣召经文  {data['call_to_worship']['title']}", "赞美祷告  Prayer"]
    # 动态添加赞美诗
    for i, hymn in enumerate(data["hymns"]):
        raw_items.append(f"赞美诗 {i+1}  {hymn['title']}")

    raw_items.extend(
        [
            "认罪祷告  Prayer",
            f"主题经文  {data['theme_scripture']['title']}",
            "讲道  Preaching",
            f"回应诗  {data['response_hymn']['title']}",
            "奉献  Offering",
            "三一颂  Doxology",
            "牧师祝福  Benediction",
            "事项报告  Announcements",
            "我今天为你祝福  Blessings",
        ]
    )

    # 转换为带序号的文本
    full_order_text = "\n".join(raw_items)

    # 5. 填充敬拜次序页 (使用新的标题和内容占位符)
    order_layout = next(l for l in prs.slide_layouts if l.name == "order")
    order_slide = prs.slides.add_slide(order_layout)
    safe_fill(order_slide, "敬拜次序", [0, 11])
    safe_fill(order_slide, full_order_text, [1, 10])

    # 6. 按流程生成具体页面
    # 宣召
    call_title = f"宣召经文：{data['call_to_worship']['title']}"
    add_paged_content(
        prs,
        "call_to_scripture",
        call_title,
        data["call_to_worship"]["lines"],
    )

    # praise prayer
    praise_layout = next(l for l in prs.slide_layouts if l.name == "praise_prayer")
    prs.slides.add_slide(praise_layout)
    # 赞美诗页 (带 Hymn 1, 2... 序号)
    for i, hymn in enumerate(data["hymns"]):
        hymn_display_title = f"Hymn{i+1}: {hymn['title']}"
        add_paged_content(prs, "Hymn", hymn_display_title, hymn["lines"])

    # Intercessory prayer
    intercessory_layout = next(
        (
            l
            for l in prs.slide_layouts
            if l.name.lower().strip() == "intercessory_prayer"
        ),
        None,
    )
    if intercessory_layout:
        prs.slides.add_slide(intercessory_layout)
    else:
        msg = "Layout 'intercessory_prayer' not found."
        logging.warning(msg)
        warnings.append(msg)
    # 主题经文
    add_paged_content(
        prs,
        "theme_scripture",
        f"主题经文：{data['theme_scripture']['title']}",
        data["theme_scripture"]["lines"],
    )

    # 讲道题目页
    s_layout = next(
        (l for l in prs.slide_layouts if l.name.lower().strip() == "sermon"), None
    )
    if s_layout:
        s_slide = prs.slides.add_slide(s_layout)
        if s_slide.shapes.title:
            s_slide.shapes.title.text = "讲道：绝境逢生"
    else:
        msg = "Layout 'sermon' not found."
        logging.warning(msg)
        warnings.append(msg)

    # 回应诗
    response_display_title = f"回应诗: {data['response_hymn']['title']}"
    add_paged_content(
        prs, "response", response_display_title, data["response_hymn"]["lines"]
    )

    # 7. 静态结尾页
    for name in [
        "Offering",
        "ode_to_the_Trinity",
        "benediction",
        "announcements",
        "wishyouwell",
    ]:
        layout = next(
            (
                l
                for l in prs.slide_layouts
                if l.name.lower().strip() == name.lower().strip()
            ),
            None,
        )
        if layout:
            prs.slides.add_slide(layout)
        else:
            msg = f"Layout '{name}' not found."
            logging.warning(msg)
            # 记录当前所有可用布局名称，方便排查
            available = [l.name for l in prs.slide_layouts]
            logging.warning(f"Available layouts in template: {available}")
            warnings.append(msg)

    prs.save(output_path)
    logging.info(f"PPT saved to {output_path}")
    return warnings


# ================= GUI 界面类 =================


class ChurchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("BSBC PPT 自动化工具 v5")
        self.root.geometry("550x400")
        self.tpl_path = None
        self.jsn_path = None

        # 界面布局
        tk.Label(
            root, text="Fredericton BSBC 崇拜助手", font=("Arial", 18, "bold"), pady=20
        ).pack()

        # 选择模板
        tk.Button(
            root, text="1. 选择 PPT 模板 (.pptx)", width=25, command=self.load_tpl
        ).pack(pady=10)
        self.lbl_tpl = tk.Label(root, text="未选择模板", fg="grey")
        self.lbl_tpl.pack()

        # 选择数据
        tk.Button(
            root, text="2. 选择 JSON 歌词数据", width=25, command=self.load_jsn
        ).pack(pady=10)
        self.lbl_jsn = tk.Label(root, text="未选择 JSON", fg="grey")
        self.lbl_jsn.pack()

        # 添加日期选择器
        tk.Label(root, text="3. 选择日期 (生成文件名)", font=("Arial", 10)).pack(pady=5)
        self.date_entry = tk.Entry(root, width=15, font=("Arial", 12))
        self.date_entry.pack(pady=5)
        self.date_entry.insert(0, "2026-03-07")  # 默认值为当前日期

        # 开始生成
        self.run_btn = tk.Button(
            root,
            text="🚀 生成完整崇拜 PPT",
            state="disabled",
            width=30,
            height=2,
            command=self.execute,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 11, "bold"),
        )
        self.run_btn.pack(pady=30)

    def load_tpl(self):
        path = filedialog.askopenfilename(filetypes=[("PowerPoint", "*.pptx")])
        if path:
            self.tpl_path = path
            self.lbl_tpl.config(text=os.path.basename(path), fg="black")
            self.check_ready()

    def load_jsn(self):
        path = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if path:
            self.jsn_path = path
            self.lbl_jsn.config(text=os.path.basename(path), fg="black")
            self.check_ready()

    def check_ready(self):
        if self.tpl_path and self.jsn_path:
            self.run_btn.config(state="normal")

    def execute(self):
        try:
            # 读取数据
            with open(self.jsn_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            # 获取用户输入的日期
            selected_date = self.date_entry.get()
            if not selected_date:
                raise ValueError("日期不能为空！")

            # 设定输出路径 (与 JSON 同目录)
            output_file = os.path.join(
                os.path.dirname(self.jsn_path), f"Worship_{selected_date}.pptx"
            )

            # 调用生成逻辑
            warnings = generate_worship_ppt(self.tpl_path, data, output_file)

            if warnings:
                warn_msg = "\n".join(warnings)
                messagebox.showwarning(
                    "完成 (含警告)",
                    f"PPT 已生成，但有以下问题：\n{warn_msg}\n\n保存在：\n{output_file}",
                )
            else:
                messagebox.showinfo("成功", f"PPT 已生成！\n\n保存在：\n{output_file}")
        except Exception as e:
            logging.error("Generation failed", exc_info=True)
            messagebox.showerror(
                "生成失败", f"错误详情：\n{str(e)}\n\n已记录到日志文件。"
            )


if __name__ == "__main__":
    root = tk.Tk()
    app = ChurchApp(root)
    root.mainloop()
