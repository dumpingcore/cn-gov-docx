from cn_gov_docx import CnGovDocx


def test_create_docx():
    doc = CnGovDocx()

    doc.add_title("测试文档")
    doc.add_h1("一级标题")
    doc.add_h2("二级标题")
    doc.add_h3("三级标题")
    doc.add_text("这里是正文内容，用于测试生成能力。")

    doc.save("test.docx")


if __name__ == "__main__":
    test_create_docx()
