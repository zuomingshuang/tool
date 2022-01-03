import tkinter as tk
from tkinter.filedialog import askdirectory, askopenfile
import os
import pdf2docx
import fitz
import re
import threading

class PdfHandle(object):
    def convert_word(self,pdf_path,word_path):
        obj = pdf2docx.Converter(pdf_path)
        obj.convert(word_path)
        obj.close()

    def get_img(self,pdf_path,img_dir):
        checkXO = r"/Type(?= */XObject)"  # 使用正则表达式来查找图片
        checkIM = r"/Subtype(?= */Image)"
        doc = fitz.open(pdf_path)  # 打开pdf文件
        imgcount = 0  # 图片计数
        lenXREF = doc.xrefLength()  # 获取对象数量长度

        # 遍历每一个对象
        for i in range(1,lenXREF):
            text = doc.xrefObject(i)  # 定义对象字符串
            isXObject = re.search(checkXO, text)  # 使用正则表达式查看是否是对象
            isImage = re.search(checkIM, text)  # 使用正则表达式查看是否是图片
            if not isXObject or not isImage:  # 如果不是对象也不是图片，则continue
                continue
            imgcount += 1
            pix = fitz.Pixmap(doc, i)  # 生成图像对象
            name_str = pdf_path.split("/")[-1].split(".pdf")[0]
            new_name = "{0}_{1}.png".format(name_str,imgcount)  # 生成图片的名称
            if pix.n < 5:  # 如果pix.n<5,可以直接存为PNG
                pix.writePNG(os.path.join(img_dir, new_name))
            else:  # 否则先转换CMYK
                pix0 = fitz.Pixmap(fitz.csRGB, pix)
                pix0.writePNG(os.path.join(img_dir, new_name))
                pix0 = None
            pix = None  # 释放资源


class Gui:
    def __init__(self,root):
        self.root = root
        self.pdf_handle = PdfHandle()

    def set_window(self):
        self.L1 = tk.Label(self.root, text="PDF文件:", font=("微软雅黑",12), background="#CDBA96")
        self.L1.place(x=10,y=30)

        self.E1 = tk.Entry(self.root, width=50, font=("微软雅黑",12))
        self.E1.place(in_=self.L1, x=75, y=0)

        self.B1 = tk.Button(self.root, text="添加", font=("微软雅黑",12), background="blue",
                       foreground="white", command=self.add_file)
        self.B1.place(in_=self.E1, x=470, y=-7)

        self.L2 = tk.Label(self.root, text="保存到:", font=("微软雅黑",12), background="#CDBA96")
        self.L2.place(in_=self.L1, x=0, y=50)

        self.E2 = tk.Entry(self.root, width=50, font=("微软雅黑",12))
        self.E2.place(in_=self.L2, x=70, y=0)

        self.B2 = tk.Button(text="选择", font=("微软雅黑",12), background="blue",
                       foreground="white", command=self.select_dir)
        self.B2.place(in_=self.E2, x=470, y=-7)

        self.B3 = tk.Button(text="转word文件", font=("微软雅黑",12), background="blue",
                       foreground="white", command=self.convert)
        self.B3.place(in_=self.L2, x=0, y=100)

        self.B4 = tk.Button(text="提取图片", font=("微软雅黑",12), background="blue",
                       foreground="white", command=self.get_img)
        self.B4.place(in_=self.L2, x=200, y=100)

        self.B5 = tk.Button(text="重置", font=("微软雅黑", 12), background="blue",
                            foreground="white", command=self.reset)
        self.B5.place(in_=self.L2, x=400, y=100)

        self.T1 = tk.Text(self.root, font=("微软雅黑", 12), height=20, width=72, background="#C6C8C2")
        self.T1.place(in_=self.B3, x=0, y=50)

    def add_file(self):
        file_name = askopenfile().name
        # print(file_name)
        self.E1.insert(tk.END, f"{file_name}|")

    def select_dir(self):
        dir_name = askdirectory()
        # print(dir_name)
        if dir_name:
            self.E2.delete(0, tk.END)
        self.E2.insert(0,dir_name)

    def convert(self):
        def convert_word():
            file_list = self.E1.get().split("|")[0:-1]
            for file in file_list:
                if not os.path.exists(file):
                    self.T1.insert(0.0, f"error: 文件'{file}' 不存在！转word文件失败\n")
                    continue
                if file.endswith(".pdf"):
                    if self.E2.get():
                        word_name = file.split("/")[-1].replace(".pdf",".docx")
                        word_path = os.path.join(self.E2.get(),word_name)
                        try:
                            self.pdf_handle.convert_word(pdf_path=file, word_path=word_path)
                            self.T1.insert(0.0, f"文件'{file}' 转换word文件成功\n")
                        except Exception as e:
                            self.T1.insert(0.0, f"error: {e} 文件'{file}' 转换word文件失败!\n")
                    else:
                        # print("请选择保存到的文件夹！")
                        self.T1.insert(0.0,"请选择保存到的文件夹！\n")
                        break
                else:
                    self.T1.insert(0.0, f"error: 文件'{file}' 不是pdf文件！转word文件失败\n")

        thd = threading.Thread(target=convert_word)
        thd.start()

    def get_img(self):
        def my_get_img():
            file_list = self.E1.get().split("|")[0:-1]
            img_dir = self.E2.get()
            for file in file_list:
                if not os.path.exists(file):
                    self.T1.insert(0.0, f"error: 文件'{file}' 不存在！提取图片失败\n")
                    continue
                if file.endswith(".pdf"):
                    if img_dir:
                        try:
                            self.pdf_handle.get_img(pdf_path=file, img_dir=img_dir)
                            self.T1.insert(0.0, f"文件'{file}' 提取图片成功\n")
                        except Exception as e:
                            self.T1.insert(0.0, f"error: {e} 文件'{file}' 提取图片失败!\n")
                    else:
                        # print("请选择保存到的文件夹！")
                        self.T1.insert(0.0, "请选择保存到的文件夹！\n")
                        break
                else:
                    self.T1.insert(0.0, f"error: 文件'{file}' 不是pdf文件！提取图片失败\n")

        thd = threading.Thread(target=my_get_img)
        thd.start()

    def reset(self):
        self.E1.delete(0,tk.END)
        self.E2.delete(0, tk.END)
        self.T1.delete(0.0, tk.END)


def main():
    root = tk.Tk()
    root.title("PDF tool")
    root.geometry("700x700+10+10")
    gui = Gui(root)
    gui.set_window()
    root.mainloop()


if __name__ == "__main__":
    # pdf_handle = PdfHandle()
    # pdf_handle.convert_word(pdf_path="test.pdf",word_path="test.docx")
    # pdf_handle.get_img(pdf_path="D:/scrips/tool/test.pdf",img_dir="img")
    main()