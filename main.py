import customtkinter 
from tkinter import messagebox,ttk
from tkinter import filedialog
from tkinter import CENTER,END,DISABLED,NORMAL
import docx
import tkinter.font as tkfont


customtkinter.set_appearance_mode("Dark")
customtkinter.set_default_color_theme("blue")


def read_tables(file_path):
    doc = docx.Document(file_path)
    tables = []
    for table in doc.tables:
        table_data = []
        for row in table.rows:
            row_data = [cell.text for cell in row.cells]
            table_data.append(row_data)
        tables.append(table_data)
    names = []

    for table in tables:
        for name in table:
            if name[1] not in ["","الاسم الثلاثي",'الاسم الثلاثي ']:
                names.append(name[1])
    return names



class PageNotFound(customtkinter.CTk):
    def __init__(self, name_not_found):
        self.name_not_found = name_not_found
        super().__init__()

        self.title("Mojtaba Al-Mayhay")
        self.minsize(850,655)
        self.geometry('850x650+360+110')
        #=============Work====================#
        heading_frame = customtkinter.CTkFrame(master=self,corner_radius=5)
        heading_frame.pack(padx=0,pady=0, ipadx=0, ipady=0,fill="x",anchor="n")

        label = customtkinter.CTkLabel(master=heading_frame, text="الاسماء الغير  مضافة",font=customtkinter.CTkFont(family="Robot", size=25, weight="bold"))
        label.pack(ipady=10)

        main_frame = customtkinter.CTkFrame(master=self,corner_radius=10,fg_color='transparent')
        main_frame.pack(padx=0,pady=0, ipadx=5, ipady=5,fill="both",expand=True)

        left_frame = customtkinter.CTkFrame(master=main_frame,corner_radius=5)
        left_frame.pack_forget()

        top_frame = customtkinter.CTkFrame(master=left_frame,corner_radius=5)
        top_frame.pack(padx=0,pady=0, ipadx=5, ipady=5,fill="both",side="top",expand=True)

        scrollbar = ttk.Scrollbar(top_frame)
        
        if name_not_found != []:
            scrollbar.pack(side="right", fill="y")#

        tre = ttk.Treeview(top_frame, columns=(1,2), show='headings', yscrollcommand=scrollbar.set)
        tre.heading(1, text='الاسم',anchor=CENTER)
        tre.column('1',anchor=CENTER)
        tre.heading(2, text='ت',anchor=CENTER)
        tre.column('2',anchor=CENTER)

        style = ttk.Style()
        style.theme_use("clam")
        _font = tkfont.nametofont('TkTextFont')
        _font.configure(size=30)
        style.configure('Treeview', rowheight=_font.metrics('linespace'))
        
        if name_not_found != []:
            tre.pack(padx=0,pady=0, ipadx=0, ipady=0,fill="both",expand=True)

        scrollbar.config(command=tre.yview)
        # دالة للتحكم بحجم سكول
        def on_mousewheel(event):
            tre.yview_scroll(int(-1*(event.delta/120)), "units")

        # ربط دالة التحكم بحجم سكول بعجلة الماوس
        tre.bind_all("<MouseWheel>", on_mousewheel)

        left_frame.pack(padx=15,pady=15, ipadx=0, ipady=0,fill="both",expand=True,side="left")

        def Get_Data_customers(event):
            selected_rowcu = tre.focus()
            item = tre.item(selected_rowcu)
            global data
            data = item['values']  

        # تغيير لون العناصر في الجدول
        tre.tag_configure('odd', background='#f0f0f0')
        tre.tag_configure('even', background='#ffffff')

        for i, row in enumerate(name_not_found):
            # if row not in ["","الاسم الثلاثي"]:
            if i % 2 == 0:
                tre.insert("",END,values=(row,i+1), tags='even')
            else:
                tre.insert("",END,values=(row,i+1), tags='odd')

        tre.bind('<<TreeviewSelect>>',Get_Data_customers)

        if name_not_found == []:
            label.configure(text="المهندس مجتبى المياحي")
            label_not_found = customtkinter.CTkLabel(master=top_frame, text="لا يوجد اسماء غير مضافة",font=customtkinter.CTkFont(family="Robot", size=25, weight="bold"))
            label_not_found.pack(padx=0,pady=0, ipadx=0, ipady=0,fill="both",expand=True)



class PageMain(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Mojtaba Al-Mayhay")
        self.minsize(300,450)
        self.maxsize(300,450)
        self.geometry('300x450')
        
        self.main_frame = customtkinter.CTkFrame(master=self,corner_radius=10)
        self.main_frame.pack(padx=10,pady=10, ipadx=5, ipady=5,fill="both",expand=True)
        
        self.label = customtkinter.CTkLabel(master=self.main_frame, text="برنامج بسيط \nيقوم بمهام متواضعة والا هو يحسب \nعدد الاسماء في الجدول \nويتحقق منها مع مقارنتها بالاسماء من جدول اخر",font=customtkinter.CTkFont(family="Robot", size=14, weight="bold"))
        self.label.pack(fill="both",expand=True)

        self.button = 0
        self.select_button1 = customtkinter.CTkButton(master=self.main_frame, text="1 تحديد الملف",font=customtkinter.CTkFont(family="Robot", size=15, weight="bold"),command=self.select_file1)
        self.select_button1.pack(padx=25,pady=10, ipadx=5, ipady=5,fill="both")
        
        self.select_button2 = customtkinter.CTkButton(master=self.main_frame, text="2 تحديد الملف",font=customtkinter.CTkFont(family="Robot", size=15, weight="bold"),command=self.select_file2)
        self.select_button2.pack(padx=25,pady=10, ipadx=5, ipady=5,fill="both")
        
        self.save_button = customtkinter.CTkButton(master=self.main_frame, text="بدء الحساب",state=DISABLED,font=customtkinter.CTkFont(family="Robot", size=15, weight="bold"),command=self.save_files)
        self.save_button.pack(padx=25,pady=30, ipadx=5, ipady=5,fill="both")
        

        self.label2 = customtkinter.CTkLabel(master=self.main_frame, text="المهندس مجتبى المياحي",font=customtkinter.CTkFont(family="Robot", size=15, weight="bold"))
        self.label2.pack(padx=10,pady=10, ipadx=5, ipady=5,fill="both",expand=True)

        self.data1 = []
        self.data2 = []
        
        
    
    def select_file1(self):
        try:
            filetypes = (('Word files', '*.docx'),)
            file = filedialog.askopenfilename(title="تحديد ملف الورد الاول",filetypes=filetypes)
            if file:
                type_files = ['docx']
                type_file = file.split(".")[1]
                
                if type_file in type_files:
                    data = read_tables(file)
                    self.data1 = data

                    self.select_button1.configure(text="تم تحميل الملف الاول بنجاح",fg_color="green",state=DISABLED)
                    self.button = self.button +1

                    if self.button == 2:
                        self.save_button.configure(state=NORMAL)
                    # return data1
                else:
                    messagebox.showerror("حدث حطأ","حدث خطأ : الملف غير مدعوم")
            else:
                messagebox.showerror("حدث خطأ","حدث خطأ : يرجى تحديد ملف الورد الاول")
        except Exception as e:
            messagebox.showerror("حدث خطأ",f"{e}")
        
    
    def select_file2(self):
        try:
            filetypes = (('Word files', '*.docx'),)
            file = filedialog.askopenfilename(title="تحديد ملف الورد الثاني",filetypes=filetypes)
            if file:
                type_files = ['docx']
                type_file = file.split(".")[1]
                
                if type_file in type_files:
                    data = read_tables(file)
                    self.data2 = data
                    
                    self.select_button2.configure(text="تم تحميل الملف الثاني بنجاح",fg_color="green",state=DISABLED)
                    self.button = self.button +1
                    if self.button == 2:
                        self.save_button.configure(state=NORMAL)
                    # return data2
                else:
                    messagebox.showerror("حدث حطأ","حدث خطأ : الملف غير مدعوم")
            else:
                messagebox.showerror("حدث خطأ","حدث خطأ : يرجى تحديد ملف الورد الثاني")
        except Exception as e:
            messagebox.showerror("حدث خطأ",f"{e}")


    def save_files(self):
        self.data1
        self.data2
        name_not_found = []
        for name in self.data2:
            if name not in self.data1:
                name_not_found.append(name)
        
        app.close_my_app()
        app.destroy()

  
        page = PageNotFound(name_not_found)
        page.mainloop()


    def close_my_app(self):
        self.label.destroy()
        self.select_button1.destroy()
        self.select_button2.destroy()
        self.save_button.destroy()
        self.label2.destroy()
        self.main_frame.destroy()


if __name__ == '__main__':
    app = PageMain()
    app.mainloop()
