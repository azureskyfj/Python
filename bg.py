import  os, sys, shutil, glob
from tkinter import messagebox, Frame, Label, StringVar, Entry, Button

class Application(Frame):
    def createWidgets(self):
        #messagebox.showinfo("Alert", "使用前关闭WORD文档")
        self.Label_num_moss = Label(self, text="MOSS顺序号：")
        self.Label_num_bg = Label(self, text="变更号：")
        self.Label_nam_system = Label(self, text="系统名称：")
        self.Label_nam_apply = Label(self, text="变更提交人：")
        self.Label_nam_dep = Label(self, text="变更提交部门：")
        self.Label_nam_manager = Label(self, text="变更执行经理：")
        self.Label_str_task = Label(self, text="变更内容：")
        self.Label_nam_yybg = Label(self, text="变更执行人：")
        self.Label_nam_bgfh = Label(self, text="变更复核人：")
        self.Label_datetime_bg = Label(self, text="变更时间：")

        self.var_num_moss = StringVar()
        self.var_num_bg = StringVar()
        self.var_nam_system = StringVar()
        self.var_nam_apply = StringVar()
        self.var_nam_dep = StringVar()
        self.var_nam_manager = StringVar()
        self.var_str_task = StringVar()
        self.var_nam_yybg = StringVar()
        self.var_nam_bgfh = StringVar()
        self.var_datetime_bg = StringVar()

        self.datetime_bg_vlu = "2000 01 01 18:00-20:00"
        self.var_datetime_bg.set(self.datetime_bg_vlu)

        self.Entry_num_moss = Entry(self)
        self.Entry_num_bg = Entry(self)
        self.Entry_nam_system = Entry(self)
        self.Entry_nam_apply = Entry(self)
        self.Entry_nam_dep = Entry(self)
        self.Entry_nam_manager = Entry(self)
        self.Entry_str_task = Entry(self)
        self.Entry_nam_yybg = Entry(self)
        self.Entry_nam_bgfh = Entry(self)
        self.Entry_datetime_bg = Entry(self,textvariable=self.var_datetime_bg)

        self.comfirm = Button(self, text="生成文档", command=self.confirm)

        self.Label_num_moss.grid(row=0,column=0,sticky='W',pady=3,padx=3)
        self.Label_num_bg.grid(row=1,column=0,sticky='W',pady=3,padx=3)
        self.Label_nam_system.grid(row=2,column=0,sticky='W',pady=3,padx=3)
        self.Label_nam_apply.grid(row=3,column=0,sticky='W',pady=3,padx=3)
        self.Label_nam_dep.grid(row=4,column=0,sticky='W',pady=3,padx=3)
        self.Label_nam_manager.grid(row=5,column=0,sticky='W',pady=3,padx=3)
        self.Label_str_task.grid(row=6,column=0,sticky='W',pady=3,padx=3)
        self.Label_nam_yybg.grid(row=7,column=0,sticky='W',pady=3,padx=3)
        self.Label_nam_bgfh.grid(row=8,column=0,sticky='W',pady=3,padx=3)
        self.Label_datetime_bg.grid(row=9,column=0,sticky='W',pady=3,padx=3)

        self.Entry_num_moss.grid(row=0,column=1,sticky='E',pady=3,padx=3)
        self.Entry_num_bg.grid(row=1,column=1,sticky='E',pady=3,padx=3)
        self.Entry_nam_system.grid(row=2,column=1,sticky='E',pady=3,padx=3)
        self.Entry_nam_apply.grid(row=3,column=1,sticky='E',pady=3,padx=3)
        self.Entry_nam_dep.grid(row=4,column=1,sticky='E',pady=3,padx=3)
        self.Entry_nam_manager.grid(row=5,column=1,sticky='E',pady=3,padx=3)
        self.Entry_str_task.grid(row=6,column=1,stick='E',pady=3,padx=3)
        self.Entry_nam_yybg.grid(row=7,column=1,sticky='E',pady=3,padx=3)
        self.Entry_nam_bgfh.grid(row=8,column=1,sticky='E',pady=3,padx=3)
        self.Entry_datetime_bg.grid(row=9,column=1,sticky='E',pady=3,padx=3)

        self.comfirm.grid(row=10,column=1,sticky='W',pady=3,padx=3)

    def confirm(self):
        num_moss = self.Entry_num_moss.get()
        num_bg = self.Entry_num_bg.get()
        nam_system = self.Entry_nam_system.get()
        nam_apply = self.Entry_nam_apply.get()
        nam_dep = self.Entry_nam_dep.get()
        nam_manager = self.Entry_nam_manager.get()
        str_task = self.Entry_str_task.get()
        nam_yybg = self.Entry_nam_yybg.get()
        nam_bgfh = self.Entry_nam_bgfh.get()
        datetime_bg = self.Entry_datetime_bg.get().split()
        year = datetime_bg[0]
        mouth = datetime_bg[1]
        date = datetime_bg[2]
        time = datetime_bg[3]
        if time == '':
            time = "18:00-20:00"
        str_datetime = year+mouth+date
        if os.path.exists(num_moss+'-bg'+num_bg) == True:
            shutil.rmtree(num_moss+'-bg'+num_bg)#
        w = win32com.client.Dispatch('Word.Application')
        w.Visible = 0
        w.DisplayAlerts = 0
        shutil.copytree('model', num_moss+'-bg'+num_bg)
        os.chdir(num_moss+'-bg'+num_bg)
        print(os.getcwd())
        filelist = glob.glob('*')
        for filename in filelist:
            filename_fmer = filename
            filename = filename.replace('{year}', year)
            filename = filename.replace('{num_moss}', num_moss)
            filename = filename.replace('{num_bg}', num_bg)
            os.replace(filename_fmer, filename)
            doc = w.Documents.Open(FileName = os.getcwd()+'\\'+filename)
            w.ActiveDocument.Sections[0].Headers[0].Range.Find.ClearFormatting()
            w.ActiveDocument.Sections[0].Headers[0].Range.Find.Replacement.ClearFormatting()
            w.ActiveDocument.Sections[0].Headers[0].Range.Find.Execute('{year}', False, False, False, False, False, True, 1, False, year, 2)
            w.ActiveDocument.Sections[0].Headers[0].Range.Find.Execute('{num_moss}', False, False, False, False, False, True, 1, False, num_moss, 2)
            w.Selection.Find.ClearFormatting()
            w.Selection.Find.Replacement.ClearFormatting()
            w.Selection.Find.Execute('{str_datetime}', False, False, False, False, False, True, 1, True, str_datetime, 2)
            w.Selection.Find.Execute('{year}', False, False, False, False, False, True, 1, True, year, 2)
            w.Selection.Find.Execute('{mouth}', False, False, False, False, False, True, 1, True, mouth, 2)
            w.Selection.Find.Execute('{date}', False, False, False, False, False, True, 1, True, date, 2)
            w.Selection.Find.Execute('{num_bg}', False, False, False, False, False, True, 1, True, num_bg, 2)
            w.Selection.Find.Execute('{nam_apply}', False, False, False, False, False, True, 1, True, nam_apply, 2)
            w.Selection.Find.Execute('{nam_manager}', False, False, False, False, False, True, 1, True, nam_manager, 2)
            w.Selection.Find.Execute('{nam_dep}', False, False, False, False, False, True, 1, True, nam_dep, 2)
            w.Selection.Find.Execute('{time}', False, False, False, False, False, True, 1, True, time, 2)
            w.Selection.Find.Execute('{nam_system}', False, False, False, False, False, True, 1, True, nam_system, 2)
            w.Selection.Find.Execute('{str_task}', False, False, False, False, False, True, 1, True, str_task, 2)
            w.Selection.Find.Execute('{nam_yybg}', False, False, False, False, False, True, 1, True, nam_yybg, 2)
            w.Selection.Find.Execute('{nam_bgfh}', False, False, False, False, False, True, 1, True, nam_bgfh, 2)
            w.Selection.Find.Execute('{h_start}', False, False, False, False, False, True, 1, True, time[0:2], 2)
            w.Selection.Find.Execute('{m_start}', False, False, False, False, False, True, 1, True, time[3:5], 2)
            w.Selection.Find.Execute('{h_end}', False, False, False, False, False, True, 1, True, time[6:8], 2)
            w.Selection.Find.Execute('{m_end}', False, False, False, False, False, True, 1, True, time[9:11], 2)
            doc.Close()
        w.Quit()
        messagebox.showinfo('生产文档', '文档生成完毕')

    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.createWidgets()
        self.pack()

app = Application()
app.master.title('XSD_YS')
app.mainloop()