import xlwt
import xlrd

class Excel:
    def __init__(self, name):
        self.workbook = xlwt.Workbook(encoding='utf-8')
        self.file_column = {}
        self.file_sheet = {}
        self.name = name
        self.style = Excel.DefaultStyle()

    # 获取默认样式
    @staticmethod
    def DefaultStyle():
        borders = Excel.allborder()
        style = xlwt.XFStyle()
        style.borders = borders
        # 设置单元格对齐方式
        alignment = xlwt.Alignment()
        # 水平位置：0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
        alignment.horz = 0x01
        style.alignment = alignment
        return style

    # 设置默认样式
    def setDefaultStyle(self, style):
        self.style = style

    # 获取当前默认样式
    def getDafaultStyle(self):
        return self.style

    # 全边框
    @staticmethod
    def allborder():
        borders = xlwt.Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        return borders

    # 通过颜色代码获取样式
    @staticmethod
    def getColorStyleByCode(code):
        style = Excel.DefaultStyle()
        style.borders = Excel.allborder()
        style.pattern.pattern = xlwt.Pattern.SOLID_PATTERN
        style.pattern.pattern_fore_colour = code
        return style

    # 浅橙色
    @staticmethod
    def important_color():
        return Excel.getColorStyleByCode(51)

    # 深橙色
    @staticmethod
    def very_important_color():
        return Excel.getColorStyleByCode(52)

    # 居中
    @staticmethod
    def center():
        '''
        设置居中
        :return:
        '''
        style = Excel.DefaultStyle()
        style.alignment.horz = 0x02
        return style

    #
    def __create_xls_obj__(self, sheet_name):
        self.file_column[sheet_name] = 0
        self.file_sheet[sheet_name] = self.workbook.add_sheet(sheet_name)

    def write_column(self, sheet_name, column, style=None):
        if style == None:
            style = self.style
        if sheet_name not in self.file_sheet:
            self.__create_xls_obj__(sheet_name)
        worksheet = self.file_sheet[sheet_name]
        row = 0
        for i in column:
            worksheet.write(self.file_column[sheet_name], row, i, style=style)
            row += 1
        self.file_column[sheet_name] += 1

    def save(self):
        self.workbook.save(self.name)

    def writeDB(self, data, header, sheetName, style=None):
        if style == None:
            style = self.style
        self.write_column(sheetName, header, style=style)
        for i in data:
            row = []
            for key in header:
                row.append(i[key])
            self.write_column(sheetName, row, style=style)

    # 得到当前时间并写入
    def writeTime(self, sheetName, row_range):
        '''
        写入当前时间
        :param sheetName: sheet名称
        :param row_range: 单元格长度
        :return: None
        '''
        import time
        time_str = time.strftime("数据统计于%Y-%m-%d %H:%M:%S", time.localtime())
        self.write_center_after_merge(sheetName, time_str, row_range)
        self.save()

    def write_center_after_merge(self, sheet_name, text, row_range):
        if sheet_name not in self.file_sheet:
            self.__create_xls_obj__(sheet_name)
        worksheet = self.file_sheet[sheet_name]
        col = self.file_column[sheet_name]
        worksheet.write_merge(col, col, 0, row_range - 1, text, Excel.center())
        self.file_column[sheet_name] += 1

    # 获取信息 返回值为xls第一行和各行数据的字典组合成的列表
    @staticmethod
    def getAllDataAsDict(name, n=0, end_colx=None, nrows=None,sheet_name=None):
        # 检测文件后缀名
        temp_flag=False
        if name.split('.')[-1] == 'xlsx':
            Excel.xlsx_to_xls(name, name.split('.')[0] + '.xls',False)
            temp_flag=True
        name = name.split('.')[0] + '.xls'
        book = xlrd.open_workbook(name)
        if sheet_name == None:
            table = book.sheet_by_index(0)
        else:
            table = book.sheet_by_name(sheet_name)
        db = []
        header = table.row_values(n, start_colx=0, end_colx=end_colx)
        if nrows == None:
            nrows = table.nrows
        for i in range(2, nrows):
            db.append(dict(zip(header, table.row_values(i, start_colx=0, end_colx=None))))
        if temp_flag:
            import os
            os.remove(name)
        return db

    @staticmethod
    def find_all_on_key_by_name(db,key,name):
        ans=[]
        for i in db:
            if i[key]==name:
                ans.append(i)
        return ans

    @staticmethod
    def xlsx_to_xls(fname, export_name, delete_flag=True):
        """
        将xlsx文件转化为xls文件
        :param fname: 传入待转换的文件路径(可传绝对路径，也可传入相对路径，都可以)
        :param export_name: 传入转换后到哪个目录下的路径(可传绝对路径，也可传入相对路径，都可以)
        :param delete_flag: 转换成功后，是否删除原来的xlsx的文件,默认删除 布尔类型
        :return:    无返回值
        """
        import win32com.client as win32 # pywin32
        import os
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False
        absolute_path = os.path.join(os.path.dirname(os.path.abspath(fname)), os.path.basename(fname))
        save_path = os.path.join(os.path.dirname(os.path.abspath(export_name)), os.path.basename(export_name))
        wb = excel.Workbooks.Open(absolute_path)
        wb.SaveAs(save_path, FileFormat=56)  # FileFormat = 51 is for .xlsx extension
        wb.Close()  # FileFormat = 56 is for .xls extension
        excel.Application.Quit()
        if delete_flag:
            os.remove(absolute_path)

if __name__ == '__main__':
    pass