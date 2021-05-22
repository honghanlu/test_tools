import xlrd
import os
import time

class Transform():
    """初始化数据"""
    def __init__(self):
        self.content = ""
        self.content_list = []

    """获取Excel表格测试数据"""
    def read_excel(self,file,sheet):
        root_path = os.path.split(os.path.split(os.path.realpath(__file__))[0])[0]
        filepath = os.path.join(root_path,file)
        print(filepath)
        xlbook = xlrd.open_workbook(filepath)
        table = xlbook.sheet_by_name(sheet)
        nrows = table.nrows
        ncols = table.ncols
        # 定义一个空数组，用来存放每行的数据
        a = []
        # 获取第一行数据作为字典的的key
        key = table.row_values(0)
        if nrows <= 1:
            print("{}未读取到数据".format(sheet))
        else:
            for i in range(1, nrows):
                # 定义一个空字典，用来存放每行的数据
                b = {}
                values = table.row_values(i)
                for j in range(ncols):
                    b[key[j]] = values[j]
                a.append(b)
            return a

    """校验模板是否准确"""
    def yanzheng(self,filepath, sheet):
        xlbook = xlrd.open_workbook(filepath)
        table = xlbook.sheet_by_name(sheet)
        bt = table.row_values(0)
        mb = ['testsuite_name', 'testcase', 'summary', 'preconditions', 'execution_type', 'importance', 'action','expect_result']
        if [False for a in mb if a not in bt]:
            print("模板错误，请更换导入模板")
            return False
        else:
            return True

    """数据写入xml文件中"""
    def write_to_file(self, ExcelFileName, SheetName):
        xmlFileName = ExcelFileName + '_' + SheetName + '.xml'
        cp = open(xmlFileName, "w",encoding='utf-8')
        cp.write(self.content)
        cp.close()

    def content_to_xml(self, key, value=None):
        if key == 'execution_type' or key == 'importance' or key == 'preconditions' or key == 'summary':
            return "<" + str(key) + "><![CDATA[" + str(value) + "]]></" + str(key) + ">"
        elif key == 'actions' or key == 'expectedresults':
            return "<" + str(key) + "><![CDATA[<p> " + str(value) + "</p> ]]></" + str(key) + ">"
        elif key == 'testcase':
            return '<testcase name="' + str(value) + '">'
        elif key == 'testsuite_name':
            return '<testsuite name="' + str(value) + '">'
        else:
            return '##########'

    def conversion(self, ExcelFileName, SheetName,filepath):
        if self.yanzheng(filepath,SheetName) == True and ExcelFileName != '':
            print("开始转化测试数据")
            testcase_list = self.read_excel(filepath,SheetName)
            i = 0
            for testcase in testcase_list:
                self.content += self.content_to_xml("actions", testcase['action'])
                self.content += self.content_to_xml("expectedresults", testcase['expect_result'])
                self.content = "<steps>" + self.content + "</steps>"
                self.content = self.content_to_xml("preconditions", testcase['preconditions']) + self.content
                self.content = self.content_to_xml("execution_type", testcase['execution_type']) + self.content
                self.content = self.content_to_xml("summary", testcase['summary']) + self.content
                self.content = self.content_to_xml("testcase", testcase['testcase']) + self.content +"</testcase>"
                i= i+1
                print("新增第{}条测试用例，测试标题：{}".format(i,testcase['testcase']))
                self.content = self.content_to_xml("testsuite_name", testcase['testsuite_name']) +self.content+"</testsuite>"
                self.content_list.append(self.content)
                self.content = ""
            self.content = "".join(self.content_list)
            self.content = '<?xml version="1.0" encoding="UTF-8"?>' +'<testsuite name ="{}">'.format(ExcelFileName)+ self.content+'</testsuite>'
            self.write_to_file(ExcelFileName, SheetName,)
            print("数据转化成功")
        else:
            print("转换后xml文件名称必填")

if __name__ == '__main__':
    # file = r'D:\study\study_data\test_link\dist\收款撤回测试用例.xlsx'
    # sheet = '服务合同'
    # ExcelFileName = ''
    try:
        file = str(input("测试用例路径（请填写绝对路径）:"))
        sheet = str(input("sheet名称:"))
        ExcelFileName =str(input("请输入转换后xml文件名称（不需文件类型）:"))
        Transform().conversion(ExcelFileName,sheet,file)
    except:
        print("路径不存在或者找不到sheet页")
    # Transform().conversion(ExcelFileName, sheet, file)
    os.system('pause')
