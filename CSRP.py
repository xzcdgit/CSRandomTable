#code = utf-8
#edit = 1.0
#专用功能库
#凸轮轴测量出货数据随机生成器
import random

import time
import sys
from camt import Ui_MainWindow
from PyQt5.QtWidgets import QApplication, QComboBox, QMainWindow
class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__( self, parent=None ):  #初始函数
        super(MyMainWindow, self).__init__(parent) #继承父类函数
        self.setupUi(self) #继承窗口函数MyMainWindow.setupUi的所有变量并显示 以便修改界面函数显示
        self.ToolsSet()#设置界面槽函数

    def ToolsSet(self):     #配置文件参数设置
        self.pushButton.clicked.connect(self.MakeTable)  #创建表格
        self.pushButton_2.clicked.connect(self.OpenFile)  #打开文件夹

    def MakeTable(self):
        '''
        创建表格的槽函数
        '''
        class_csdata_creator = CamshaftTable() #创建一个类实例
        index = self.comboBox.currentIndex()
        if index == 0:
            path = class_csdata_creator.method_create_cstable('145')
            string = '已保存为:\n'+path+'\n! 如果电脑出现异常，请按一次 alt 键'
        elif index == 1:
            path =  class_csdata_creator.method_create_cstable('80')
            string = '已保存为:\n'+path+'\n! 如果电脑出现异常，请按一次 alt 键'
        else:
            string = '错误，操作失败'
        self.textBrowser.setText(string)

    def OpenFile(self):
        import os
        import AutoExcel
        sorcepath = AutoExcel.func_openfiles()
        os.startfile(sorcepath+'\\excels')

class CsAttribute():
    type_list = ['145','80']
    
    def func_re_attribute(self,tmptype:str, ires_type:str)->str:
        '''
        通过输入凸轮轴的型号和所需要返回的信息键名 就可以获得键值
        @param tmptype: 查询的凸轮轴型号 字符串类型
        @param ires_type: 查询的信息的键名 字符串类型
        @return (str): 返回的信息 键值 字符串类型
        '''
        if tmptype == self.type_list[0]:
            type = '145'
            mode_file_name = 'RV145S凸轮轴成品出货全尺寸检测报告模板.xlsx'
            exact_file_name = time.strftime(r'%Y%m%d')+'_145'
        elif tmptype == self.__type_list[1]:
            type = '80'
            mode_file_name = 'RV80S凸轮轴成品出货全尺寸检测报告模板.xlsx'
            exact_file_name = time.strftime(r'%Y%m%d')+'_80'
        else:
            raise ValueError('意外的错误')
        ires_dict = {'type':type, 'mode_file_name':mode_file_name, 'exact_file_name':exact_file_name}
        return ires_dict[ires_type]

    def __doc__()->str:
        return '查询信息的键名只能是 type: 键名  mode_file_name: 模板文件名  exact_file_name: 保存文件的文件名（不包括标号）'

class CamshaftTable():
    '''
    创建特定数据的excel表格类
    '''
    myexact = CsAttribute
    def __init__(self) -> None:
        pass

    def method_create_cstable(self,cs_type:str)->str:
        '''
        创建指定类型的凸轮轴的数据表格
        @param cs_type: 凸轮轴类型
        @return (str): 返回保存的完整路径
        '''
        import AutoExcel
        if cs_type not in self.myexact.type_list:
            return -1
        #打开源文件
        sorcepath = AutoExcel.func_openfiles()
        modepath = sorcepath+'\\mode\\'+self.myexact.func_re_attribute(self.myexact,cs_type,'mode_file_name')
        AutoExcel.func_openfile(modepath)
        time.sleep(2) #等待源文件打开完成

        #生成日期字符串
        date_str = time.strftime(r'%Y/%m/%d')
        #生成合理的随机数据 
        ires,data_list1 = self.__method_data_generator(cs_type) #数据列1
        ires,data_list2 = self.__method_data_generator(cs_type) #数据列2
        ires,data_list3 = self.__method_data_generator(cs_type) #数据列3
        AutoExcel.func_set_inputer('english')#切换英文输入
        #开始输入
        #填入日期
        AutoExcel.func_local_input('C3',date_str)
        #填入各列数据
        AutoExcel.func_series_input('F8',data_list1)
        AutoExcel.func_series_input('G8',data_list2)
        AutoExcel.func_series_input('H8',data_list3)

        #保存数据
        i=1
        while i<99:
            filename = 'excels\\'+str(i)+'_'+time.strftime(r'%Y%m%d')+self.myexact.func_re_attribute(self.myexact,cs_type,'exact_file_name')
            path = sorcepath+'\\'+ filename+'.xlsx'
            ires =  AutoExcel.func_fileexist(path)
            print(path)
            if ires == 0:
                break
            i += 1
        AutoExcel.func_simulate_saveexcel(filename)
        time.sleep(0.7)
        AutoExcel.func_simulate_close()
        AutoExcel.func_inputkey('N')
        return path

    def __method_data_generator(self,cs_type:str)->tuple[int,list[int]]:
        '''
        创建凸轮轴表格信息
        @param cs_type: 选择型号数据 字符串类型
        @return (list): 返回一个元组
        * (tuple)[0]: int类型的数据表明该函数执行是否成功 -1表示生成失败 其他非负整数表示匹配类型位于字典中的序号
        * (tuple)[1]: list[int]类型的数据为生成的制定数据list
        '''
        measure_data = []
        key = -1
        if cs_type == self.myexact.type_list[0]:
            measure_data.append(str(random.randint(6870,6890)*0.01))
            measure_data.append(str(random.randint(9365,9378)*0.01))
            #
            if random.randint(0,100)>98:
                measure_data.append(str(11.98))
            else:
                measure_data.append(str(11.97))

            measure_data.append(str(random.randint(7162,7172)*0.01))
            measure_data.append(str(random.randint(701,710)*0.01))
            measure_data.append(str(random.randint(845,854)*0.01))
            measure_data.append(str(random.randint(1897,1908)*0.01))
            measure_data.append(str(random.randint(1781,1792)*0.01))
            measure_data.append(str(random.randint(1748,1757)*0.01))

            measure_data.append(str(min(random.randint(990,998)*0.01,random.randint(990,998)*0.01)))
            measure_data.append(str(random.randint(2508,2519)*0.01))

            measure_data.append(str( min(random.randint(990,997)*0.01,min(random.randint(990,997)*0.01,random.randint(990,997)*0.01)) ))
            measure_data.append(str(random.randint(1085,1112)*0.01))
            measure_data.append(str(random.randint(2436,2443)*0.01))
            measure_data.append(str(random.randint(2438,2447)*0.01))
            measure_data.append(str(random.randint(797,814)*0.01))
            measure_data.append(str(random.randint(10460,10540)*0.01))

            #进排气基圆跳动
            k = random.randint(0,100)
            if k>70:
                measure_data.append(str(0.07))
            elif k>50:
                measure_data.append(str(0.06))
            else:
                measure_data.append(str(random.randint(4,7)*0.01))

            k = random.randint(0,100)
            if k>70:
                measure_data.append(str(0.07))
            elif k>40:
                measure_data.append(str(0.06))
            else:
                measure_data.append(str(random.randint(4,6)*0.01))

            measure_data.append(str(random.randint(3955,4044)*0.01))
            #进排气升程跳动

            k = random.randint(0,100)
            if k>60:
                measure_data.append(str(0.07))
            elif k>40:
                measure_data.append(str(0.06))
            else:
                measure_data.append(str(random.randint(5,9)*0.01))


            k = random.randint(0,100)
            if k>60:
                measure_data.append(str(0.07))
            elif k>30:
                measure_data.append(str(0.06))
            else:
                measure_data.append(str(random.randint(5,9)*0.01))

            k = random.randint(0,100)
            if k>3:
                measure_data.append(str(max(random.randint(60,64)*10,random.randint(60,64)*10)))
            else:
                measure_data.append(str(650))

            measure_data.append(str(random.randint(110,120)*0.01))
            key = 0
        elif cs_type == self.myexact.type_list[1]:
            measure_data.append(str(random.randint(6870,6890)*0.01))
            measure_data.append(str(random.randint(9365,9378)*0.01))
            #
            if random.randint(0,100)>98:
                measure_data.append(str(11.98))
            else:
                measure_data.append(str(11.97))

            measure_data.append(str(random.randint(7162,7172)*0.01))
            measure_data.append(str(random.randint(701,710)*0.01))
            measure_data.append(str(random.randint(845,854)*0.01))
            measure_data.append(str(random.randint(1897,1908)*0.01))
            measure_data.append(str(random.randint(1781,1792)*0.01))
            measure_data.append(str(random.randint(1748,1757)*0.01))

            measure_data.append(str(min(random.randint(990,998)*0.01,random.randint(990,998)*0.01)))
            measure_data.append(str(random.randint(2508,2519)*0.01))

            measure_data.append(str( min(random.randint(990,997)*0.01,min(random.randint(990,997)*0.01,random.randint(990,997)*0.01)) ))
            measure_data.append(str(random.randint(1085,1112)*0.01))
            measure_data.append(str(random.randint(2436,2443)*0.01))
            measure_data.append(str(random.randint(2438,2447)*0.01))
            measure_data.append(str(random.randint(797,814)*0.01))     #切边距
            measure_data.append(str(random.randint(10460,10540)*0.01)) #进排气夹角

            #进排气基圆跳动
            k = random.randint(0,100)
            if k>90:
                measure_data.append(str(0.07))
            elif k>50:
                measure_data.append(str(0.06))
            else:
                measure_data.append(str(random.randint(4,7)*0.01))

            k = random.randint(0,100)
            if k>85:
                measure_data.append(str(0.07))
            elif k>40:
                measure_data.append(str(0.06))
            else:
                measure_data.append(str(random.randint(3,6)*0.01))

            measure_data.append(str(random.randint(1896,2084)*0.01))
            #进排气升程跳动

            k = random.randint(0,100)
            if k>60:
                measure_data.append(str(0.07))
            elif k>40:
                measure_data.append(str(0.06))
            else:
                measure_data.append(str(random.randint(5,9)*0.01))

            k = random.randint(0,100)
            if k>60:
                measure_data.append(str(0.07))
            elif k>30:
                measure_data.append(str(0.06))
            else:
                measure_data.append(str(random.randint(5,9)*0.01))

            k = random.randint(0,100)
            if k>3:
                measure_data.append(str(max(random.randint(60,64)*10,random.randint(60,64)*10)))
            else:
                measure_data.append(str(650))

            measure_data.append(str(random.randint(110,120)*0.01))
            key = 1
        return key, measure_data

if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MyMainWindow() #初始化窗口
    win.show()
    sys.exit(app.exec_())
