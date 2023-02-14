# _*_ coding: utf-8 _*_
#edition = 1.0
#自动填写Excel表格 专用函数配合数据库使用
import os
import shutil
import time
import pyautogui
from win32con import WM_INPUTLANGCHANGEREQUEST
import win32gui
import win32api
import ctypes

def func_inquiry_inputer()->int: 
	'''
	用于查询当前的输入法类型
	@ return (int): 返回值为int类型 
	* 1 英文输入法 
	* 2 中文输入法 
	* -1 未查询到结果
	'''
	user32 = ctypes.WinDLL('user32', use_last_error=True)
	curr_window = user32.GetForegroundWindow()
	thread_id = user32.GetWindowThreadProcessId(curr_window, 0)
	klid = user32.GetKeyboardLayout(thread_id)
	lid = klid & (2**16 - 1)
	lid_hex = hex(lid)
	if lid_hex == '0x409':
		return 1
	elif lid_hex == '0x804':
		return 2
	else:
		return -1

def func_set_inputer(language:str)->int:  #设置输入法
	'''
	切换输入法，如果存在目标输入法
	@param language: 语言参数 目前支持 chinese 和 english
	@return (int) 
	* 0 切换成功 
	* -1 切换失败 
	* -2 不存在目标语言的输入法
	'''
	language_dict = {'chinese':0x0804, 'english':0x0409}
	# 0x0409为英文输入法的lid_hex的 中文一般为0x0804
	hwnd = win32gui.GetForegroundWindow()
	#title = win32gui.GetWindowText(hwnd)
	#im_list = win32api.GetKeyboardLayoutList()
	#im_list = list(map(hex, im_list))
	#print(im_list)
	if language not in language_dict: #如果目标关键字不在字典内 返回错误代码
		return -2
	result = win32api.SendMessage(hwnd, WM_INPUTLANGCHANGEREQUEST, 0, language_dict[language])
	if result == 0:
		return 0
	else:
		return -1

def func_mkdir(fullpath:str)->int:
	'''
	需要创建的的目标文件夹的完整路径
	@param fullpath: 目标文件夹的完整路径
	@return (int): 
	* 0 文件夹创建成功。
	* -1 目标文件夹已存在，未创建。
	'''
	path = fullpath
	# 去除首位空格
	path=path.strip()
	# 去除尾部 \ 符号
	path=path.rstrip("\\")

	# 判断路径是否存在
	# 存在     True
	# 不存在   False
	isExists=os.path.exists(path)
	# 判断结果
	if not isExists:
		# 如果不存在则创建目录
		# 创建目录操作函数
		os.makedirs(path) 
		return 0
	else:
		# 如果目录存在则不创建，并提示目录已存在
		return -1

def func_filecopy(origin_file:str, tar_folder:str, file_name:str)->int:
	'''
	复制文件函数
	@param origin_file: 源文件
	@param tar_folder: 目标文件夹
	@param name_file: 复制后的新文件的文件名
	@return (int): 
	* 0 复制成功
	* 1 未复制，文件已存在
	* -1 未知错误
	'''
	try:
		if os.path.exists(tar_folder+"\\"+file_name):
			return 1
		else:
			shutil.copy(origin_file, tar_folder)
			return 0
	except Exception as e:
		return -1

def func_fileremove(fullpath:str)->int:
	'''
	删除指定的文件
	@param fullpath 要删除的文件的完整路径
	@return (int): 
	* 0 执行成功
	* -1 执行失败
	'''
	try:
		os.remove(fullpath)
		return 0
	except Exception:
		return -1

def func_fileexist(fullpath:str)->int:
	'''
	判断制定文件是否存在
	@param fullpath: 目标文件的完整路径
	@return (int): 
	* 0 文件不存在
	* 1 文件存在
	* -1 执行失败
	'''
	try:
		ires = os.path.isfile(fullpath)
		if ires == True:
			return 1
		else:
			return 0
	except Exception:
		return -1



def func_openfile(fullpath:str)->int:
	'''
	打开文件
	@param fullpath: 文件的完整路径
	@return (int):
	* 0: 打开成功
	* 1: 文件已经存在
	'''
	if os.path.exists(fullpath):
		r_v = os.system('start '+fullpath)
		return 0
	else:
		return 1

def func_openfiles()->str:
	'''
	跳出一个资源管理器的文件夹选择界面，选择结束后返回路径
	@return (str): 返回选择的路径
	'''
	current_path = os.getcwd()
	return current_path


def func_simulate_close():
	'''
	模拟键盘alt+F4关闭文件
	@return (int) 1: 操作成功
	'''
	pyautogui.hotkey("alt", "F4")
	return 1

def func_simulate_saveexcel(filename:str)->int:
	'''
	模拟键盘保存excel文件，如果存在同名文件会操作错误
	@param filename: 保存的文件名
	@return (int): 0 保存成功
	''' 
	pyautogui.press("F12")
	pyautogui.hotkey("alt", "up")
	pyautogui.typewrite(filename)
	pyautogui.press("enter")
	return 0

def func_inputkey(key:str)->int:
	'''
	模拟键盘触发单个按键功能
	@param key: 需要触发的按键，字符串类型
	@return (int): 0 运行成功
	'''
	pyautogui.press(key)
	return 0

def func_local_input(position:str,value:str)->int:
	'''
	excel专用定位输入信息
	@param position: excel专用定位输入 需要填写的表格位置
	@param value: excel专用定位输入 需要填写的内容
	@return (int) 0: 运行正确
	'''
	pyautogui.press('F5')
	pyautogui.typewrite(position)
	pyautogui.press("enter")
	pyautogui.typewrite(value) 
	pyautogui.press("enter")
	return 0

def func_series_input(position:str, str_list:list)->int:
	'''
	excel专用连续输入信息
	@param position(str): 连续数据的输入起始位置
	@param str_list(list): 需要输入的字符串列表
	@return (int) 0:运行正确
	'''
	pyautogui.press('F5')
	pyautogui.typewrite(position)
	pyautogui.press("enter")
	length = len(str_list) #求出list的长度
	for i in str_list:
		pyautogui.typewrite(i)
		pyautogui.press('enter')
	return 0

if __name__ == '__main__':
	import pyqt5_plugins
	#ires = func_fileexist(r'D:\Code\Python\CamshaftRandomTable\excels\1_20230214RV145S凸轮轴成品出货全尺寸检测报告模板.xlsx')
	#ires = os.path.exists(r'D:\Code\Python\CamshaftRandomTable\excels\1_20230214RV145S凸轮轴成品出货全尺寸检测报告模板.xlsx')
	#print(ires)
	print(pyqt5_plugins.__file__)
	pass