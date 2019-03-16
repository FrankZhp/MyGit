# -*- coding:utf-8 -*-
import os

print('***获取当前目录***')
print(os.getcwd())


print('\n')

print('***获取上级目录***')
print(os.path.abspath(os.path.dirname(os.getcwd())))
print(os.path.abspath(os.path.join(os.getcwd(), "..")))

print('\n')

print('***获取上上级目录***')
print(os.path.abspath(os.path.join(os.getcwd(), "../..")))
