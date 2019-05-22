# EasyReplace
一款批量替换word 内容的python 小程序

[![Travis](https://img.shields.io/travis/rust-lang/rust.svg)](https://github.com/songluyi/EasyReplace)
[![PyPI](https://img.shields.io/pypi/wheel/Django.svg)](https://github.com/songluyi/EasyReplace)
[![Hex.pm](https://img.shields.io/hexpm/l/plug.svg)](https://github.com/songluyi/EasyReplace)
## Introduction
A small tool to replace something in Office Word docx document with reading replace 
information from Office Excel.

从excel 中读取原内容和要替换的内容，然后再制定data 目录下寻找word 文档进行批量替换。

这里简单说明一下对于原生word 还是尽量使用win32com 库对office 的事务进行操作，特别是
所谓的python-wordx 操作和替换功能类似于玩具，对于复杂格式和工业级应用上难以满足。

目前该版本有一个highlight 函数无法做高亮处理。
