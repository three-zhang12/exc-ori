from PySide2.QtWidgets import QApplication, QMessageBox, QTableWidget, QFileDialog, QTableWidgetItem
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import QFile, QUrl
from PySide2.QtGui import QIcon
from PySide2.QtWebEngineWidgets import QWebEngineView

import pandas as pd
import os

from pyecharts.charts import Line
from pyecharts import options as opts

from scipy import signal
from threading import Thread

class MainWin:
    def __init__(self):
        qfile = QFile('main.ui')
        qfile.open(qfile.ReadOnly)
        qfile.close()

        self.ui = QUiLoader().load(qfile)
        self.ui.setWindowTitle('excigin')
        # self.ui.openfilebutton.clicked.connect(self.openfilebutton)
        # self.ui.savefilebutton.clicked.connect(self.savefilebutton)
        self.ui.actionopen.triggered.connect(self.openfilebutton)
        self.ui.actionsave.triggered.connect(self.savefilebutton)
        self.ui.operation_button.clicked.connect(self.startbutton)  # 行列基本运算
        self.ui.eval_button.clicked.connect(self.evalbutton)  # 行列最大值等
        self.ui.pic_button.clicked.connect(self.picbutton)
        self.ui.filter_button.clicked.connect(self.filterbutton)
        self.ui.section_button.clicked.connect(self.sectionbutton)

        self.file = None
        self.isnull = None
        self.pic_show = QWebEngineView()  # 图片预览框
        self.ui.pic_show_layout.addWidget(self.pic_show)

    def openfilebutton(self):

        filepath = QFileDialog.getOpenFileName(self.ui, "请选择需要打开的文件",
                                           r'c:/', '文件 (*.xlsx);;文件 (*.xls);;文件 (*.csv);;文件 (*.txt)')

        if filepath[0]:
            self.file = pd.read_excel(filepath[0], header=None)
            # self.file.index = self.file.index * 1000
            self.isnull = self.file.isnull().any()
            # 表格显示添加数据
            showline = int(self.ui.data_show_cbox.currentText())
            for x in range(self.file.shape[1]):
                self.ui.data_show_table.insertColumn(x)
                for y in range(showline):
                    if self.ui.data_show_table.rowCount() < showline:
                        self.ui.data_show_table.insertRow(y)
                    thread_open = Thread(target=self.thread_open,
                                         args=(x, y))
                    thread_open.start()
            QMessageBox.about(self.ui, '文件导入成功', '请进行下一步操作')

    def savefilebutton(self):
        filepath = QFileDialog.getSaveFileName(self.ui, "请选择需要打开的文件",
                                               r'c:/', '文件 (*.xls);;文件 (*.csv);;文件 (*.txt)')
        if filepath[0]:
            self.file.to_excel(filepath[0], header=None, index=False)
            QMessageBox.about(self.ui, '文件保存成功', '请进行下一步操作')

    def startbutton(self):
        def trans_col(x):
            return eval(dowhat.replace('x', str(x[int(col) - 1])))

        def trans_cols(x):
            return eval(dowhat.replace('x', str(x[int(each - 1)])))

        def trans_row(x):
            return eval(dowhat.replace('x', str(x.iloc[int(col) - 1])))

        def trans_rows(x):
            return eval(dowhat.replace('x', str(x.iloc[int(each) - 1])))

        def trans_area(x):
            return eval(dowhat.replace('x', str(x[each-1])))

        order = self.ui.operation_cbox.currentText()
        col = self.ui.operation_index_col.text()
        flag = 1
        # 检查行列输入是否正确
        try:
            if order == '整体操作':
                self.ui.operation_index_col.clear()
            else:
                if col.isdigit():
                    if int(col) < 0 or (order == '单列操作' and int(col) >= self.file.shape[1]):
                        flag = 0
                        QMessageBox.critical(self.ui, '错误', '请输入正确行列数！')
                    elif int(col) < 0 or (order == '单行操作' and int(col) >= self.file.shape[0]):
                        flag = 0
                        QMessageBox.critical(self.ui, '错误', '请输入正确行列数！')
                elif ',' in col and ';' not in col:
                    a, b = col.split(',')
                    if int(a) < 0 or int(b) > self.file.shape[1]:
                        flag = 0
                        QMessageBox.critical(self.ui, '错误', '请输入正确行列数！')
                elif ';' in col and ',' in col:
                    col_row, col_col = col.split(';')
                    if ',' in col_col:
                        a, b = col_col.split(',')
                        if int(a) < 0 or int(b) >= self.file.shape[1]:
                            flag = 0
                            QMessageBox.critical(self.ui, '错误', '请输入正确行列数！')
                    else:
                        if int(col_col) < 0 or int(col_col) >= self.file.shape[1]:
                            flag = 0
                            QMessageBox.critical(self.ui, '错误', '请输入正确行列数！')
                    if ',' in col_row:
                        c, d = col_row.split(',')
                        if int(c) < 0 or int(d) >= self.file.shape[0]:
                            flag = 0
                            QMessageBox.critical(self.ui, '错误', '请输入正确行列数！')
                    else:
                        if int(col_row) < 0 or int(col_row) >= self.file.shape[0]:
                            flag = 0
                            QMessageBox.critical(self.ui, '错误', '请输入正确行列数！')
                else:
                    flag = 0
                    QMessageBox.critical(self.ui, '错误', '请输入正确的格式！')
        except:
            flag = 0
            QMessageBox.critical(self.ui, '错误', '请输入正确的格式！')

        dowhat = self.ui.operation_do.text()
        # 检查表达式是否正确
        try:
            eval(dowhat.replace('x', '1'))
        except:
            flag = 0
            QMessageBox.critical(self.ui, '错误', '请输入正确的表达式！')

        if flag:
            if order == '单列操作':
                try:
                    self.file[int(col) - 1] = self.file.apply(trans_col, axis=1)
                except:
                    QMessageBox.critical(self.ui, '错误', '数据列存在缺失，请使用区域操作！')

            elif order == '单行操作':
                try:
                    self.file.iloc[int(col) - 1] = self.file.apply(trans_row, axis=0)
                except:
                    QMessageBox.critical(self.ui, '错误', '数据列存在缺失，请使用区域操作！')

            elif order == '多列操作':
                if ',' in col:
                    for each in range(int(col.split(',')[0]), int(col.split(',')[1]) + 1):
                        try:
                            self.file[each - 1] = self.file.apply(trans_cols, axis=1)  # 后期改用多线程处理
                        except:
                            pass
                else:
                    QMessageBox.critical(self.ui, '错误', '请输入正确行列数！')

            elif order == '多行操作':
                if ',' in col:
                    for each in range(int(col.split(',')[0]), int(col.split(',')[1]) + 1):
                        try:
                            self.file.iloc[int(each) - 1] = self.file.apply(trans_rows, axis=0)  # 后期改用多线程处理
                        except:
                            pass
                else:
                    QMessageBox.critical(self.ui, '错误', '请输入正确行列数！')

            elif order == '区域操作':
                if ';' in col:
                    if ',' in col:
                        if ',' in col.split(';')[1]:
                            if ',' in col.split(';')[0]:
                                for each in range(int(col.split(';')[1].split(',')[0]), int(col.split(';')[1].split(',')[1]) + 1):
                                    self.file.iloc[int(col.split(';')[0].split(',')[0]) - 1:int(col.split(';')[0].split(',')[1]), each - 1] = self.file.apply(trans_area, axis=1)
                            else:
                                for each in range(int(col.split(';')[1].split(',')[0]), int(col.split(';')[1].split(',')[1]) + 1):
                                    print(int(col.split(';')[1].split(',')[0]), int(col.split(';')[1].split(',')[0]) + 1)
                                    self.file.iloc[int(col.split(';')[0]) - 1, each - 1] = eval(
                            dowhat.replace('x', str(self.file.iloc[int(col.split(';')[0]) - 1, each - 1])))
                        elif ',' in col.split(';')[0]:
                            each = int(col.split(';')[1])
                            self.file.iloc[int(col.split(';')[0].split(',')[0]) - 1:int(col.split(';')[0].split(',')[1]),
                            int(col.split(';')[1])-1] = self.file.apply(trans_area, axis=1)
                    else:
                        self.file.iloc[int(col) - 1, int(col) - 1] = eval(
                            dowhat.replace('x', str(self.file.iloc[int(col) - 1, int(col) - 1])))

                else:
                    QMessageBox.critical(self.ui, '错误', '请输入正确行列数！')

            elif order == '整体操作':
                for each in range(1, self.file.shape[1] + 1):
                    if not self.isnull[each - 1]:
                        self.file[each - 1] = self.file.apply(trans_cols, axis=1)

            QMessageBox.about(self.ui, '操作成功', '请进行下一步操作')

    def evalbutton(self):
        flag = 1
        choice = self.ui.eval_choice_cbox.currentText()
        row_col = self.ui.eval_line.text()
        # 检查行列输入是否正确
        try:
            int(row_col)
            if choice == '单列' and (int(row_col) >= self.file.shape[1] or int(row_col) < 1):
                flag = 0
                QMessageBox.critical(self.ui, '错误', '请输入正确行列数！')
            elif choice == '单行' and (int(row_col) >= self.file.shape[0] or int(row_col) < 1):
                flag = 0
                QMessageBox.critical(self.ui, '错误', '请输入正确行列数！')
        except:
            flag = 0
            QMessageBox.critical(self.ui, '错误', '请输入正确行列数！')
        func = self.ui.eval_cbox.currentText()
        self.ui.eval_show.clear()

        if flag:
            if func == '和':
                self.ui.eval_show.append(str(self.file[int(row_col)-1].sum())) if choice == '单列' else self.ui.eval_show.append(str(self.file.iloc[int(row_col) - 1].sum()))
            elif func == '平均值':
                self.ui.eval_show.append(str(self.file[int(row_col)-1].mean())) if choice == '单列' else self.ui.eval_show.append(str(self.file.iloc[int(row_col) - 1].mean()))
            elif func == '最大值':
                self.ui.eval_show.append(str(self.file[int(row_col)-1].max())) if choice == '单列' else self.ui.eval_show.append(str(self.file.iloc[int(row_col) - 1].max()))
            elif func == '最小值':
                self.ui.eval_show.append(str(self.file[int(row_col)-1].min())) if choice == '单列' else self.ui.eval_show.append(str(self.file.iloc[int(row_col) - 1].min()))
            elif func == '标准误差':
                self.ui.eval_show.append(str(self.file[int(row_col)-1].std())) if choice == '单列' else self.ui.eval_show.append(str(self.file.iloc[int(row_col) - 1].std()))
            elif func == '众数':
                self.ui.eval_show.append(str(self.file[int(row_col)-1].mode())) if choice == '单列' else self.ui.eval_show.append(str(self.file.iloc[int(row_col) - 1].mode()))
            elif func == '无偏误差':
                self.ui.eval_show.append(str(self.file[int(row_col)-1].var())) if choice == '单列' else self.ui.eval_show.append(str(self.file.iloc[int(row_col) - 1].var()))
            elif func == '偏差':
                self.ui.eval_show.append(str(self.file[int(row_col)-1].mad())) if choice == '单列' else self.ui.eval_show.append(str(self.file.iloc[int(row_col) - 1].mad()))
            elif func == '中位数':
                self.ui.eval_show.append(str(self.file[int(row_col)-1].median())) if choice == '单列' else self.ui.eval_show.append(str(self.file.iloc[int(row_col) - 1].median()))
            elif func == '连乘积':
                self.ui.eval_show.append(str(self.file[int(row_col)-1].prod())) if choice == '单列' else self.ui.eval_show.append(str(self.file.iloc[int(row_col) - 1].prod()))

            QMessageBox.about(self.ui, '操作成功', '请进行下一步操作')

    def picbutton(self):
        def pic(col, path, file):
            col_1 = col + 1
            chart = Line()
            chart.add_xaxis(file.index.to_list())
            chart.add_yaxis(f'第{col_1}列', file[col].to_list(), is_symbol_show=False)
            chart.set_global_opts(
                title_opts=opts.TitleOpts(title='数据处理'),
                tooltip_opts=opts.TooltipOpts(trigger='axis', axis_pointer_type='cross'),  # 交互性十字标
                toolbox_opts=opts.ToolboxOpts(),
                legend_opts=opts.LegendOpts(),

                xaxis_opts=opts.AxisOpts(name='时间'),  # 设置x轴名字属性
                yaxis_opts=opts.AxisOpts(name='电荷')
            )
            chart.render(path+'-'+str(col+1)+'.html')
            self.pic_show.load(QUrl.fromLocalFile(os.path.realpath(path+'-'+str(col+1)+'.html')))

        col_cols = self.ui.pic_cbox.currentText()
        col = self.ui.pic_line.text()
        show = self.ui.pic_choose_cbox.currentText()
        flag = 1
        # 检查行列数
        try:
            if col_cols == '单列' and (int(col) >= self.file.shape[1] or int(col) < 1):
                flag = 0
                QMessageBox.critical(self.ui, '错误', '请输入正确行列数！')
            elif show or '全部列':
                flag = 1
        except:
            flag = 0
            QMessageBox.critical(self.ui, '错误', '请输入正确行列数！')

        if show:
            self.ui.pic_line.clear()
            self.pic_show.load(QUrl.fromLocalFile(os.path.realpath('pic/pic-' + show[1: -1] + '.html')))

        else:
            # 创建图像文件保存位置，每次生成前删除上一次保存
            path = 'pic'
            if not os.path.exists(path):
                os.mkdir(path)
            else:
                file_list = os.listdir(path)
                for each in file_list:
                    os.remove(path + '/' + each)
            os.chdir(path)
            if col_cols == '单列':
                if not self.isnull[int(col) - 1]:
                    pic(int(col) - 1, path, self.file)
                else:
                    QMessageBox.critical(self.ui, '错误', '请输入正确行列数！')
            elif col_cols == '全部列':
                self.ui.pic_line.clear()
                for each in range(self.file.shape[1]):
                    if not self.isnull[each]:
                        pic(each, path, self.file)
                        self.ui.pic_choose_cbox.addItem(f'第{each + 1}列')
            os.chdir('..')

    def filterbutton(self):
        choice = self.ui.filter_cbox.currentText()
        try:
            hz = int(self.ui.filter_line.text())
            hz_1 = int(self.ui.filter_line_1.text())
            wn = 2*hz/hz_1
            b, a = signal.butter(8, wn, choice)

            for each in range(1, self.file.shape[1]):
                if not self.isnull[each]:
                    thread_filter = Thread(target=self.thread_filter,
                                           args=(b, a, each))
                    thread_filter.start()
            print(self.file.head())
            QMessageBox.about(self.ui, '滤波完成', '请进行下一步操作')
        except:
            QMessageBox.critical(self.ui, '错误', '请输入正确数值！')

    def sectionbutton(self):
        section = self.ui.section_line.text()
        flag = 1
        try:
            a, b = section.split()
            if int(a) < 0 or int(b) >= self.file.shape[0]:
                flag = 0
                QMessageBox.critical(self.ui, '错误', '请输入可用区间！')
        except:
            flag = 0
            QMessageBox.critical(self.ui, '错误', '请输入正确格式！')
        if flag:
            show = self.ui.pic_choose_cbox.currentText()
            col = self.ui.pic_line.text()
            if show:
                result = abs(self.file.loc[int(a):int(b), int(show[1:-1]) - 1 ].max() - self.file.loc[int(a):int(b),
                             int(show[1:-1]) - 1].min())
            else:
                result = abs(self.file.loc[int(a):int(b), int(col) - 1].max() - self.file.loc[int(a):int(b),
                             int(col) - 1].min())
            self.ui.section_show.clear()
            self.ui.section_show.append(str(result))

    def thread_filter(self, b, a, each):
        self.file[each] = signal.filtfilt(b, a, self.file[each])

    def thread_open(self, x, y):
        item = QTableWidgetItem()
        item.setText(str(self.file.loc[y, x]))
        self.ui.data_show_table.setItem(y, x, item)


app = QApplication([])
app.setWindowIcon(QIcon('ico.png'))
mainwin = MainWin()
mainwin.ui.show()
app.exec_()