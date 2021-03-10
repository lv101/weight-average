import os
import time
import xlwt
import pandas
import warnings


class get_weight_score():
    def __init__(self):
        warnings.filterwarnings("ignore")
        self.value = []
        self.file_list = []
        self.end = 1
        self.flag = 1
        self.mode = 0
        self.count = -1
        self.debug_mode = 0
        self.fina_file = "综测统计表.xls"
        self.dir_name = "成绩单"
        self.path = f"{os.getcwd()}\\{self.dir_name}\\"
        self.headline = ""

    def self_round(self, value, decimal=0):
        multi = pow(10, decimal)
        return round(value * multi) / multi

    def self_init(self):
        self.mode = input("\n  请选择你所在学校加权平均分计算方式[1-2]\n"
                          "\t1.只包含必修课程\n"
                          "\t2.包含必修课程+选修课程\n"
                          "> ")
        if self.mode == '1':
            self.headline = ["序号", "学号", "姓名", "全部课程", "必修课程", "全部成绩", "必修成绩", "已修学分", "必修学分", "加权平均分", "平均绩点"]
            return False
        elif self.mode == '2':
            self.headline = ["序号", "学号", "姓名", "全部课程", "必修课程", "选修课程", "全部成绩", "必修成绩", "选修成绩", "已修学分", "必修学分", "选修学分", "加权平均分", "平均绩点"]
            return False
        else:
            print("Input Error.")
            self.end = 0
            return True

    def self_exit(self):
        print()
        if self.end:
            for i in range(6, 0, -1):
                # print(f"\r  程序将在 {i}s 后自动退出...", end='')
                print(f"\r  The program will exit automatically after {i}s ...", end='')
                time.sleep(1)
        else:
            input("  Enter any key to continue...")

    def get_dir_file(self):
        '''
        获取指定路径下的所有文件名,若为Excel文件,则返回文件名
        :return: None
        '''
        try:
            os.mkdir(self.dir_name)
        except FileExistsError:
            pass

        file_list = os.listdir(self.path)
        if not file_list:
            self.end = 0
            print(f"\n  请将Excel文档复制至 -> {self.path}")
            return True

        for file_name in file_list:
            tmp_file_list = file_name.split('.')
            for ch in ['xls', 'xlsx']:
                if ch in tmp_file_list:
                    self.file_list.append(file_name)

    def write_excel(self):
        '''
        新建文档写入数据(覆盖写入)
        :param data_file_name: 文档名
        :return: None
        '''
        try:
            file = xlwt.Workbook(encoding='utf-8', style_compression=0)
            sheet = file.add_sheet('sheet1', cell_overwrite_ok=True)
            for i in range(len(self.value)):
                self.count += 1
                for j in range(len(self.value[i])):
                    sheet.write(i, j, self.value[i][j])

            file.save(self.fina_file)
            print('-'*46)
            print(f'『{self.fina_file}』写入成功 | 共写入 {self.count} 条信息')
        except:
            print('-'*46)
            print(f'『{self.fina_file}』写入失败')

    def weight_score(self, filename):
        '''
        读取成绩单，提取必修课程相关信息并计算
        :param filename: 成绩单文档名
        :return: None
        '''
        studentID = ''     # 学生学号
        studentName = ''   # 学生姓名
        all_count = 0      # 全部课程总门数
        all_credit = 0     # 全部课程总学分
        all_grade = 0      # 全部课程总成绩

        sum_count1 = 0     # 必修课程总门数
        sum_credit1 = 0    # 必修课程总学分
        sum_grade1 = 0     # 必修课程总成绩

        sum_count2 = 0     # 必修课程总门数
        sum_credit2 = 0    # 必修课程总学分
        sum_grade2 = 0     # 必修课程总成绩

        sum_GDA = 0        # 必修课程学分绩点乘积求和
        sum_score = 0      # 必修课程成绩学分乘积求和

        data = pandas.read_excel(self.path+filename)

        for ch in ["学号", "姓名", "成绩", "学分", "绩点"]:
            if ch not in data:
                self.end = 0
                self.flag = 0
                print(f"  [{filename}] 缺失 [{ch}] 数据！")

        if not self.flag:
            print("  - 请在文档中添加缺失数据或检查更正相关数据表头")
            return

        for i in range(len(data)):
            if "学号" in data and data["学号"][i]:
                studentID = int(data["学号"][i])
            if "姓名" in data and data["姓名"][i]:
                studentName = data["姓名"][i]

            if data["成绩"][i] == "优秀":
                data["成绩"][i] = 95
            elif data["成绩"][i] == "良好":
                data["成绩"][i] = 85
            elif data["成绩"][i] == "中等":
                data["成绩"][i] = 75
            elif data["成绩"][i] == "及格":
                data["成绩"][i] = 65
            elif data["成绩"][i] == "不及格":
                data["成绩"][i] = 55

            all_count += 1
            all_credit += float(data["学分"][i])
            all_grade += float(data["成绩"][i])

            if data["课程性质"][i] == "必修课程" and self.mode in ["1", "2"]:
                sum_count1 += 1
                sum_credit1 += float(data["学分"][i])
                sum_grade1 += float(data["成绩"][i])
                sum_GDA += float(data["绩点"][i]) * float(data["学分"][i])
                sum_score += float(data["成绩"][i]) * float(data["学分"][i])
                # print(f'{data["课程名称"][i]}   {data["成绩"][i]}  {data["学分"][i]}  {type(data["学分"][i])}')
            elif data["课程性质"][i] == "选修课程" and self.mode in ["2"]:
                sum_count2 += 1
                sum_credit2 += float(data["学分"][i])
                sum_grade2 += float(data["成绩"][i])
                sum_GDA += float(data["绩点"][i]) * float(data["学分"][i])
                sum_score += float(data["成绩"][i]) * float(data["学分"][i])

        if self.mode == "1":
            score = sum_score / (sum_credit1)    # 加权平均分(只含必修课)
            GDA = sum_GDA / (sum_credit1)        # 平均绩点 (只含必修课)
        elif self.mode == "2":
            score = sum_score / (sum_credit1+sum_credit2)    # 加权平均分(必修+选修)
            GDA = sum_GDA / (sum_credit1+sum_credit2)        # 平均绩点(必修+选修)

        score = self.self_round(score, 2)
        GDA = self.self_round(GDA, 2)

        if self.mode == "1":
            self.value.append(["", studentID, studentName, all_count, sum_count1, all_grade, sum_grade1, all_credit, sum_credit1, score, GDA])
        elif self.mode == "2":
            self.value.append(["", studentID, studentName, all_count, sum_count1, sum_count2, all_grade, sum_grade1, sum_grade2, all_credit, sum_credit1, sum_credit2, score, GDA])

        if self.debug_mode:
            print("共计必修课程 {:d}\n已修学分 {:.2f}\n必修学分 {:.2f}\n总成绩 {:.2f}\n加权平均分 {:.2f}\n平均绩点 {:.2f}"
                  .format(sum_count1, all_credit, sum_credit1, sum_grade1, score, GDA))

    def main(self):
        if self.get_dir_file() or self.self_init():  # 用户输入错误,退出程序
            self.self_exit()
            return

        self.value.append(self.headline)

        if self.file_list:
            for file in self.file_list:
                print(f"\n> {self.path+file}")
                self.flag = 1
                self.weight_score(file)

            self.write_excel()

        self.self_exit()


get_weight_score().main()