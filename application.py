import collections
import datetime
import json
import os
import tkinter as tk
from tkinter import messagebox, ttk
from tkinter.filedialog import askdirectory

import pandas as pd

from salary import SalaryCalculator


class MyApplication(object):
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry("1920x1080")
        self.root.title("随便试试")
        self.persons = [
            # 小库
            "何宛余",
            "杨小荻",
            "李春",
            "郑夏丽",
            "曾古",
            "魏启赟",
            "吕沙",
            "邓燊",
            "刘轲",
            "聂广洋",
            "杨蕴琳",
            "谈仁超",
            "王箫",
            "孔明",
            "吴磊",
            # 建筑
            "刘云祥",
            "吴佳晶",
            "杨良崧",
            "黄辉雄",
            "罗佳敏",
            "李政",
            "沈睿卿",
            # 阿姨
            "王梅香",
        ]
        self.uids = {
            # 小库
            "何宛余": "001",
            "杨小荻": "002",
            "李春":  "003",
            "郑夏丽": "101",
            "曾古":  "102",
            "魏启赟": "103",
            "吕沙":  "104",
            "邓燊":  "105",
            "刘轲":  "106",
            "聂广洋": "107",
            "杨蕴琳": "108",
            "谈仁超": "109",
            "王箫":  "110",
            "孔明":  "I102",
            "吴磊":  "I101",
            # 建筑
            "刘云祥": "103",
            "吴佳晶": "101",
            "杨良崧": "102",
            "黄辉雄": "104",
            "罗佳敏": "105",
            "李政":  "I104",
            "沈睿卿": "I105",
            # 阿姨
            "王梅香": "A",
        }

        self.foreigners = ["孔明", "杨良崧"]
        self.shenzheners = ["吕沙", "刘轲", "聂广洋", "何宛余", "王箫", "曾古", "杨蕴琳", "邓燊", "罗佳敏", "吴佳晶", "黄辉雄", "李政"]
        self.interns = ["吴磊", "李政", "孔明", "沈睿卿"]
        self.table_edited = True
        self.edited_people = []
        self.init_data()
        self.init_app()
        self.root.mainloop()

    def init_app(self):
        self.init_combobox()
        self.init_global_data_entry()
        self.init_button()
        self.init_person_detail_entries()
        self.init_info()

    def init_global_data_entry(self):
        row = 0
        default_year_label = ttk.Label(self.root, text="年:")
        default_year_label.grid(column=0, row=row)
        self.default_year = ttk.Entry(self.root, width=12)
        self.default_year.insert(tk.END, datetime.datetime.now().year)
        self.default_year.grid(column=1, row=row)
        default_month_label = ttk.Label(self.root, text="月:")
        default_month_label.grid(column=2, row=row)
        self.default_month = ttk.Entry(self.root, width=12)
        self.default_month.insert(tk.END, datetime.datetime.now().month)
        self.default_month.grid(column=3, row=row)

        row = 1
        default_working_day_label = ttk.Label(self.root, text="本周工作日天数是:")
        default_working_day_label.grid(column=0, row=row)
        self.default_working_day = ttk.Entry(self.root, width=12)
        self.default_working_day.insert(tk.END, '22')
        self.default_working_day.grid(column=1, row=row)

        row = 2
        default_salary_label = ttk.Label(self.root, text="默认计算社保部分:")
        default_salary_label.grid(column=0, row=row)
        self.default_salary = ttk.Entry(self.root, width=12)
        self.default_salary.insert(tk.END, self.default_salary_cache or '2130')
        self.default_salary.grid(column=1, row=row)

        row = 3
        default_max_transfer_value_label = ttk.Label(self.root, text="宛余设定的报销限额:")
        default_max_transfer_value_label.grid(column=0, row=row)
        self.default_max_transfer_value = ttk.Entry(self.root, width=12)
        self.default_max_transfer_value.insert(tk.END, self.default_max_transfer_value_cache or '10000')
        self.default_max_transfer_value.grid(column=1, row=row)

        row = 4
        self.folder = tk.StringVar()
        ttk.Label(self.root, text="保存excel的文件夹:").grid(row=row, column=0)
        ttk.Entry(self.root, textvariable=self.folder, width=60).grid(row=row, column=1, columnspan=5)
        ttk.Button(self.root, text="选择", command=self._select_folder).grid(row=row, column=6)

    def init_person_detail_entries(self, row=7):
        self.columns = ["是否中国人", "是否深户", "是否实习生", "编号", "姓名", "合计款项", "正式/试用期工资占比", "本周工作日", "出勤天数", "本月应收款项",
                        "补贴", "报销", "工资+补贴+报销", "医保档次", "社保减扣", "公积金比例", "公积金减扣",
                        "基础薪金(计税部分)", "税前工资", "代扣个人所得税", "实发转账工资", "实发报销工资", "实发保险", "合计金额"]
        self.eng_columns = ["is_chinese", "is_shenzhen", "is_intern", "uid", "name", "salary", "salary_rate", "working_day", "present_working_day",
                            "real_salary",
                            "pension", "reimbursement", "real_total_salary", "yibao_level", "social_security_total", "housing_fund_rate", "housing_fund",
                            "base_salary", "salary_for_tax", "tax", "transfer_salary", "transfer_reimbursement", "transfer_insurance", "transfer_total"]

        self.read_only_columns = ["is_chinese", "is_shenzhen", "is_intern", "uid", "name", "real_salary", "working_day", "yibao_level", "real_total_salary",
                                  "social_security_total", "housing_fund",
                                  "salary_for_tax", "tax",
                                  "transfer_insurance", "transfer_total"]

        max_column_per_line = 12
        for i, column in enumerate(self.columns):
            j = i
            this_row = row + int(j / max_column_per_line) * 2
            this_column = j % max_column_per_line
            ttk.Label(self.root, text=column, width=18).grid(column=this_column, row=this_row, ipadx=2, ipady=2, sticky="s")
            entry = ttk.Entry(self.root, width=18)
            if self.eng_columns[i] in self.read_only_columns:
                entry["state"] = 'readonly'
            setattr(self, self.eng_columns[i], entry)
            entry.insert(tk.END, "0")
            entry.grid(column=this_column, row=this_row + 1)
            entry.bind("<KeyRelease>", self._table_changed)

    def init_button(self, row=11):
        self.recalculate_button = ttk.Button(self.root, text="确认并重新计算", command=self._bind_table_to_salary_instance)
        if not self.selected_person:
            self.recalculate_button["state"] = 'disabled'
        self.recalculate_button.grid(column=0, row=row)

        self.export_single_excel_button = ttk.Button(self.root, text="导出当前童鞋的excel", command=self._export_single_excel)
        if not self.selected_person:
            self.export_single_excel_button["state"] = 'disabled'
        self.export_single_excel_button.grid(column=1, row=row)

        self.export_all_excel_button = ttk.Button(self.root, text="导出所有已编辑同学们的excel", command=self._export_all_edited_excel)
        if not self.selected_person:
            self.export_all_excel_button["state"] = 'disabled'
        self.export_all_excel_button.grid(column=2, row=row)

    def init_info(self):
        ttk.Label(self.root, text="已经计算过的盘友们是:").grid(column=0, row=6)
        self.edited_people_entry = ttk.Entry(self.root, width=180)
        self.edited_people_entry["state"] = 'readonly'
        self.edited_people_entry.grid(column=1, row=6, columnspan=10)
        info1 = ttk.Entry(self.root, width=80)
        info1.insert(tk.END, "社保计算方式参考: https://wenku.baidu.com/view/6dded89c710abb68a98271fe910ef12d2bf9a96c.html")
        info1["state"] = 'readonly'
        info1.grid(column=0, row=12, columnspan=8)
        info2 = ttk.Entry(self.root, width=80)
        info2.insert(tk.END, "个人所得税计算参考：http://www.gerensuodeshui.cn/")
        info2["state"] = 'readonly'
        info2.grid(column=0, row=13, columnspan=8)

    def init_data(self):
        self.selected_person = None
        self.default_max_transfer_value_cache, self.default_salary_cache, self.salary_dict = self._load_salary_dict_from_cache()

    def init_combobox(self, row=5):
        ttk.Label(self.root, text="你选中的童鞋是").grid(column=0, row=row)
        tk_string = tk.StringVar()
        person_chosen = ttk.Combobox(self.root, width=10, textvariable=tk_string)
        person_chosen["values"] = self.persons
        person_chosen.grid(column=1, row=row)
        person_chosen.current(0)
        self.person_chosen = person_chosen
        self.person_chosen.bind("<<ComboboxSelected>>", self._update_person)

    def _get_person_info(self, person_name):
        return (self.uids[person_name],
                "是" if person_name not in self.foreigners else "否",
                "是" if person_name in self.shenzheners else "否",
                "是" if person_name in self.interns else "否")

    def _update_person(self, event):
        result = True
        if self.table_edited and self.selected_person:
            result = messagebox.askokcancel("提示", "{}的表格被修改过，但是并未重新计算并导出，是否立即跳转？".format(self.selected_person))
        if not result:
            # 回滚
            self.person_chosen.set(self.selected_person)
        else:
            self.selected_person = self.person_chosen.get()
            print("被选中的人是" + self.selected_person)
            # 更新名字
            self._update_entry(entry=self.name,
                               new_value=self.selected_person)
            # 把表单状态改为编辑完成
            self.table_edited = False
            # 按钮启用
            self.recalculate_button["state"] = 'active'
            self.export_single_excel_button["state"] = 'active'
            self.export_all_excel_button["state"] = "active"
            # self.export_all_excel_button["state"] = 'active'
            if self.selected_person not in self.salary_dict:
                uid, is_chinese, is_shenzhen, is_intern = self._get_person_info(person_name=self.selected_person)
                yibao_level = 1 if is_shenzhen == "是" else 3
                salary_instance = SalaryCalculator(default_max_transfer_value=float(self.default_max_transfer_value.get()),
                                                   uid=uid,
                                                   name=self.selected_person,
                                                   is_chinese=is_chinese,
                                                   is_shenzhen=is_shenzhen,
                                                   is_intern=is_intern,
                                                   salary=30000,
                                                   salary_rate=1,
                                                   working_day=int(self.default_working_day.get()),
                                                   present_working_day=int(self.default_working_day.get()),
                                                   base_salary=10400,
                                                   pension=5100,
                                                   reimbursement=0,
                                                   yibao_level=yibao_level,
                                                   housing_fund_rate=0.05,
                                                   social_security_base=float(self.default_salary.get()),
                                                   transfer_reimbursement=10000)
                self.salary_dict[self.selected_person] = salary_instance
                self._refresh_salary_table(salary_instance)
            else:
                salary_instance = self.salary_dict[self.selected_person]
                self._refresh_salary_table(salary_instance)

    def _refresh_salary_table(self, salary_instance):
        """
        eng_columns = ["uid", "name", "salary", "salary_rate", "working_day", "present_working_day", "real_salary",
               "pension", "reimbursement", "real_total_salary", "yibao_level", "social_security_total", "housing_fund_rate", "housing_fund",
               "base_salary", "salary_for_tax", "tax", "transfer_salary", "transfer_reimbursement", "transfer_insurance", "transfer_total"]
        :param salary_instance:
        :return:
        """
        self._update_entry(self.is_chinese, "是" if salary_instance.is_chinese else "否")
        self._update_entry(self.is_shenzhen, "是" if salary_instance.is_shenzhen else "否")
        self._update_entry(self.is_intern, "是" if salary_instance.is_intern else "否")
        self._update_entry(self.uid, salary_instance.uid)
        self._update_entry(self.salary, salary_instance.salary)
        self._update_entry(self.salary_rate, salary_instance.salary_rate)
        self._update_entry(self.working_day, salary_instance.working_day)
        self._update_entry(self.present_working_day, salary_instance.present_working_day),
        self._update_entry(self.real_salary, salary_instance.real_salary),
        self._update_entry(self.pension, salary_instance.pension),
        self._update_entry(self.reimbursement, salary_instance.reimbursement),
        self._update_entry(self.real_total_salary, salary_instance.real_total_salary)
        self._update_entry(self.yibao_level, salary_instance.yibao_level)
        self._update_entry(self.social_security_total, salary_instance.social_security_total),
        self._update_entry(self.housing_fund, salary_instance.housing_fund),
        self._update_entry(self.housing_fund_rate, salary_instance.housing_fund_rate)
        self._update_entry(self.base_salary, salary_instance.base_salary)
        self._update_entry(self.salary_for_tax, salary_instance.salary_for_tax)
        self._update_entry(self.tax, salary_instance.tax)
        self._update_entry(self.transfer_salary, salary_instance.transfer_salary)
        self._update_entry(self.transfer_reimbursement, salary_instance.transfer_reimbursement)
        self._update_entry(self.transfer_insurance, salary_instance.transfer_insurance)
        self._update_entry(self.transfer_total, salary_instance.transfer_total)

    def _bind_table_to_salary_instance(self, *args, **kwargs):
        """
        eng_columns = ["uid", "name", "salary", "salary_rate", "working_day", "present_working_day", "real_salary",
               "pension", "reimbursement", "real_total_salary", "yibao_level", "social_security_total", "housing_fund_rate", "housing_fund",
               "base_salary", "salary_for_tax", "tax", "transfer_salary", "transfer_reimbursement", "transfer_insurance", "transfer_total"]
        :return:
        """
        self.table_edited = False
        uid, is_chinese, is_shenzhen, is_intern = self._get_person_info(self.selected_person)
        salary = float(self.salary.get())
        salary_rate = float(self.salary_rate.get())
        working_day = int(self.working_day.get())
        present_working_day = int(self.present_working_day.get())
        pension = float(self.pension.get())
        reimbursement = float(self.reimbursement.get())
        yibao_level = 1 if is_shenzhen == "是" else 3
        housing_fund_rate = float(self.housing_fund_rate.get())

        # 基础薪金
        base_salary = float(self.base_salary.get())
        transfer_reimbursement = float(self.transfer_reimbursement.get())
        # 计算社保的部分
        social_security_base = float(self.default_salary.get())
        new_salary_instance = SalaryCalculator(default_max_transfer_value=float(self.default_max_transfer_value.get()),
                                               uid=uid,
                                               name=self.selected_person,
                                               is_chinese=is_chinese,
                                               is_shenzhen=is_shenzhen,
                                               is_intern=is_intern,
                                               salary=salary,
                                               salary_rate=salary_rate,
                                               working_day=working_day,
                                               present_working_day=present_working_day,
                                               base_salary=base_salary,
                                               pension=pension,
                                               reimbursement=reimbursement,
                                               yibao_level=yibao_level,
                                               housing_fund_rate=housing_fund_rate,
                                               social_security_base=social_security_base,
                                               transfer_reimbursement=transfer_reimbursement)
        self.salary_dict[self.selected_person] = new_salary_instance
        self._refresh_salary_table(salary_instance=new_salary_instance)
        self._save_current_stats()
        if self.selected_person not in self.edited_people:
            self.edited_people.append(self.selected_person)
        self._update_entry(self.edited_people_entry, self.edited_people)

    def _export_single_excel(self, *args, selected_person=None, ask_confirm=True, **kwargs):
        selected_person = selected_person or self.selected_person
        folder = self.folder.get()
        if not folder:
            messagebox.showinfo("警告", "你还没选要把excel存在哪啦")
            return
        if self.table_edited:
            messagebox.showinfo("警告", "你上次修改过后还没重新计算哦")
            return
        file_path = os.path.join(folder, "{}年{}月薪酬详情表格{}.xlsx".format(
            self.default_year.get(),
            self.default_month.get(),
            selected_person))
        if ask_confirm:
            result = messagebox.askokcancel('提示', "你真的要生成{}的薪酬excel么\n{}的excel将会存储在:\n{}".format(selected_person, selected_person, file_path))
        else:
            result = True
        if result:
            columns = ["编号", "姓名", "合计款项", "正式/试用期工资占比", "本周工作日", "出勤天数", "本月应收款项",
                       "补贴", "报销", "工资+补贴+报销", "社保减扣", "公积金减扣",
                       "基础薪金", "税前工资", "代扣个人所得税", "实发转账工资", "实发报销工资", "实发保险", "合计金额"]
            data_dict = collections.OrderedDict()
            salary_instance = self.salary_dict[selected_person]
            salary_dict = salary_instance.export()
            for column in columns:
                if column in data_dict:
                    data_dict[column].append(salary_dict.get(column))
                else:
                    data_dict[column] = [salary_dict.get(column)]
            df = pd.DataFrame(data_dict)

            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            df.to_excel(writer, sheet_name='Sheet1', index=False)
            worksheet = writer.sheets['Sheet1']
            for i, column in enumerate(columns):
                width = max(len(column) * 2, 10)
                worksheet.set_column(firstcol=i, lastcol=i, width=width)
            writer.save()

    def _export_all_edited_excel(self, *args, **kwargs):
        folder = self.folder.get()
        if not folder:
            messagebox.showinfo("警告", "你还没选要把excel存在哪啦")
            return
        file_path = os.path.join(folder, "小库科技{}年{}月薪酬详情表格总表.xlsx".format(
            self.default_year.get(),
            self.default_month.get()))
        result = messagebox.askokcancel('提示', "你真的要生成薪酬总表么，当前编辑过的同学们有{}\nexcel将会存储在:\n{}".format(self.edited_people, folder))
        if result:
            columns = ["编号", "姓名", "合计款项", "正式/试用期工资占比", "本周工作日", "出勤天数", "本月应收款项",
                       "补贴", "报销", "工资+补贴+报销", "社保减扣", "公积金减扣",
                       "基础薪金", "税前工资", "代扣个人所得税", "实发转账工资", "实发报销工资", "实发保险", "合计金额"]
            data_dict = collections.OrderedDict()
            for selected_person in self.edited_people:
                self._export_single_excel(selected_person=selected_person, ask_confirm=False)
                salary_instance = self.salary_dict[selected_person]
                salary_dict = salary_instance.export()
                for column in columns:
                    if column in data_dict:
                        data_dict[column].append(salary_dict.get(column))
                    else:
                        data_dict[column] = [salary_dict.get(column)]
                df = pd.DataFrame(data_dict)

                writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
                df.to_excel(writer, sheet_name='Sheet1', index=False)
                worksheet = writer.sheets['Sheet1']
                for i, column in enumerate(columns):
                    width = max(len(column) * 2, 10)
                    worksheet.set_column(firstcol=i, lastcol=i, width=width)
                writer.save()

    def _update_entry(self, entry, new_value):
        readonly = str(entry["state"]) == 'readonly'
        entry["state"] = 'active'
        entry.delete(0, tk.END)
        if isinstance(new_value, float):
            new_value = round(new_value, 2)
        entry.insert(tk.END, new_value)
        if readonly:
            entry["state"] = 'readonly'

    def _select_folder(self, *args, **kwargs):
        folder = askdirectory()
        self.folder.set(folder)

    def _table_changed(self, *args, **kwargs):
        self.table_edited = True

    def _save_current_stats(self):
        with open("salary.cache", "w+") as f:
            result_dict = dict()
            for key, salary_instance in self.salary_dict.items():
                result_dict[key] = salary_instance.export()
            data = json.dumps(result_dict)
            f.write(data)
        with open("global.cache", "w+") as f:
            global_dict = dict()
            global_dict["default_max_transfer_value"] = float(self.default_max_transfer_value.get())
            global_dict["default_salary"] = float(self.default_salary.get())
            data = json.dumps(global_dict)
            f.write(data)

    def _load_salary_dict_from_cache(self):
        if os.path.exists("global.cache"):
            with open("global.cache", "r+") as f:
                json_str = f.read()
                global_dict = json.loads(json_str)
                default_max_transfer_value_cache = global_dict["default_max_transfer_value"]
                default_salary_cache = global_dict["default_salary"]
        else:
            default_max_transfer_value_cache = None
            default_salary_cache = None
        if os.path.exists("salary.cache"):
            with open("salary.cache", "r+") as f:
                json_str = f.read()
                salary_dict_cache = dict()
                salary_dict = json.loads(json_str)
                for selected_person, salary_person_dict in salary_dict.items():
                    uid, is_chinese, is_shenzhen, is_intern = self._get_person_info(selected_person)
                    yibao_level = 1 if is_shenzhen == "是" else 3
                    salary_instance = SalaryCalculator(default_max_transfer_value=default_max_transfer_value_cache,
                                                       uid=salary_person_dict["编号"],
                                                       name=selected_person,
                                                       is_chinese=is_chinese,
                                                       is_shenzhen=is_shenzhen,
                                                       is_intern=is_intern,
                                                       salary=salary_person_dict["合计款项"],
                                                       salary_rate=salary_person_dict["正式/试用期工资占比"],
                                                       working_day=22,
                                                       present_working_day=22,
                                                       base_salary=salary_person_dict["基础薪金"],
                                                       pension=salary_person_dict["补贴"],
                                                       reimbursement=salary_person_dict["报销"],
                                                       yibao_level=yibao_level,
                                                       housing_fund_rate=0.05,
                                                       social_security_base=default_salary_cache,
                                                       transfer_reimbursement=salary_person_dict["实发报销工资"])
                    salary_dict_cache[selected_person] = salary_instance
        else:
            salary_dict_cache = dict()
        return default_max_transfer_value_cache, default_salary_cache, salary_dict_cache


if __name__ == '__main__':
    app = MyApplication()
