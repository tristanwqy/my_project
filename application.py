import collections
import datetime
import tkinter as tk
from tkinter import messagebox, ttk

import pandas as pd

from salary import SalaryCalculator


class MyApplication(object):
    def __init__(self):
        self.root = tk.Tk()
        self.root.geometry("1920x1080")
        self.root.title("随便试试")
        self.persons = ["何宛余", "杨小荻", "李春", "郑夏丽", "魏启赟", "曾古", "邓燊", "孔明"]
        self.init_data()
        self.init_app()
        self.root.mainloop()

    def init_app(self):
        self.init_combobox()
        self.init_global_data_entry()
        self.init_button()
        self.init_person_detail_entries()

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
        self.default_salary.insert(tk.END, '2130')
        self.default_salary.grid(column=1, row=row)

        row = 3
        default_max_transfer_value_label = ttk.Label(self.root, text="宛余设定的报销限额:")
        default_max_transfer_value_label.grid(column=0, row=row)
        self.default_max_transfer_value = ttk.Entry(self.root, width=12)
        self.default_max_transfer_value.insert(tk.END, '10000')
        self.default_max_transfer_value.grid(column=1, row=row)

    def init_person_detail_entries(self, row=5):
        # 是否歪果仁
        ttk.Label(self.root, text="是否歪果仁(肯定不导出)").grid(column=0, row=row)
        tk_bool = tk.StringVar()
        self.is_waiguoren = ttk.Combobox(self.root, width=10, textvariable=tk_bool, values=("是", "否"))
        self.is_waiguoren.current(1)
        self.is_waiguoren.grid(column=0, row=row + 1)
        # self.is_waiguoren.bind("<<ComboboxSelected>>", self._show_person_chosen2)

        self.columns = ["编号", "姓名", "合计款项", "正式\试用期工资占比", "本周工作日", "出勤天数", "本月应收款项",
                        "补贴", "报销", "工资+补贴+报销", "医保档次(共三档，不会导出)", "社保减扣", "公积金比例(不会导出)", "公积金减扣",
                        "基础薪金(计税)", "税前工资", "代扣个人所得税", "实发转账工资", "实发报销工资", "实发保险", "合计金额"]
        self.eng_columns = ["uid", "name", "salary", "salary_rate", "working_day", "present_working_day", "real_salary",
                            "pension", "reimbursement", "real_total_salary", "yibao_level", "social_security_total", "housing_fund_rate", "housing_fund",
                            "base_salary", "salary_for_tax", "tax", "transfer_salary", "transfer_reimbursement", "transfer_insurance", "transfer_total"]

        self.read_only_columns = ["name", "working_day", "real_total_salary", "social_security_total", "housing_fund", "salary_for_tax", "tax",
                                  "transfer_insurance", "transfer_total"]

        max_column_per_line = 12
        for i, column in enumerate(self.columns):
            j = i + 1
            this_row = row + int(j / max_column_per_line) * 2
            this_column = j % max_column_per_line
            ttk.Label(self.root, text=column, width=18).grid(column=this_column, row=this_row, ipadx=2, ipady=2, sticky="s")
            entry = ttk.Entry(self.root, width=18)
            if self.eng_columns[i] in self.read_only_columns:
                entry["state"] = 'readonly'
            setattr(self, self.eng_columns[i], entry)
            entry.insert(tk.END, "0")
            entry.grid(column=this_column, row=this_row + 1)
            # entry.bind("<KeyRelease>", self._bind_table_to_salary_instance)

    def init_button(self, row=10):
        self.recalculate_button = ttk.Button(self.root, text="重新计算", command=self._bind_table_to_salary_instance)
        if not self.selected_person:
            self.recalculate_button["state"] = 'disabled'
        self.recalculate_button.grid(column=0, row=row)

        self.export_single_excel_button = ttk.Button(self.root, text="导出当前童鞋的excel", command=self._export_single_excel)
        if not self.selected_person:
            self.export_single_excel_button["state"] = 'disabled'
        self.export_single_excel_button.grid(column=1, row=row)

        self.export_all_excel_button = ttk.Button(self.root, text="导出所有已编辑人的excel", command=self._bind_table_to_salary_instance)
        if not self.selected_person:
            self.export_all_excel_button["state"] = 'disabled'
        self.export_all_excel_button.grid(column=2, row=row)

    def init_data(self):
        self.selected_person = None
        self.salary_dict = dict()

    def init_combobox(self, row=4):
        ttk.Label(self.root, text="你选中的童鞋是").grid(column=0, row=row)
        tk_string = tk.StringVar()
        person_chosen = ttk.Combobox(self.root, width=10, textvariable=tk_string)
        person_chosen["values"] = self.persons
        person_chosen.grid(column=1, row=row)
        person_chosen.current(0)
        self.person_chosen = person_chosen
        self.person_chosen.bind("<<ComboboxSelected>>", self._update_person)

    def _update_person(self, event):
        self.selected_person = self.person_chosen.get()
        print("被选中的人是" + self.selected_person)
        # 更新名字
        self._update_entry(entry=self.name,
                           new_value=self.selected_person)
        # 按钮启用
        self.recalculate_button["state"] = 'active'
        self.export_single_excel_button["state"] = 'active'
        self.export_all_excel_button["state"] = 'active'
        if self.selected_person not in self.salary_dict:
            if self.selected_person.lower() in ["jakie", "孔明"]:
                self.is_waiguoren.set("是")
            else:
                self.is_waiguoren.set("否")
            is_chinese = self.is_waiguoren.get() == "否"
            salary_instance = SalaryCalculator(default_max_transfer_value=float(self.default_max_transfer_value.get()),
                                               uid="编号别忘了啦",
                                               name=self.selected_person,
                                               is_chinese=is_chinese,
                                               salary=30000,
                                               salary_rate=1,
                                               working_day=int(self.default_working_day.get()),
                                               present_working_day=int(self.default_working_day.get()),
                                               base_salary=10400,
                                               pension=0,
                                               reimbursement=0,
                                               yibao_level=2,
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
        is_chinese = self.is_waiguoren.get() == "否"
        uid = self.uid.get()
        salary = float(self.salary.get())
        salary_rate = float(self.salary_rate.get())
        working_day = int(self.working_day.get())
        present_working_day = int(self.present_working_day.get())
        pension = float(self.pension.get())
        reimbursement = float(self.reimbursement.get())
        yibao_level = int(self.yibao_level.get())
        housing_fund_rate = float(self.housing_fund_rate.get())

        # 基础薪金
        base_salary = float(self.base_salary.get())
        transfer_reimbursement = float(self.transfer_reimbursement.get())
        real_total_salary = float(self.real_total_salary.get())
        transfer_reimbursement = min(real_total_salary - base_salary, transfer_reimbursement)
        # 计算社保的部分
        social_security_base = float(self.default_salary.get())
        new_salary_instance = SalaryCalculator(default_max_transfer_value=float(self.default_max_transfer_value.get()),
                                               uid=uid,
                                               name=self.selected_person,
                                               is_chinese=is_chinese,
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

    def _export_single_excel(self, *args, **kwargs):
        result = messagebox.askokcancel('提示', "你真的要生成{}的薪酬excel么".format(self.selected_person))
        if result:
            columns = ["编号", "姓名", "合计款项", "正式\试用期工资占比", "本周工作日", "出勤天数", "本月应收款项",
                       "补贴", "报销", "工资+补贴+报销", "社保减扣", "公积金减扣",
                       "基础薪金", "税前工资", "代扣个人所得税", "实发转账工资", "实发报销工资", "实发保险", "合计金额"]
            data_dict = collections.OrderedDict()
            salary_instance = self.salary_dict[self.selected_person]
            salary_dict = salary_instance.export()
            for column in columns:
                if column in data_dict:
                    data_dict[column].append(salary_dict.get(column))
                else:
                    data_dict[column] = [salary_dict.get(column)]
            df = pd.DataFrame(data_dict)

            writer = pd.ExcelWriter("{}年{}月薪酬详情表格{}.xlsx".format(self.default_year.get(),
                                                                 self.default_month.get(),
                                                                 self.selected_person), engine='xlsxwriter')
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


if __name__ == '__main__':
    app = MyApplication()
