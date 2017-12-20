class SocialSecurity(object):
    """
    http://news.vobao.com/zhuanti/885553609845687774.shtml
    """
    LOWEST_SALARY = 2130

    def __init__(self, salary=None, housing_fund_rate=0.05, level=3, is_shenzhen=False):
        if is_shenzhen:
            SHENZHEN_AVAERAGE_SALARY = 7480 * 0.6
        else:
            SHENZHEN_AVAERAGE_SALARY = 7480
        if not salary:
            self.salary = self.LOWEST_SALARY
        else:
            self.salary = salary
        self.yanglao = 0.08 * self.salary
        if is_shenzhen or level == 1:
            self.yiliao = 0.02 * SHENZHEN_AVAERAGE_SALARY
        elif not is_shenzhen and level == 2:
            self.yiliao = 0.002 * SHENZHEN_AVAERAGE_SALARY
        elif not is_shenzhen and level == 3:
            self.yiliao = 0.001 * SHENZHEN_AVAERAGE_SALARY
        else:
            raise Exception
        self.shiye = 0.005 * self.LOWEST_SALARY
        self.security_total = self.yanglao + self.yiliao + self.shiye
        self.housing_fund = housing_fund_rate * self.salary


class TaxCalculator(object):
    """
    http://www.gerensuodeshui.cn/
    """

    @staticmethod
    def calculate_tax(total, is_chinese=True):
        if is_chinese:
            total_for_tax = total - 3500
        else:
            total_for_tax = total - 4800
        if total_for_tax < 0:
            return 0
        elif total_for_tax < 1455:
            return total_for_tax * 0.03
        elif total_for_tax < 4155:
            return total_for_tax * 0.1 - 105
        elif total_for_tax < 7755:
            return total_for_tax * 0.2 - 555
        elif total_for_tax < 27255:
            return total_for_tax * 0.25 - 1005
        elif total_for_tax < 41255:
            return total_for_tax * 0.3 - 2755
        elif total_for_tax < 57505:
            return total_for_tax * 0.35 - 5505
        else:
            return total_for_tax * 0.45 - 13505


class SalaryCalculator(object):
    def __init__(self,
                 default_max_transfer_value,  # 全局的转账限额
                 uid,  # 编号
                 name,  # 姓名
                 is_chinese,  # 是否中国人
                 is_shenzhen,  # 是否深圳人
                 salary,  # 合计款项
                 salary_rate,  # 百分比（实习，正式）
                 working_day,  # 工作日数
                 present_working_day,  # 出勤日数
                 base_salary,  # 基础薪金
                 pension,  # 各类补贴
                 reimbursement,  # 报销

                 # transfer_salary,  # 实发转账
                 transfer_reimbursement,  # 实发报销
                 # transfer_insurance,  # 实发保险
                 yibao_level=2,
                 housing_fund_rate=0.05,
                 social_security_base=None,  # 计算社保的工资，一般为2130
                 ):
        self.uid = uid
        self.name = name
        self.is_chinese = is_chinese
        self.is_shenzhen = is_shenzhen
        self.salary = salary
        self.salary_rate = salary_rate
        self.present_working_day = present_working_day
        self.working_day = working_day
        self.real_salary = self.salary * self.salary_rate * present_working_day / working_day
        self.social_security_base = social_security_base
        # 计算社保和公积金扣减
        self.yibao_level = yibao_level
        self.housing_fund_rate = housing_fund_rate
        if not is_chinese:
            # 歪果仁不交社保
            self.social_security_total = 0
            self.housing_fund = 0
        else:
            social_security = SocialSecurity(salary=social_security_base,
                                             housing_fund_rate=housing_fund_rate,
                                             level=yibao_level,
                                             is_shenzhen=is_shenzhen)
            self.social_security_total = social_security.security_total
            self.housing_fund = social_security.housing_fund

        # 补贴、报销
        self.pension = pension
        self.reimbursement = reimbursement
        self.real_total_salary = self.real_salary + self.pension + self.reimbursement
        # 基础薪金不可以超过总薪金
        self.base_salary = min(base_salary, self.real_salary)

        self.salary_for_tax = self.base_salary - self.social_security_total - self.housing_fund
        self.tax = TaxCalculator.calculate_tax(total=self.salary_for_tax,
                                               is_chinese=is_chinese)
        self.transfer_salary = self.salary_for_tax - self.tax
        # 不可以超过宛余的上限, 也不可以超过了总计薪金减去基础薪金的差额
        self.transfer_reimbursement = min(transfer_reimbursement, default_max_transfer_value, self.real_total_salary - self.base_salary)
        self.transfer_insurance = (self.real_total_salary - self.base_salary) - self.transfer_reimbursement
        self.transfer_total = self.transfer_salary + self.transfer_reimbursement + self.transfer_insurance

    def export(self):
        return {"编号":         self.uid,
                "姓名":         self.name,
                "合计款项":       self.salary,
                "正式/试用期工资占比": self.salary_rate,
                "本周工作日":      self.working_day,
                "出勤天数":       self.present_working_day,
                "本月应收款项":     self.real_salary,
                "补贴":         self.pension,
                "报销":         self.reimbursement,
                "工资+补贴+报销":   self.real_total_salary,
                "社保减扣":       self.social_security_total,
                "公积金减扣":      self.housing_fund,
                "基础薪金":       self.base_salary,
                "税前工资":       self.salary_for_tax,
                "代扣个人所得税":    self.tax,
                "实发转账工资":     self.transfer_salary,
                "实发报销工资":     self.transfer_reimbursement,
                "实发保险":       self.transfer_insurance,
                "合计金额":       self.transfer_total,
                }


if __name__ == '__main__':
    with open("1.txt", "w+") as f:
        f.write("hello2")
