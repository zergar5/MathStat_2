import pandas as pd
import scipy.stats as stats


def excel_read(file_name):
    with pd.ExcelFile(file_name) as data_file:
        df = pd.read_excel(data_file, 'Лист1', engine='openpyxl')
        pm1_1 = read_data(df, 'ПМ1 1 семестр')
        pm2_1 = read_data(df, 'ПМ2 1 семестр')
        pm1_2 = read_data(df, 'ПМ1 2 семестр')
        pm2_2 = read_data(df, 'ПМ2 2 семестр')

    return pm1_1, pm2_1, pm1_2, pm2_2


def read_data(df, name):
    data = df.loc[:, name]
    return data


def concat_data(pm_1, pm_2):
    pm = pd.concat([pm_1, pm_2], ignore_index=True)
    return pm


def calc_mean(data):
    middle = float(data.mean())
    return formating(middle)


def calc_t_criterion_for_dependent(pm_1, pm_2):
    t_dep, _ = stats.ttest_rel(pm_1, pm_2)

    return formating(t_dep)


def calc_t_criterion_for_independent(pm_1, pm_2):
    t_indep, _ = stats.ttest_ind(pm_1, pm_2)

    return formating(t_indep)

def calc_U_criterion(pm_1, pm_2):
    u1, _ = stats.mannwhitneyu(pm_1, pm_2)
    u2 = len(pm_1) * len(pm_2) - u1
    u = min(u1, u2)
    return formating(u)

def calc_W_criterion(pm_1, pm_2):
    w, _ = stats.wilcoxon(pm_1, pm_2, method='approx')

    return formating(w)


def formating(stat):
    return float('{0:.5}'.format(stat))


data = excel_read("Stat.xlsx")

pm1_1 = data[0]
pm2_1 = data[1]
pm1_2 = data[2]
pm2_2 = data[3]
pm1 = concat_data(pm1_1, pm1_2)
pm2 = concat_data(pm2_1, pm2_2)

pm1_1_mean = calc_mean(pm1_1)
pm2_1_mean = calc_mean(pm2_1)
pm1_2_mean = calc_mean(pm1_2)
pm2_2_mean = calc_mean(pm2_2)
pm1_mean = calc_mean(pm1)
pm2_mean = calc_mean(pm2)

print('Средний балл ПМ1 за 1-ый семестр: ', pm1_1_mean)
print('Средний балл ПМ1 за 2-ой семестр: ', pm1_2_mean)
print('Средний балл ПМ1 за год: ', pm1_mean)
print()
print('Средний балл ПМ2 за 1-ый семестр: ', pm2_1_mean)
print('Средний балл ПМ2 за 2-ой семестр: ', pm2_2_mean)
print('Средний балл ПМ2 за год: ', pm2_mean)
print()

t_dep_1 = formating(calc_t_criterion_for_dependent(pm1_1, pm1_2))
t_dep_2 = calc_t_criterion_for_dependent(pm2_1, pm2_2)

print('t-критерий для зависимой выборки(ПМ1): ', t_dep_1)
print('t-критерий для зависимой выборки(ПМ2): ', t_dep_2)
print('Критический t-критерий для зависимой выборки: ', 2.064)
print()

t_indep_1 = calc_t_criterion_for_independent(pm1_1, pm2_1)
t_indep_2 = calc_t_criterion_for_independent(pm1_2, pm2_2)

print('t-критерий для независимой выборки(1-ый семестр ПМ1 и ПМ2): ', t_indep_1)
print('t-критерий для независимой выборки(2-ой семестр ПМ1 и ПМ2): ', t_indep_2)
print('Критический t-критерий для независимой выборки: ', 2.011)
print()

U_1 = calc_U_criterion(pm1_1, pm2_1)
U_2 = calc_U_criterion(pm1_2, pm2_2)

print('U-критерий для 1-ого семестра: ', U_1)
print('U-критерий для 2-ого семестра: ', U_2)
print('Критический U-критерий: ', 211)
print()

W_1 = calc_W_criterion(pm1_1, pm1_2)
W_2 = calc_W_criterion(pm2_1, pm2_2)

print('W-критерий для ПМ1: ', W_1)
print('Критический W-критерий для ПМ1: ', 60)
print('W-критерий для ПМ2: ', W_2)
print('Критический W-критерий для ПМ2: ', 100)
print()

