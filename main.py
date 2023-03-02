from openpyxl import load_workbook


def pars(c_app):
	wb = load_workbook('./res/applications.xlsx')
	sheet = wb.active
	weight = [cell.value for cell in sheet['B1':'G1'][0]]
	print(weight)
	apps = {}
	for row in sheet['A3':'H'+str(3+c_app-1)]:
		apps[row[0].value] = [cell.value for cell in row[1:]]
	print(apps)
	return weight, apps



if __name__ == '__main__':
	weight, apps = pars(4)
	print(list(apps.keys()))
	print(list(apps.values()))
