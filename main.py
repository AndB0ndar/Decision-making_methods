from openpyxl import load_workbook


def pars(c_app):
	wb = load_workbook('./res/applications.xlsx')
	sheet = wb.active
	weight = [cell.value for cell in sheet['B1':'H1'][0]]
	apps = {}
	for row in sheet['A3':'H'+str(3+c_app-1)]:
		apps[row[0].value] = [cell.value for cell in row[1:]]
	return weight, apps


def normalization(apps):
	mx = list(apps.values())
	x_max = [max([x[i] for x in mx]) for i in range(len(mx[0]))]
	vec = mx.copy()
	for i in range(len(vec)):
		for j in range(len(vec[i])):
			vec[i][j] /= x_max[j]
	print(vec)


def saw(w, apps):
	return [round(sum([apps[i][j] * w[j] for j in range(len(apps[i]))]), 3) for i in range(len(apps))]


def mout(w, apps):
	x_max = [max([x[i] for x in apps]) for i in range(len(apps[0]))]
	x_min = [min([x[i] for x in apps]) for i in range(len(apps[0]))]
	matrix = apps.copy()
	for i in range(len(matrix)):
		for j in range(len(matrix[0])):
			matrix[i][j] = round((matrix[i][j] - x_min[j]) / (x_max[j] - x_min[j]), 3)
	return [round(sum([matrix[i][j] * w[j] for j in range(len(matrix[i]))]), 3) for i in range(len(matrix))]


if __name__ == '__main__':
	c_app = 4
	methods = ['saw', 'mout', 'topical']

	weight, dt_apps = pars(c_app)
	apps = list(dt_apps.keys())
	matrix = list(dt_apps.values())

	r_saw = saw(weight, matrix)
	r_maut = mout(weight, matrix)

	print('\\', *methods)
	for i in range(len(r_saw)):
		print(apps[i], r_saw[i], r_maut[i])
