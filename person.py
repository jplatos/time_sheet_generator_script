import datetime

class MonthRecord:
	def __init__(self):
		self.contract_type = ""
		self.multiplicative = 0.0
		self.additive = 0.0
		self.projectTotal = 0
		self.projectAmount = 0
		self.work_holyday_hours = 0
		self.project_holyday_hours =0
		self.work_sickness_days = 0
		self.project_sickness_hours = 0
		self.work_obstacles_hours = 0
		self.project_obstacles_hours = 0
		self.work_state_holydays_hours = 0
		self.project_state_holyday_hours = 0
		self.obstacles_hours = 0
		self.study_program = ''
		self.work_done_start = 0

		self.attendance = {}

	@property
	def projectAmount_str(self):
		if self.projectAmount==1.0:
			return '1,0'
		else:
			return '{0:.03g}'.format(self.projectAmount).replace('.', ',')
		

		

	@property
	def multiplicative_str(self):
		if self.multiplicative==1.0:
			return '1,0'
		else:
			return '{0:.03g}'.format(self.multiplicative).replace('.', ',')
		 
	@property
	def work_total_str(self):
		if (self.additive is None or self.additive == 0):
			return self.multiplicative_str
		else:
			return ("{0:}+{1:.3g}".format(self.multiplicative_str, self.additive)).replace(".", ",");
	
	@property
	def is_contract(self):
		return self.contract_type!='DPP' and self.contract_type!='DPČ'
	
	def holidays_str(self):
		result = []
		for day,value in sorted(self.attendance.items()):
			if value=='d' or value=='D':
				result.append(day)
		return ', '.join(str(x) for x in result)

	def holidays_days(self):
		result = 0
		for day,value in sorted(self.attendance.items()):
			if value=='d':
				result += 0.5
			elif value=='D':
				result += 1
		return result

	def illnesses_str(self):
		result = []
		for day,value in sorted(self.attendance.items()):
			if value=='N':
				result.append(day)
		return ','.join(str(x) for x in result)

	def illnesses_days(self):
		result = 0
		for day,value in sorted(self.attendance.items()):
			if value=='N':
				result += 1
		return result

	def obstacles_str(self):
		result = []
		for day,value in sorted(self.attendance.items()):
			if value=='P' or value=='p' or value=='*'  or value=='C'  or value=='O'  or value=='o':
				result.append(day)
		return ','.join(str(x) for x in result)

	def obstacles_days(self):
		result = 0
		for day,value in sorted(self.attendance.items()):
			if value=='P' or value=='C'  or value=='O':
				result += 1
			elif value=='*':
				result += 1
			elif value=='p'  or value=='o':
				result += 0.5
		return result
		
class Person:
	def __init__(self):
		self.sap_number = ""
		self.nick = ""
		self.firstname = ""
		self.lastname = ""
		self.project = ""
		self.row_code = ""
		self.position = ""
		self.spp = ""
		self.key_activity = ""
		self.activity = ""
		self.email = ""
		self.records = []
		self.last_date = datetime.datetime(1900, 1, 1)
		self.approver = ""
		self.approver_position = ""
		self.report_year = ""
		self.report_month = ""
		self.file_name = ''

	def __repr__(self):
		return "[{}, {}]".format(self.name, self.activity)

	def contract_amounts(self, month_names):
		if (self.records[0].is_contract):
			amounts = [x.projectAmount for x in self.records]
			if len(set(amounts)) == 1:
				return amounts[0]
			else:
				zipped = zip(month_names, amounts)
				lines = ["{}: {:.3g}".format(x[0], x[1]) for x in zipped]
				return "\n".join(lines).replace(".", ",")
		else:
			amounts = [x.projectTotal for x in self.records]
			if (len(amounts)==1): return amounts[0]
			zipped = zip(month_names, amounts)
			lines = ["{}: {:.3g}".format(x[0], x[1]) for x in zipped]
			return "\n".join(lines).replace(".", ",")


	def contract_amounts_str(self, month_names):
		if (self.records[0].is_contract):
			amounts = [x.projectAmount for x in self.records]
			if len(set(amounts)) == 1:
				return amounts[0]
			else:
				zipped = zip(month_names, amounts)
				lines = ["{}: {:.3g}".format(x[0], x[1]) for x in zipped]
				return "\n".join(lines).replace(".", ",")
		else:
			amounts = [x.projectTotal for x in self.records]
			if (len(amounts)==1): return amounts[0]
			zipped = zip(month_names, amounts)
			lines = ["{}: {:.3g}".format(x[0], x[1]) for x in zipped]
			return "\n".join(lines).replace(".", ",")

	def contract_types(self, month_names):
		types = [x.contract_type for x in self.records]
		if len(set(types)) == 1:
			return types[0]
		else:
			return ",".join(types)
	
	def total_amounts(self, month_names):
		types = [x.work_total_str for x in self.records]
		if len(set(types)) == 1:
			return types[0]
		else:
			zipped = zip(month_names, types)
			lines = ["{}: {}".format(x[0], x[1]) for x in zipped]
			return "\n".join(lines).replace(".", ",")
	