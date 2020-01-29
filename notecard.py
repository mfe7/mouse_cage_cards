import xlsxwriter
from xlrd import open_workbook
import re
import yaml

if __name__ == "__main__":

	# Create a workbook and add a worksheet.
	workbook = xlsxwriter.Workbook('notecards.xlsx')
	worksheet = workbook.add_worksheet()
	worksheet.set_paper(1)  # Letter
	worksheet.set_landscape()

	# Rows/cols are for the spreadsheet. Nominally, 4 cards
	# per sheet of paper, arranged in a 2x2 pattern.
	max_mice_per_cage = 6
	rows_per_card = max_mice_per_cage + 10
	cols_per_card = 6
	cards_per_sheet = 4
	cards_on_current_sheet = 0
	cards_per_row_on_sheet = 2
	cards_per_col_on_sheet = cards_per_sheet / cards_per_row_on_sheet

	page_breaks = []

	# Set Column Widths
	column_widths = [6,11,3,29,3,2,10]
	num_cols = len(column_widths)
	for c in range(num_cols):
		worksheet.set_column(c,c,column_widths[c])
		if c == num_cols - 1:
			worksheet.set_column(c+cols_per_card,c+cols_per_card,0)
			break
		worksheet.set_column(c+cols_per_card,c+cols_per_card,column_widths[c])
	outline_format = workbook.add_format({'border': 1})
	bottom_format = workbook.add_format({'bottom': 1})
	left_format = workbook.add_format({'left': 1})
	right_format = workbook.add_format({'right': 1})
	bold_format = workbook.add_format({'bold': 0})
	merge_format = workbook.add_format({
	    'bold': 1,
	    'border': 1,
	    'align': 'center',
	    'valign': 'vcenter',
	    'fg_color': 'black',
	    'color':'white'})


	prev_mouseline = None

	# Start from the first cell. Rows and columns are zero indexed.
	row = 0
	col = 0

	columns = {
		'cage_tag': 0,
		'num_mice': 1,
		'disposition': 2,
		'cage_mouseline': 3,
		'mice_tags': 4,
		'genotypes': 5,
		'litter_sids': 6,
		'comment': 7,
		'setup_date': 8,
	}

	# Iterate over the data and write it out row by row.
	wb = open_workbook('softmousedb.xlsx')
	s = wb.sheets()[0]
	data = [s.row_values(i) for i in xrange(s.nrows)]
	labels = data[0]    # Don't sort our headers
	data = data[1:]     # Data begins on the second row
	data.sort(key=lambda x: x[columns['cage_mouseline']])
	paper_dict = {}
	paper_order = []

	import yaml

	with open("settings.yaml", 'r') as stream:
	    try:
	        settings = yaml.safe_load(stream)
	    except yaml.YAMLError as exc:
	        print(exc)

	for cage_row, cage_data in enumerate(data):
		# Each "mouseline" (e.g. p53 flox, atf4 flox) goes on its
		# own card, so this first section goes line-by-line
		# through the xlsx input, which has been sorted by mouseline,
		# and if the mouseline has changed, it jumps ahead to the next
		# sheet of paper. Otherwise, it just advances to the next
		# box within the same sheet of paper.
		mouseline = cage_data[columns['cage_mouseline']]
		print "mouseline:", mouseline
		print "Before this card, cards on current sheet:", cards_on_current_sheet
		if prev_mouseline != mouseline and prev_mouseline is not None:
			# This isn't the very first card, and we have moved
			# to a new mouseline, so advance to the next sheet of paper.
			print "-----New Mousline----"
			col = 0 # always move to leftmost card position
			# To calculate how many spreadsheet rows to advance,
			# define a vertical slot to be a row of cards (without
			# using the term row, since that refers to the spreadsheet).
			# The number of vertical slots remaining tells you how many
			# vertical slots to move forward, modulo the number of vertical
			# slots per card, because we have already advanced to the
			# next sheet in the previous iteration if we filled up a sheet.
			if cards_on_current_sheet > 0:
				print "cards on current sheet:", cards_on_current_sheet
				vertical_slots_used_on_prev_card = (cards_on_current_sheet/cards_per_row_on_sheet) + 1
				print "vert slots used on prev card:", vertical_slots_used_on_prev_card
				vertical_slots_remaining_on_prev_card = cards_per_col_on_sheet - vertical_slots_used_on_prev_card
				print "vert slots left on prev card:", vertical_slots_remaining_on_prev_card
				num_vertical_slots_to_advance = vertical_slots_remaining_on_prev_card+1

				row += rows_per_card * num_vertical_slots_to_advance
			cards_on_current_sheet = 0
			paper_dict[mouseline] = 1
			paper_order.append(mouseline)
			page_breaks.append(row)
		if prev_mouseline is None:
			# Don't need to advance to next card on 1st mouseline,
			# but still need to keep track of info re: how much paper
			# and what order paper should be loaded
			paper_dict[mouseline] = 1
			paper_order.append(mouseline)
		prev_mouseline = mouseline

		print "On Card:", paper_dict[mouseline], "Row:", cards_on_current_sheet/cards_per_row_on_sheet+1,"Col:",cards_on_current_sheet%cards_per_row_on_sheet+1
		print "spreadsheet row:", row, "col:",col

		cards_on_current_sheet += 1

		worksheet.write(row, col+0, 'PI')
		worksheet.write(row, col+1, settings["PI_name"])
		worksheet.write(row, col+3, 'Protocol: {}'.format(settings["protocol_num"]))
		worksheet.write(row+1, col+0, 'Contact')
		worksheet.write(row+1, col+1, settings["contact_name"])
		worksheet.write(row+2, col+1, settings["contact_phone"])
		worksheet.write(row+3, col+0, 'Species')
		worksheet.write(row+3, col+1, settings["species"])
		gender_row = row + 2
		male_col = col+2
		female_col = col+3
		mf_col = col+4
		worksheet.write(gender_row, male_col, 'M')
		worksheet.write(gender_row+1, male_col, 'F')
		worksheet.write(gender_row+2, male_col, 'M/F')
		worksheet.write(row+4, col+0, 'Strain')
		worksheet.write(row+5, col+0, 'Cage #')
		worksheet.write(row+7, col+0, 'Tag')
		worksheet.write(row+7, col+1, 'DOB')
		worksheet.write(row+7, col+2, 'Sex')
		# worksheet.write(row+7, col+3, 'Age')
		worksheet.write(row+7, col+3, 'Genotype')
		worksheet.write(row+3, col+1, mouseline)
		# worksheet.write(row+4, col+1, s.cell(cage_row,0).value)

		num_mice = int(cage_data[columns['num_mice']])
		mouse_data = re.split('\n',cage_data[columns['mice_tags']])
		mouse_genotypes = re.split('\n',cage_data[columns['genotypes']])
		
		num_males = 0
		num_females = 0
		for i in range(num_mice):

			# Mouse Tag Number
			mouse_num = re.split('\[',mouse_data[i])[0]
			worksheet.write(row+8+i,col+0,mouse_num)

			# Gender of particular mouse
			male = re.search('\[M',mouse_data[i])
			female = re.search('\[F',mouse_data[i])
			if male:
				num_males += 1
				worksheet.write(row+8+i,col+2,'M')
			if female:
				num_females += 1
				worksheet.write(row+8+i,col+2,'F')

			# DOB
			mouse_dob = re.search('[0-1][0-9]\-[0-3][0-9]\-20[0-9][0-9]',mouse_data[i])
			if mouse_dob:
				worksheet.write(row+8+i,col+1,mouse_dob.group())

			# # Age
			# mouse_age = re.search('[0-9]*[dw]',mouse_data[i])
			# if mouse_age:
			# 	worksheet.write(row+8+i,col+3,mouse_age.group())

			# Genotype
			worksheet.write(row+8+i,col+3,mouse_genotypes[i])

		if num_males > 0:
			if num_females > 0:
				cell = chr(65+male_col)+str(gender_row+3)
				worksheet.write(cell, 'M/F', outline_format)
				mf_range = chr(65+female_col)+str(gender_row+4)+":"+chr(65+female_col)+str(gender_row+5)
				worksheet.merge_range(mf_range, 'MATING', merge_format)
			else:
				cell = chr(65+male_col)+str(gender_row+1)
				worksheet.write(cell, 'M', outline_format)
		else:
			cell = chr(65+male_col)+str(gender_row+2)
			worksheet.write(cell, 'F', outline_format)


		if col == 0:
			col += cols_per_card
		else:
			col = 0
			row += rows_per_card
		if cards_on_current_sheet == cards_per_sheet: # Filled up a page
			cards_on_current_sheet = 0
			paper_dict[mouseline] += 1
			page_breaks.append(row)

	worksheet.print_area(0, 0, row+rows_per_card, 2*cols_per_card-1)
	total_num_pages = sum(paper_dict.values())
	worksheet.set_h_pagebreaks(page_breaks)
	worksheet.set_column(11, None, None, {'hidden': True})

	print "--------------------------------------"
	print "Load this many pages into the printer:"
	for mouseline in paper_order:
		print mouseline, paper_dict[mouseline]
	print "--------------------------------------"


	worksheet.conditional_format("A1:K1000", {'type': 'no_errors',
                                          'format': bold_format})



	workbook.close()






