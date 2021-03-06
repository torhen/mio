import os
from tkinter import *
from tkinter import messagebox
import textwrap
import random

def init_app(master):
	canvas = Canvas(master)
	canvas.pack(expand=True, fill='both')
	return canvas

def special_chars():
	""" every char only once and in scrambled order"""
	s = ""
	while True:
		i = random.randint(0, len(g_special_chars)-1)
		s = s + g_special_chars[i]
		s = ''.join(set(s))
		if len(s) >= g_special_chars_count:
			l = list(s)
			random.shuffle(l)
			return ''.join(l) + chr(g_placeholder_tab)


def read_text(file, letters_per_line):
	with open(file, encoding=g_text_encoding) as fin:
		text = fin.read()
		text = text.replace('„', '"')
		text = text.replace('“', '"')
		text = text.replace('”', '"')
		text = text.replace('–', '-')
		text = text.replace('—','-')
		text = text.replace('ß', 'ss')
		text = text.replace('‚', "'")
		text = text.replace('‘', "'")
		text = text.replace('»', '"')
		text = text.replace('«', '"')
		text = text.replace('›', "'")
		text = text.replace('‹', "'")
		text = text.replace('…', "...")
		#text = text.replace('Ü', 'Ue')
		#text = text.replace('Ö', 'Oe')
		#text = text.replace('Ä', 'Ae')

		text2 = ''
		for c in text:
			if ord(c)!=173:  # line break 'minus'
				text2 += c
		text = text2

		# keep the excisting linebreaks
		text = text.split('\n')

		new_text = []
		for line in text:
			splitted = textwrap.wrap(line, width= letters_per_line)
			new_text = new_text + splitted

		text3 = []
		for i, line in enumerate(new_text):
			#s = str(i) + '⋅'  + special_chars() + line + "¬"
			s = str(i) + chr(g_placeholder_tab)  + special_chars() + line + "¬"

			text3.append(s)

		return text3

def draw_text():
	g_canvas.delete("all")
	g_canvas.create_rectangle((0,0,g_width, g_height), fill='white', outline='white')
	for line, line_text in enumerate(g_text):
		t = g_canvas.create_text(5, 5 + 1.5 * g_font_size * line, text = line_text, anchor='nw', fill='black', font=g_font)


def clear_cursor():
	y = 5 + 1.5 * g_font_size * g_cur_line
	t = g_canvas.create_text(5, y, text=g_text[g_cur_line], anchor='nw', fill='black',  font=g_font)
	g_canvas.create_rectangle(g_canvas.bbox(t), fill='white', outline='white')
	t = g_canvas.create_text(5, y, text=g_text[g_cur_line], anchor='nw', fill='black',  font=g_font)

def draw_cursor():
	y = 5 + 1.5 * g_font_size * g_cur_line

	# just calulate the coordinates of the cursor

	if g_cur_line < 0 or g_cur_line > g_visible_lines:
		# cusor not visible
		return

	t = g_canvas.create_text(5, y, text=g_text[g_cur_line][0:g_cur_char], anchor='nw', fill='white',  font=g_font)
	x0, y0, w0, h0 = g_canvas.bbox(t)
	t = g_canvas.create_text(5, y, text=g_text[g_cur_line][0:g_cur_char+1], anchor='nw', fill='white',  font=g_font)
	x1, y1, w1, h1 = g_canvas.bbox(t)

	coord0 = (x0,  y0,  w0 , h0)
	coord1 = (x1,  y1,  w1 , h1)
	coord = (w0,   y0, w1, h0)

	if g_system_state == 0:
		cursor_color_fill = 'light green'
		cursor_color_outline = 'green'
	else:
		cursor_color_fill = '#FAAFBA'
		cursor_color_outline = 'red'

	g_canvas.create_rectangle(coord, fill=cursor_color_fill, outline=cursor_color_outline)
	t = g_canvas.create_text(5, y, text=g_text[g_cur_line], anchor='nw', fill='black',  font=g_font)

def move_cursor(n):
	global g_cur_char, g_cur_line

	clear_cursor()

	g_cur_char += n

	if g_cur_char >= len(g_text[g_cur_line]):
		g_cur_line += n
		g_cur_char = 0

	if g_cur_char < 0:
		g_cur_char = 0

	if g_cur_line < 0:
		g_cur_line = 0

	if g_cur_line >= len(g_text):
		g_cur_line = 0
		g_cur_char = 0

	if g_cur_line >= g_scroll_after_lines:
		scroll(n)

	draw_cursor()


def keydown(event):
	global g_typed_all, g_typed_wrong, g_system_state
	c_soll = ord(g_text[g_cur_line][g_cur_char])

	if event.char == '§':
		print('PANIC!')
		g_master.wm_state('iconic')
		return

	try:
		c_ist = ord(event.char)
		#print('Key pressed ord=', c_ist)


		if c_ist == 10 or c_ist == 13: 
			c_ist = 172

		# tabulator key
		if c_ist == 9: c_ist = 8901

		# Ctrl + C
		if c_ist == 3: c_ist = 9426	

		# Ctrl + V
		if c_ist == 22: c_ist = 9445	


		# Ctrl + X
		if c_ist == 24: c_ist = 9447			

		g_typed_all += 1
		perc = round(100*g_typed_wrong/g_typed_all,3)

	except:
		return

	if g_system_state == 0:
		# normal state
	
		if c_ist == c_soll or c_ist==27:
			g_system_state = 0
			move_cursor(1)
			title = f"{g_file_name}: {g_typed_wrong}/{g_typed_all} ({perc:.2f}%) "
			g_master.title(title)
		else:
			g_typed_wrong += 1
			g_system_state = 1

			move_cursor(0)

			error_text = f"Expected '{chr(c_soll)}' ({c_soll}) but received '{chr(c_ist)}' ({c_ist})"
			g_master.title(error_text)

			print(error_text)
	else:
		# system is in error state
		if c_ist == 8: # Back delete key
			g_system_state = 0
			move_cursor(0)
		else:
			error_text = f"Expected '{chr(c_soll)}' ({c_soll}) but received '{chr(c_ist)}' ({c_ist})"
			g_master.title(error_text)
		


def scroll(n):
	global g_first_visible_line, g_text, g_cur_line, g_last_position

	g_first_visible_line += n
	if g_first_visible_line <0:
		g_first_visible_line = 0
		return

	g_text = g_full_text[g_first_visible_line:g_first_visible_line + g_visible_lines + 1]
	g_text = g_text.copy()
	draw_text()
	g_cur_line = g_cur_line - n
	draw_cursor()
	g_last_position = g_last_position +n 

def downKey(event):
	scroll_line(1)

def scroll_line(n):
	global g_cur_line, g_last_position
	g_cur_line = g_cur_line + n
	scroll(n)


def upKey(event):
	global g_cur_line
	g_cur_line = g_cur_line -1
	scroll(-1)

def rightKey(event):
	move_cursor(1)

def leftKey(event):
	move_cursor(-1)



# global settings
g_width = 700
g_height = 400
g_font_size = 18
g_font_name = "Courier"
g_font = 0
g_letters_per_line = 80
g_cur_line = 0
g_cur_char = 0
g_canvas = 0
g_full_text = 0
g_text = 0
g_first_visible_line = 0
g_visible_lines = 100
g_scroll_after_lines = 10
g_typed_all = 0
g_typed_wrong = 0
g_master = 0
g_file_name = ''
g_special_chars = r"{}*#%&/[]+@_$|\<>=^~ⓒⓥⓧⓒⓥⓧⓒⓥⓧⓒⓥⓧⓒⓥⓧ"
g_special_chars_count = 3
g_text_encoding = 'utf-8'
g_last_position = 0
g_placeholder_tab = 8901
g_system_state = 0

def end_app():
	print('write last position',g_last_position)
	with open('last_position.txt', 'w') as fout:
		fout.write(str(g_last_position))
	g_master.quit()


def main():
	global g_canvas, g_full_text, g_text, g_master, g_file_name, g_font

	if os.path.isfile('last_position.txt'):
		with open('last_position.txt') as fin:
			s = fin.read()
			g_last_position = int(s) 
	else:
		g_last_position = 0


	g_master = Tk()
	g_master.geometry(f'{g_width}x{g_height}')
	g_canvas = init_app(g_master)

	g_font = g_font_name + " " + str(g_font_size)

	g_file_name = 'text.txt'
	g_full_text = read_text(g_file_name, g_letters_per_line)

	g_master.title(g_file_name)
	g_text = g_full_text[g_first_visible_line:g_first_visible_line + g_visible_lines + 1]

	draw_text()

	draw_cursor()

	scroll_line(g_last_position)


	g_master.bind("<KeyPress>", keydown)
	g_master.bind('<Up>', upKey)
	g_master.bind('<Down>', downKey)
	g_master.bind('<Right>', rightKey)
	g_master.bind('<Left>', leftKey)

	g_master.protocol("WM_DELETE_WINDOW", end_app)
	mainloop()


main()