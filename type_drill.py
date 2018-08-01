
from tkinter import *
from tkinter import messagebox
import textwrap


def init_app(master):
	canvas = Canvas(master)
	canvas.pack(expand=True, fill='both')
	return canvas

def read_text(file, letters_per_line):
	with open(file) as fin:
		text = fin.read()
		text = text.replace('„', '"')
		text = text.replace('“', '"')
		text = text.replace('”', '"')
		text = text.replace('–', '-')
		text = text.replace('ß', 'ss')
		text = text.replace('‚', "'")
		text = text.replace('‘', "'")

		text = text.replace('»', '"')
		text = text.replace('«', '"')

		text = text.replace('›', "'")
		text = text.replace('‹', "'")

		text = text.replace('…', "...")


		text = text.replace('Ü', 'Ue')
		text = text.replace('Ö', 'Oe')
		text = text.replace('Ä', 'Ae')

		# keep the excisting linebreaks
		text = text.split('\n')

		new_text = []
		for line in text:
			splitted = textwrap.wrap(line, width= letters_per_line)
			new_text = new_text + splitted

		new_text = [line + "¬" for line in new_text]
		return new_text

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

	g_canvas.create_rectangle(coord, fill='light green', outline='green')
	t = g_canvas.create_text(5, y, text=g_text[g_cur_line], anchor='nw', fill='black',  font=g_font)

def move_cursor():
	global g_cur_char, g_cur_line

	clear_cursor()

	g_cur_char += 1
	if g_cur_char >= len(g_text[g_cur_line]):
		g_cur_line += 1
		g_cur_char = 0
	if g_cur_line >= len(g_text):
		g_cur_line = 0
		g_cur_char = 0

	if g_cur_line >= g_scroll_after_lines:
		scroll(1)

	draw_cursor()


def keydown(event):
	global g_typed_all, g_typed_wrong
	c_soll = ord(g_text[g_cur_line][g_cur_char])



	try:
		c_ist = ord(event.char)
		if c_ist == 10 or c_ist == 13: c_ist = 172
		g_typed_all += 1
		perc = round(100*g_typed_wrong/g_typed_all,1)
		title = f"{g_file_name}: {g_typed_wrong}/{g_typed_all} ({perc}%) "

		g_master.title(title)
	except:
		return
	
	if c_ist == c_soll or c_ist==27:
		move_cursor()
	else:
		g_typed_wrong += 1

		error_text = f"Expected '{chr(c_soll)}' ({c_soll}) but received '{chr(c_ist)}' ({c_ist})"
		messagebox.showinfo('Error', error_text)
		print(error_text)
		


def scroll(n):
	print('scroll',n)
	global g_first_visible_line, g_text, g_cur_line

	g_first_visible_line += n
	if g_first_visible_line <0:
		g_first_visible_line = 0
		return

	g_text = g_full_text[g_first_visible_line:g_first_visible_line + g_visible_lines + 1]
	g_text = g_text.copy()
	draw_text()
	g_cur_line = g_cur_line - n
	draw_cursor()

def downKey(event):
    print("Down key pressed")
    scroll(1)

def upKey(event):
    print("Up key pressed")
    scroll(-1)

# global settings
g_width = 700
g_height = 400
g_font_size = 14
g_font_name = "Courier"
g_font = 0
g_letters_per_line = 60
g_cur_line = 0
g_cur_char = 0
g_canvas = 0
g_full_text = 0
g_text = 0
g_first_visible_line = 0
g_visible_lines = 30
g_scroll_after_lines = 10
g_typed_all = 0
g_typed_wrong = 0
g_master = 0
g_file_name = ''

def main():
	global g_canvas, g_full_text, g_text, g_master, g_file_name, g_font
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


	g_master.bind("<KeyPress>", keydown)
	g_master.bind('<Up>', upKey)
	g_master.bind('<Down>', downKey)
	mainloop()


main()