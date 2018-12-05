from tkinter import *
from tkinter.filedialog import *
from tkinter.ttk import Button, Progressbar
from threading import Thread
import CalcPoints as cp
# import openpyxl

def selfile():
	filePath = askopenfilename(filetypes = (("Excel files","*.xlsx"),))
	if filePath == "":
		pass
	elif filePath[-5:] != '.xlsx':
		instr.set("Please select an excel (.xlsx) file")
		infstr.set("")
	else:
		try:
			infstr.set(filePath)
		except Exception as e:
			infstr.set(e)

def go():
	t = Thread(target = cp.calculate, args = (calcing, wbl, instr, infstr, progress))
	t.start()

def displayRes():
	x = optionbox.get()
	try:
		cp.showRes(wbl[0], int(x), rankstr, partistr, pointstr)
	except:
		infstr.set("Please enter an integer")

def save():
	savePath = asksaveasfilename(filetypes = (("Excel files","*.xlsx"),), defaultextension = ".xlsx")
	if savePath[-5:] == ".xlsx":
		try:
			wbl[0].save(savePath)
			infstr.set("Saved. In registrations, column G is the points. In event sheets, column U is.")
		except:
			infstr.set("Unable to save (make sure the selected save file is not open elsewhere)")
	else:
		infstr.set("Please select excel file (.xlsx)")

def mtChecker():
	while running:
		if progress.get() == 0 or progress.get() == 100:
			pb.grid_forget()
			instrLbl.grid(row = 0, column = 0, padx = 20, pady = 10)
		else:
			instrLbl.grid_forget()
			pb.grid(row = 0, column = 0, pady = 10, sticky = "ew", padx = 50)
		if infstr.get() == "":
			infLbl.grid_forget()
			runBtn.grid_forget()
		else:
			infLbl.grid(row = 1, column = 0)
			if calcing.get() == 0 and infstr.get()[-5:] == ".xlsx":
				runBtn.grid(row = 3, column = 0, pady = (0,10))
			else:
				runBtn.grid_forget()
		if calcing.get() == 0:
			selBtn.grid(row = 2, column = 0, pady = 10)
		else:
			selBtn.grid_forget()
		if rankstr.get() == "":
			resultFrm.grid_forget()
		else:
			resultFrm.grid(row = 5, column = 0)
		if progress.get() == 100:
			optionFrm.grid(row = 4, column = 0)
		else:
			optionFrm.grid_forget()

if __name__ == "__main__":
	wbl = []
	running = True
	root = Tk()
	calcing = IntVar(value = 0)
	root.configure(bg = "#FFFFFF")

	# Instructions
	instr = StringVar(value = 'Select the cubecomps results export of the competition')
	instrLbl = Label(root,
					textvariable = instr,
					font = "defaultwindowsfont 15",
					background = "#FFFFFF")
	# Additional info (selected file name etc)
	infstr = StringVar()
	infLbl = Label(root,
				textvariable = infstr,
				background = "#FFFFFF")
	# Averaging progress
	progress = DoubleVar()
	progress.set(0)
	pb = Progressbar(root,
					orient = "horizontal",
					mode = "determinate",
					variable = progress)
	# Select export
	selBtn = Button(root,
					text = "Select file",
					command = lambda: selfile())
	# Run the points calculation module
	runBtn = Button(root,
					text = "Calculate points",
					command = go)
	# Result options
	optionFrm = Frame(root,
					background = "#FFFFFF")
	optionLbl = Label(optionFrm,
					text = "How many people are being awarded?",
					background = "#FFFFFF")
	optionLbl.grid(row = 0, column = 0)
	optionbox = Entry(optionFrm)
	optionbox.grid(row = 1, column = 0)
	optionBtn = Button(optionFrm,
					text = "Submit",
					command = displayRes)
	optionBtn.grid(row = 2, column = 0)
	orLbl = Label(optionFrm,
				text = "OR",
				background = "#FFFFFF")
	orLbl.grid(row = 3, column = 0)
	saveBtn = Button(optionFrm,
					text = "Save as excel file with details",
					command = save)
	saveBtn.grid(row = 4, column = 0, pady = (0,10))
	# Results display
	resultFrm = Frame(root,
					background = "#FFFFFF")
	rankHead = Label(resultFrm,
					text = "Rank",
					font = "Helvetica",
					background = "#FFFFFF")
	rankHead.grid(row = 0, column = 0)
	participantHead = Label(resultFrm,
					text = "Participant",
					font = "Helvetica",
					background = "#FFFFFF")
	participantHead.grid(row = 0, column = 1)
	pointsHead = Label(resultFrm,
					text = "Points (max 100)",
					font = "Helvetica",
					background = "#FFFFFF")
	pointsHead.grid(row = 0, column = 2)
	rankstr = StringVar()
	rankBoard = Message(resultFrm,
						textvariable = rankstr,
						background = "#FFFFFF")
	rankBoard.grid(row = 1, column = 0)
	partistr = StringVar()
	partiBoard = Message(resultFrm,
						textvariable = partistr,
						background = "#FFFFFF")
	partiBoard.grid(row = 1, column = 1)
	pointstr = StringVar()
	pointBoard = Message(resultFrm,
						textvariable = pointstr,
						background = "#FFFFFF")
	pointBoard.grid(row = 1, column = 2)
	# Check if any widgets are empty and update UI accordingly
	mtCheck = Thread(target = mtChecker)
	mtCheck.start()

	root.mainloop()
	running = False