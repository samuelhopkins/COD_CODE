from wx import *
import config
import cal
import sys
import os



months=["January","February","March","April","May","June","July"
		,"August","September","October","November","December"]

wildcard_1= "Excel Files (*.xlsx)|*.xlsx" 
		   

wildcard_2= "ICS File (*.ics)|*.ics"  




class mainFrame(Frame):
	def __init__(self,parent,id):
		Frame.__init__(self,parent,id,'COD Scheduler 3000',size=(400,250))

		self.InitUI()
		self.Centre()

	def InitUI(self):
		self.panel_1=Panel(self)

		#initial dialogs
		monthBox = SingleChoiceDialog(None,"What is the first month you need scheduled?","Initialize Month",months)

		if monthBox.ShowModal()==ID_OK:
			config.month=int(months.index(monthBox.GetStringSelection())+1)
	
										

		yearBox = TextEntryDialog(None,"Enter the year","Initialize year","")
		if yearBox.ShowModal()==ID_OK:
			config.year=int(yearBox.GetValue())

		shiftBox = TextEntryDialog(None,"Enter the total number of shifts","Initialize shifts","")
		if shiftBox.ShowModal()==ID_OK:
			config.total_Shifts=int(shiftBox.GetValue())
		#add static text to panel
		text_list=StaticText(self.panel_1,-1,"Please upload the list of counselors below")
		text_list.SetFont(Font(14,SWISS,NORMAL,BOLD))
		text_list.SetSize(text_list.GetBestSize())
		text_out=StaticText(self.panel_1,-1,"Please upload the travel calendar below")
		text_out.SetFont(Font(14,SWISS,NORMAL,BOLD))
		text_out.SetSize(text_out.GetBestSize())

		#bind buttons to handlers
		button=Button(self.panel_1,label="Browse")
		button2=Button(self.panel_1,label="Browse")
		button3=Button(self.panel_1,label="Run")
		button4=Button(self.panel_1,label="Exit")
		self.Bind(EVT_BUTTON,self.openButton,button)
		self.Bind(EVT_BUTTON,self.openButton_2,button2)
		self.Bind(EVT_BUTTON,self.run,button3)
		self.Bind(EVT_BUTTON,self.close,button4)
		self.Bind(EVT_CLOSE,self.closewindow)

		#sizer to layout controls
		self.sizer=BoxSizer(VERTICAL)
		self.sizer.Add(text_list,0,ALL,10)
		self.sizer.Add(button,0,ALL,10)
		self.sizer.Add(text_out,0,ALL,10)
		self.sizer.Add(button2,0,ALL,10)
		self.sizer.Add(button3,0,ALL,10)
		self.sizer.Add(button4,0,ALL,10)
		self.panel_1.SetSizer(self.sizer)
		self.panel_1.Layout()



	def run(self,evt):
		dlg=MessageDialog(self.panel_1,"Do you want a COD schedule?","Final Answer", wx.YES_NO | wx.ICON_QUESTION)
		if dlg.ShowModal()==ID_YES:
			cal.COD_Scheduler_3000(config.path_1,config.path_2,config.month,config.year,config.total_Shifts)
			dlg.Destroy()
		

	def close(self,evt):
		dlg=MessageDialog(self.panel_1,"Do you wish to exit?","Exit Program", wx.YES_NO | wx.ICON_QUESTION)
		if dlg.ShowModal()==ID_YES:
			Exit()

	def openButton(self,evt):
		dlg=FileDialog(self,message="Choose a file",
			defaultDir=os.getcwd(),defaultFile="",
			wildcard=wildcard_1,style= OPEN | MULTIPLE |
			CHANGE_DIR)

		if dlg.ShowModal() == ID_OK:
			path=dlg.GetPaths()[0]
			config.path_1=path
			output=StaticText(self.panel_1,-1,config.path_1,pos=(10,75))
			output.SetFont(Font(10,SWISS,NORMAL,NORMAL))
			output.SetSize(output.GetBestSize())

	def openButton_2(self,evt):
		dlg=FileDialog(self,message="Choose a file",
			defaultDir=os.getcwd(),defaultFile="",
			wildcard=wildcard_2,style= OPEN | MULTIPLE |
			CHANGE_DIR)

		if dlg.ShowModal() == ID_OK:
			path=dlg.GetPaths()[0]
			config.path_2=path
			output=StaticText(self.panel_1,-1,config.path_2,pos=(10,150))
			output.SetFont(Font(10,SWISS,NORMAL,NORMAL))
			output.SetSize(output.GetBestSize())


	def closewindow(self, evt):
		self.Destroy()



if __name__ == "__main__":
	app=PySimpleApp()
	frame=mainFrame(parent=None,id=-1)
	frame.Show()
	app.MainLoop()

