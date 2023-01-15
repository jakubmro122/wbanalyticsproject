from flask import Flask, render_template, request
from openpyxl import Workbook, load_workbook
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import configparser


app = Flask(__name__)

@app.route('/')
def index():

    wb = load_workbook('ZebraneDane.xlsx')
    ws = wb.active
    machineName = ws.cell
    
    return render_template('index.html', machineName=machineName)
        
@app.route('/oneDay/')
def about():

    wb = load_workbook('ZebraneDane.xlsx')
    ws = wb.active
    columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
    data = [ws.cell(row=4,column=i).value for i in range(1,6)]
    machineName = (ws['B1'].value)
    machineStoppage = data[1] - data[2]
    
    return render_template('oneDay.html', columName = columName, data=data, machineName = machineName, machineStoppage = machineStoppage)

@app.route('/oneDay/181', methods=['POST', "GET"])
def oneDay181():
    if request.method == "POST":
     wb = load_workbook('ZebraneDane.xlsx')
     wb.active = wb['181']
     ws = wb.active
     columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
     data = [ws.cell(row=4,column=i).value for i in range(1,6)]
     machineName = (ws['B1'].value)
     machineStoppage = data[1] - data[2]

    return render_template('oneDay.html', columName = columName, data=data, machineName = machineName, machineStoppage = machineStoppage) 

@app.route('/oneDay/230', methods=['POST', "GET"])
def oneDay230():
    if request.method == "POST":
     wb = load_workbook('ZebraneDane.xlsx')
     wb.active = wb['230']
     ws = wb.active
     columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
     data = [ws.cell(row=4,column=i).value for i in range(1,6)]
     machineName = (ws['B1'].value)
     machineStoppage = data[1] - data[2]
    
     return render_template('oneDay.html', columName = columName, data=data, machineName = machineName, machineStoppage = machineStoppage)

@app.route('/oneDay/254', methods=['POST', "GET"])
def oneDay254():
    if request.method == "POST":
     wb = load_workbook('ZebraneDane.xlsx')
     wb.active = wb['254']
     ws = wb.active
     columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
     data = [ws.cell(row=4,column=i).value for i in range(1,6)]
     machineName = (ws['B1'].value)
     machineStoppage = data[1] - data[2]
    
     return render_template('oneDay.html', columName = columName, data=data, machineName = machineName, machineStoppage = machineStoppage)       

@app.route('/oneDay/268', methods=['POST', "GET"])
def oneDay268():
    if request.method == "POST":
     wb = load_workbook('ZebraneDane.xlsx')
     wb.active = wb['268']
     ws = wb.active
     columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
     data = [ws.cell(row=4,column=i).value for i in range(1,6)]
     machineName = (ws['B1'].value)
     machineStoppage = data[1] - data[2]
    
     return render_template('oneDay.html', columName = columName, data=data, machineName = machineName, machineStoppage = machineStoppage)      

@app.route('/oneDay/273', methods=['POST', "GET"])
def oneDay273():
    if request.method == "POST":
     wb = load_workbook('ZebraneDane.xlsx')
     wb.active = wb['273']
     ws = wb.active
     columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
     data = [ws.cell(row=4,column=i).value for i in range(1,6)]
     machineName = (ws['B1'].value)
     machineStoppage = data[1] - data[2]
    
     return render_template('oneDay.html', columName = columName, data=data, machineName = machineName, machineStoppage = machineStoppage)  

@app.route('/oneDay/269', methods=['POST', "GET"])
def oneDay269():
    if request.method == "POST":
     wb = load_workbook('ZebraneDane.xlsx')
     wb.active = wb['269']
     ws = wb.active
     columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
     data = [ws.cell(row=4,column=i).value for i in range(1,6)]
     machineName = (ws['B1'].value)
     machineStoppage = data[1] - data[2]
    
     return render_template('oneDay.html', columName = columName, data=data, machineName = machineName, machineStoppage = machineStoppage)  



@app.route('/sevenDays/')
def seven():
 wb = load_workbook('ZebraneDane.xlsx')
 ws = wb.active
 columName = []

 columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
 sevenDaysData = []
 sevenDaysMachineStoppage = []


 for r in range(0,7):
  data = [ws.cell(row=4+r, column=i).value for i in range(1,6)]
  sevenDaysData.append(data)


  if (data[1] == None):
      sevenDaysMachineStoppage.append(0)
  else:
       machineStoppage = data[1] - data[2]
       sevenDaysMachineStoppage.append(machineStoppage)
  machineName = (ws['B1'].value)
 
 return render_template('sevenDays.html', columName = columName , data=data, machineName = machineName, machineStoppage = machineStoppage, sevenDaysData = sevenDaysData, sevenDaysMachineStoppage = sevenDaysMachineStoppage)

@app.route('/sevenDays/181', methods=['POST', "GET"])
def sevenDays181():
    if request.method == "POST":
      wb = load_workbook('ZebraneDane.xlsx')
      wb.active = wb['181']
      ws = wb.active
      columName = []

      columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
      sevenDaysData = []
      sevenDaysMachineStoppage = []

      for r in range(0,7):
       data = [ws.cell(row=4+r, column=i).value for i in range(1,6)]
       sevenDaysData.append(data)
       machineStoppage = data[1] - data[2]
       sevenDaysMachineStoppage.append(machineStoppage)
      machineName = (ws['B1'].value)
    
    return render_template('sevenDays.html', columName = columName, data=data, machineName = machineName, machineStoppage = machineStoppage, sevenDaysData = sevenDaysData, sevenDaysMachineStoppage = sevenDaysMachineStoppage) 

@app.route('/sevenDays/230', methods=['POST', "GET"])
def sevenDays230():
    if request.method == "POST":
      wb = load_workbook('ZebraneDane.xlsx')
      wb.active = wb['230']
      ws = wb.active
      columName = []

      columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
      sevenDaysData = []
      sevenDaysMachineStoppage = []

      for r in range(0,7):
       data = [ws.cell(row=4+r, column=i).value for i in range(1,6)]
       sevenDaysData.append(data)
       if (data[1] == None):
        sevenDaysMachineStoppage.append(0)
       else:
        machineStoppage = data[1] - data[2]
        sevenDaysMachineStoppage.append(machineStoppage)
      machineName = (ws['B1'].value)
    
    return render_template('sevenDays.html', columName = columName, data=data, machineName = machineName, machineStoppage = machineStoppage, sevenDaysData = sevenDaysData, sevenDaysMachineStoppage = sevenDaysMachineStoppage)


@app.route('/sevenDays/254', methods=['POST', "GET"])
def sevenDays254():
    if request.method == "POST":
      wb = load_workbook('ZebraneDane.xlsx')
      wb.active = wb['254']
      ws = wb.active
      columName = []

      columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
      sevenDaysData = []
      sevenDaysMachineStoppage = []

      for r in range(0,7):
       data = [ws.cell(row=4+r, column=i).value for i in range(1,6)]
       sevenDaysData.append(data)
       machineStoppage = data[1] - data[2]
       sevenDaysMachineStoppage.append(machineStoppage)
      machineName = (ws['B1'].value)
    
    return render_template('sevenDays.html', columName = columName, data=data, machineName = machineName, machineStoppage = machineStoppage, sevenDaysData = sevenDaysData, sevenDaysMachineStoppage = sevenDaysMachineStoppage)

@app.route('/sevenDays/268', methods=['POST', "GET"])
def sevenDays268():
    if request.method == "POST":
      wb = load_workbook('ZebraneDane.xlsx')
      wb.active = wb['268']
      ws = wb.active
      columName = []

      columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
      sevenDaysData = []
      sevenDaysMachineStoppage = []

      for r in range(0,7):
       data = [ws.cell(row=4+r, column=i).value for i in range(1,6)]
       sevenDaysData.append(data)
       machineStoppage = data[1] - data[2]
       sevenDaysMachineStoppage.append(machineStoppage)
      machineName = (ws['B1'].value)
    
    return render_template('sevenDays.html', columName = columName, data=data, machineName = machineName, machineStoppage = machineStoppage, sevenDaysData = sevenDaysData, sevenDaysMachineStoppage = sevenDaysMachineStoppage)

@app.route('/sevenDays/273', methods=['POST', "GET"])
def sevenDays273():
    if request.method == "POST":
      wb = load_workbook('ZebraneDane.xlsx')
      wb.active = wb['273']
      ws = wb.active
      columName = []

      columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
      sevenDaysData = []
      sevenDaysMachineStoppage = []

      for r in range(0,7):
       data = [ws.cell(row=4+r, column=i).value for i in range(1,6)]
       sevenDaysData.append(data)
       machineStoppage = data[1] - data[2]
       sevenDaysMachineStoppage.append(machineStoppage)
      machineName = (ws['B1'].value)
    
    return render_template('sevenDays.html', columName = columName, data=data, machineName = machineName, machineStoppage = machineStoppage, sevenDaysData = sevenDaysData, sevenDaysMachineStoppage = sevenDaysMachineStoppage)

@app.route('/sevenDays/269', methods=['POST', "GET"])
def sevenDays269():
    if request.method == "POST":
      wb = load_workbook('ZebraneDane.xlsx')
      wb.active = wb['269']
      ws = wb.active
      columName = []

      columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
      sevenDaysData = []
      sevenDaysMachineStoppage = []

      for r in range(0,7):
       data = [ws.cell(row=4+r, column=i).value for i in range(1,6)]
       sevenDaysData.append(data)
       machineStoppage = data[1] - data[2]
       sevenDaysMachineStoppage.append(machineStoppage)
      machineName = (ws['B1'].value)
    
    return render_template('sevenDays.html', columName = columName, data=data, machineName = machineName, machineStoppage = machineStoppage, sevenDaysData = sevenDaysData, sevenDaysMachineStoppage = sevenDaysMachineStoppage)

@app.route('/thirtyDays/')
def thirty():
 wb = load_workbook('ZebraneDane.xlsx')
 ws = wb.active
 columName = []

 columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
 sevenDaysData = []
 sevenDaysMachineStoppage = []

 for r in range(0,30):
  data = [ws.cell(row=4+r, column=i).value for i in range(1,6)]
  sevenDaysData.append(data)

  if (data[1] == None):
        sevenDaysMachineStoppage.append(0)
  else:
        machineStoppage = data[1] - data[2]
        sevenDaysMachineStoppage.append(machineStoppage)
  machineName = (ws['B1'].value)

 return render_template('thirtyDays.html',  columName = columName , data=data, machineName = machineName, machineStoppage = machineStoppage, sevenDaysData = sevenDaysData, sevenDaysMachineStoppage = sevenDaysMachineStoppage)


@app.route('/thirtyDays/181', methods=['POST', "GET"])
def thirtyDays181():
    if request.method == "POST":
     wb = load_workbook('ZebraneDane.xlsx')
     wb.active = wb['181']
     ws = wb.active
     columName = []

     columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
     sevenDaysData = []
     sevenDaysMachineStoppage = []

     for r in range(0,30):
      data = [ws.cell(row=4+r, column=i).value for i in range(1,6)]
      sevenDaysData.append(data)
      machineStoppage = data[1] - data[2]
      sevenDaysMachineStoppage.append(machineStoppage)

     machineName = (ws['B1'].value)

    return render_template('thirtyDays.html',  columName = columName , data=data, machineName = machineName, machineStoppage = machineStoppage, sevenDaysData = sevenDaysData, sevenDaysMachineStoppage = sevenDaysMachineStoppage)

@app.route('/thirtyDays/230', methods=['POST', "GET"])
def thirtyDays230():
    if request.method == "POST":
     wb = load_workbook('ZebraneDane.xlsx')
     wb.active = wb['230']
     ws = wb.active
     columName = []

     columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
     sevenDaysData = []
     sevenDaysMachineStoppage = []

     for r in range(0,30):
      data = [ws.cell(row=4+r, column=i).value for i in range(1,6)]
      sevenDaysData.append(data)
      if (data[1] == None):
        sevenDaysMachineStoppage.append(0)
      else:
       machineStoppage = data[1] - data[2]
       sevenDaysMachineStoppage.append(machineStoppage)
      

     machineName = (ws['B1'].value)

    return render_template('thirtyDays.html',  columName = columName , data=data, machineName = machineName, machineStoppage = machineStoppage, sevenDaysData = sevenDaysData, sevenDaysMachineStoppage = sevenDaysMachineStoppage)

@app.route('/thirtyDays/254', methods=['POST', "GET"])
def thirtyDays254():
    if request.method == "POST":
     wb = load_workbook('ZebraneDane.xlsx')
     wb.active = wb['254']
     ws = wb.active
     columName = []

     columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
     sevenDaysData = []
     sevenDaysMachineStoppage = []

     for r in range(0,30):
      data = [ws.cell(row=4+r, column=i).value for i in range(1,6)]
      sevenDaysData.append(data)
       
      machineStoppage = data[1] - data[2]
      sevenDaysMachineStoppage.append(machineStoppage)

     machineName = (ws['B1'].value)

    return render_template('thirtyDays.html',  columName = columName , data=data, machineName = machineName, machineStoppage = machineStoppage, sevenDaysData = sevenDaysData, sevenDaysMachineStoppage = sevenDaysMachineStoppage)

@app.route('/thirtyDays/268', methods=['POST', "GET"])
def thirtyDays268():
    if request.method == "POST":
     wb = load_workbook('ZebraneDane.xlsx')
     wb.active = wb['268']
     ws = wb.active
     columName = []

     columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
     sevenDaysData = []
     sevenDaysMachineStoppage = []

     for r in range(0,30):
      data = [ws.cell(row=4+r, column=i).value for i in range(1,6)]
      sevenDaysData.append(data)
      if (data[1] == None):
        sevenDaysMachineStoppage.append(0)
      else:
       machineStoppage = data[1] - data[2]
       sevenDaysMachineStoppage.append(machineStoppage)

     machineName = (ws['B1'].value)

    return render_template('thirtyDays.html',  columName = columName , data=data, machineName = machineName, machineStoppage = machineStoppage, sevenDaysData = sevenDaysData, sevenDaysMachineStoppage = sevenDaysMachineStoppage)

@app.route('/thirtyDays/273', methods=['POST', "GET"])
def thirtyDays273():
    if request.method == "POST":
     wb = load_workbook('ZebraneDane.xlsx')
     wb.active = wb['273']
     ws = wb.active
     columName = []

     columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
     sevenDaysData = []
     sevenDaysMachineStoppage = []

     for r in range(0,30):
      data = [ws.cell(row=4+r, column=i).value for i in range(1,6)]
      sevenDaysData.append(data)
      machineStoppage = data[1] - data[2]
      sevenDaysMachineStoppage.append(machineStoppage)

     machineName = (ws['B1'].value)

    return render_template('thirtyDays.html',  columName = columName , data=data, machineName = machineName, machineStoppage = machineStoppage, sevenDaysData = sevenDaysData, sevenDaysMachineStoppage = sevenDaysMachineStoppage)

@app.route('/thirtyDays/269', methods=['POST', "GET"])
def thirtyDays269():
    if request.method == "POST":
     wb = load_workbook('ZebraneDane.xlsx')
     wb.active = wb['269']
     ws = wb.active
     columName = []

     columName = [ws.cell(row=3,column=i).value for i in range(1,6)]
     sevenDaysData = []
     sevenDaysMachineStoppage = []

     for r in range(0,30):
      data = [ws.cell(row=4+r, column=i).value for i in range(1,6)]
      sevenDaysData.append(data)
      machineStoppage = data[1] - data[2]
      sevenDaysMachineStoppage.append(machineStoppage)

     machineName = (ws['B1'].value)

    return render_template('thirtyDays.html',  columName = columName , data=data, machineName = machineName, machineStoppage = machineStoppage, sevenDaysData = sevenDaysData, sevenDaysMachineStoppage = sevenDaysMachineStoppage)

if __name__ =='__main__':
    app.run(host="0.0.0.0", port=5000)


