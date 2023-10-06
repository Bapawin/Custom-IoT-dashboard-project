from win32api import RGB
import win32com.client
import os
import PySimpleGUI as sg

#Function inch to Pixel
def inch(A):
    por = A*72.4
    return por
#Function Color R G B
def colorcode(R,G,B):
    por = R + G*(256) + B*(256*256)
    return por
#Function change Size in textbox
def Sizetext(Target,Size):
    Target.TextFrame.TextRange.font.size = Size
    return
#Function position and size shapes
def PS(Target,Left,Top,Width,Height):
    Target.left = Left
    Target.top  = Top
    Target.width = Width
    Target.height = Height
    return
#Function Copy and Paste Graph
def CopyPaste(Worksheet,Chart,Target,type):
    excel_worksheet = xlWorkbook.Worksheets(Worksheet)
    Chrt = excel_worksheet.ChartObjects(Chart)
    Chrt.Copy()
    Target.Shapes.PasteSpecial(DataType=type)
    A = Target.Shapes(Target.Shapes.Count)
    return A
#Function Copy and Paste Table
def CPtable(sheet,range,target,type) :
    excel_worksheet = xlWorkbook.Worksheets(sheet)
    excel_range = excel_worksheet.Range(range) # Range of 2 tables
    excel_range.Copy()
    target.Shapes.PasteSpecial(DataType=type)
    A = target.Shapes(powerpoint_slide.Shapes.Count)
    return A
#Function Color text
def colortext(Target,R,G,B):
    Target.TextFrame.TextRange.font.color.rgb = colorcode(R,G,B)
    return
def Texttitle(target,text):
    target.TextFrame.TextRange.Text = text
    return
def Fillcolor(target,R,G,B):
    target.fill.forecolor.rgb = colorcode(R,G,B)
    return
def Texttable(Target,Row,Column,text):
    Target.Table.Cell(Row,Column).Shape.TextFrame.TextRange.text = text
    return
#Function
def addLine(target,weight,R,G,B):
    target.line.weight = weight
    target.line.forecolor.rgb = RGB(R,G,B)
    return

def updateTarget(period_value):
    if period_value == 1:
        target_value = [1]
    elif period_value == 2:
        target_value = [1, 2]
    else:
        target_value = []
    window['-Target-'].update(values=target_value)

sg.theme("Reddit")

layout = [[sg.Text("")],
        [sg.Text("       Data Report Generator", font=("Arial Rounded MT Bold",30))],
        [sg.Text("")],
        [sg.Text("Import OEE Trend file:"),
         sg.InputText(default_text = 'choose an excel file...',key = '-TREND-', disabled=True, size=52),
         sg.FileBrowse(file_types=(('excel', '*.xlsx'),('all types','*.*'),), size=8, initial_folder="C:/Users")],
        [sg.Text("Import OEE Loss analysis file:"),
         sg.InputText(default_text= 'choose an excel file...',key='-LOSS-', disabled=True),
         sg.FileBrowse(file_types=(('macro-excel', '*.xlsm'),('all types','*.*'),), size=8, initial_folder="C:/Users")],
        [sg.Text("Month Period: "), 
         sg.Combo(values=[1,2], default_value='', key="-Period-", readonly=True, enable_events=True, size=3),
         sg.Text("Month Target: "), 
         sg.Combo(values=[], default_value='', key="-Target-", readonly=True, size=3)],
        [sg.Text("")],
        [sg.Button("Generate",size=(16,2), pad=((230,230), 0))],[sg.Exit("Exit"),sg.Push(),sg.Image("densologo.png")]]

window = sg.Window("Denso Data Report Generator", layout, icon='dogicon.ico', resizable=False)

while True:
    event, values = window.read()
    print(event, values)
    if event in(sg.WINDOW_CLOSED, "Exit"):
        break

    if event == "-Period-":
        updateTarget(values['-Period-'])

    if event == "Generate":

        trendfile = values['-TREND-']
        lossfile = values['-LOSS-']
        
        if trendfile == 'choose an excel file...':
            sg.popup_error("Please choose OEE Trend file !!! ")
            continue
        if lossfile == 'choose an excel file...':
            sg.popup_error("Please choose OEE Loss analysis file !!! ")
            continue
        if values['-Period-'] == 1:
            month = 1
        else: 
            month = 2
        if values['-Target-'] == 1:
            Target = 1      
        else: 
            Target = 2
 
        window.close()     


        #ExcelApp = win32com.client.GetActiveObject("Excel.Application")
        ExcelApp = win32com.client.Dispatch("Excel.Application")
        ExcelApp.Visible = True

        # Grab the workbook with the charts.
        xlWorkbook = ExcelApp.Workbooks.Open(trendfile)
        xlWorkbook2 = ExcelApp.Workbooks.Open(lossfile)

        # Create a new instance of PowerPoint and make sure it's visible.
        PPTApp = win32com.client.Dispatch("PowerPoint.Application")
        PPTApp.Visible = True

        # Add a presentation to the PowerPoint Application, returns a Presentation Object. 
        powerpoint_slide = os.getcwd() + "\Data Report Format.pptx"
        PPTPresentation = PPTApp.Presentations.Open(r''+powerpoint_slide)

        #First Page name
        page = 0
        powerpoint_slide = PPTPresentation.Slides[page]
        textbelow = powerpoint_slide.Shapes.AddTextbox(1,inch(0.37),inch(2.2),inch(7.8),inch(2.19))
        Texttitle(textbelow,"[" + xlWorkbook2.Worksheets("Calculation Sheet").Range("B1").text + "]" + "\nProductivity Data Report\n" + "(" + xlWorkbook2.Worksheets("Calculation Sheet").Range("I2").text + " - " + xlWorkbook2.Worksheets("Calculation Sheet").Range("I3").text + " Analysis" + ")")
        Sizetext(textbelow,26)
        textbelow.TextFrame.TextRange.font.bold = True

        title = powerpoint_slide.Shapes.AddTextbox(1,inch(0.37),inch(4.54),inch(5.04),inch(0.4))
        Texttitle(title,xlWorkbook2.Worksheets("Calculation Sheet").Range("D1").text)
        Sizetext(title,24)

        '''Part1'''
        LineName = xlWorkbook2.Worksheets("Analysis Table").Range("BH2").text
        page = 4 
        month = int(month)
        Target = int(Target)
        '''powerpoint_slide = PPTPresentation.Slides.Add(page,12)''' # Position of Slide, Layout
        #Title

        if month < 1.5 :
            DNS = xlWorkbook.Worksheets("Transfer to PPT").Range("X4").text
            DS =  xlWorkbook.Worksheets("Transfer to PPT").Range("AA4").text
            NS =  xlWorkbook.Worksheets("Transfer to PPT").Range("AD4").text
            M1 = xlWorkbook.Worksheets("Transfer to PPT").Range("L4").text
            M1text = xlWorkbook.Worksheets("Transfer to PPT").Range("L3").text
            if DNS > DS:
                if NS > DNS:MaxtoMin = DS + " - " + NS
                if DNS > NS:
                    if NS > DS: MaxtoMin = DS + " - " + DNS
                    else: MaxtoMin = NS + " - " + DNS
            else:
                if NS > DS : MaxtoMin = DNS + " - " + NS
                if DS > NS : 
                    if NS > DNS: MaxtoMin = DNS + " - " + DS
                    else: MaxtoMin = NS + " - " + DS       
        else:
            MaxtoMin = "..."
            M1text = xlWorkbook.Worksheets("Transfer to PPT").Range("L3").text
            M2text = xlWorkbook.Worksheets("Transfer to PPT").Range("M3").text
            M1 = xlWorkbook.Worksheets("Transfer to PPT").Range("L4").text
            M2 = xlWorkbook.Worksheets("Transfer to PPT").Range("M4").text
            if M1 > M2 : MtoM = M2 + " - " + M1
            else: MtoM = M1 + " - " + M2
            FM1 = xlWorkbook.Worksheets("Transfer to PPT").Range("L7").text
            FM2 = xlWorkbook.Worksheets("Transfer to PPT").Range("M7").text
            if FM1 > FM2 :
                Ftext = "The fluctuation of "+M1text+" (σ = "+FM1+") was worse than "+M2text+" (σ = "+FM2+")."
            else : Ftext = "The fluctuation of "+M2text+" (σ = "+FM2+") was worse than "+M1text+" (σ = "+FM1+")."



        '''Part1/2-4''' 
        if month < 1.5 :
            i = 0
            monthtext = xlWorkbook.Worksheets("Transfer to PPT").Range("X3").text
            while i < 3 : 
                powerpoint_slide = PPTPresentation.Slides.Add(page,12) 
                title = powerpoint_slide.Shapes.AddTextbox(1, 0, 0, 724, 100) 
                if i == 0 :
                    table = CPtable("Transfer to PPT","A11:J17",powerpoint_slide,1)
                    graph = CopyPaste("OEE_Daily_Month1 DS+NS","KOPCHART",powerpoint_slide,1)
                    title.TextFrame.TextRange.Text = 'Part 1: Daily OEE Trend ('+monthtext+', Dayshift + Nightshift)'
                    AVGOEE = xlWorkbook.Worksheets("Transfer to PPT").Range("C14").text
                    FLUOEE = xlWorkbook.Worksheets("Transfer to PPT").Range("C17").text
                    SDS = "(Dayshift + Nightshift)"
                if i == 1 :
                    table = CPtable("Transfer to PPT","A18:J24",powerpoint_slide,1)
                    graph = CopyPaste("OEE_Daily_Month1 DS","KOPCHART",powerpoint_slide,1)
                    title.TextFrame.TextRange.Text = 'Part 1: Daily OEE Trend ('+monthtext+', Dayshift)'
                    AVGOEE = xlWorkbook.Worksheets("Transfer to PPT").Range("C21").text
                    FLUOEE = xlWorkbook.Worksheets("Transfer to PPT").Range("C24").text
                    SDS = "(Dayshift)"
                if i == 2 :
                    table = CPtable("Transfer to PPT","A25:J31",powerpoint_slide,1)
                    graph = CopyPaste("OEE_Daily_Month1 NS","KOPCHART",powerpoint_slide,1)
                    title.TextFrame.TextRange.Text = 'Part 1: Daily OEE Trend ('+monthtext+', Nightshift)'
                    AVGOEE = xlWorkbook.Worksheets("Transfer to PPT").Range("C28").text
                    FLUOEE = xlWorkbook.Worksheets("Transfer to PPT").Range("C31").text
                    SDS = "(Nightshift)"
                table.LockAspectRatio = False
                graph.LockAspectRatio = False
                PS(table,inch(0.1),inch(4.02),inch(9.79),inch(2))   
                Sizetext(title,20)
                title.TextFrame.TextRange.font.bold = True
                title.Left = 15
                title.Top = 5
                PS(graph,10,45,inch(9.7),inch(3.3))
                
                textbelow = powerpoint_slide.Shapes.AddTextbox(1, inch(0.1), inch(6.15), inch(9.85),inch(0.57))
                textbelow.TextFrame.TextRange.Text = 'The average OEE of '+monthtext+' '+SDS+' was '+AVGOEE+', and the OEE fluctuation was '+FLUOEE+'.'
                colortext(textbelow,0,0,255)
                Sizetext(textbelow,14)
                textbelow.fill.forecolor.rgb = colorcode(255,255,0)
                textbelow.TextFrame.TextRange.font.bold = True

                i = i+1
                page = page+1
        else:
            i = 0
            while i < 6 :
                powerpoint_slide = PPTPresentation.Slides.Add(page,12)
                title = powerpoint_slide.Shapes.AddTextbox(1, 0, 0, 724, 100)  
                if i == 0 :
                    table = CPtable("Transfer to PPT","A11:J17",powerpoint_slide,1)
                    graph = CopyPaste("OEE_Daily_Month1 DS+NS","KOPCHART",powerpoint_slide,1)
                    title.TextFrame.TextRange.Text = 'Part 1: Daily OEE Trend ('+M1text+', Dayshift + Nightshift)'
                    monthtext = M1text
                    AVGOEE = xlWorkbook.Worksheets("Transfer to PPT").Range("C14").text
                    FLUOEE = xlWorkbook.Worksheets("Transfer to PPT").Range("C17").text
                    SDS = "(Dayshift + Nightshift)"
                if i == 1 :
                    table = CPtable("Transfer to PPT","A18:J24",powerpoint_slide,1)
                    graph = CopyPaste("OEE_Daily_Month1 DS","KOPCHART",powerpoint_slide,1)
                    title.TextFrame.TextRange.Text = 'Part 1: Daily OEE Trend ('+M1text+', Dayshift)'
                    monthtext = M1text
                    AVGOEE = xlWorkbook.Worksheets("Transfer to PPT").Range("C21").text
                    FLUOEE = xlWorkbook.Worksheets("Transfer to PPT").Range("C24").text
                    SDS = "(Dayshift)"
                if i == 2 :
                    table = CPtable("Transfer to PPT","A25:J31",powerpoint_slide,1)
                    graph = CopyPaste("OEE_Daily_Month1 NS","KOPCHART",powerpoint_slide,1)
                    title.TextFrame.TextRange.Text = 'Part 1: Daily OEE Trend ('+M1text+', Nightshift)'
                    monthtext = M1text
                    AVGOEE = xlWorkbook.Worksheets("Transfer to PPT").Range("C28").text
                    FLUOEE = xlWorkbook.Worksheets("Transfer to PPT").Range("C31").text
                    SDS = "(Nightshift)"
                if i == 3 :
                    table = CPtable("Transfer to PPT","A33:J39",powerpoint_slide,1)
                    graph = CopyPaste("OEE_Daily_Month2 DS+NS","KOPCHART",powerpoint_slide,1)
                    title.TextFrame.TextRange.Text = 'Part 1: Daily OEE Trend ('+M2text+', Dayshift + Nightshift)'
                    monthtext = M2text
                    AVGOEE = xlWorkbook.Worksheets("Transfer to PPT").Range("C36").text
                    FLUOEE = xlWorkbook.Worksheets("Transfer to PPT").Range("C39").text
                    SDS = "(Dayshift + Nightshift)"
                if i == 4 :
                    table = CPtable("Transfer to PPT","A40:J46",powerpoint_slide,1)
                    graph = CopyPaste("OEE_Daily_Month2 DS","KOPCHART",powerpoint_slide,1)
                    title.TextFrame.TextRange.Text = 'Part 1: Daily OEE Trend ('+M2text+', Dayshift)'
                    monthtext = M2text
                    AVGOEE = xlWorkbook.Worksheets("Transfer to PPT").Range("C43").text
                    FLUOEE = xlWorkbook.Worksheets("Transfer to PPT").Range("C46").text
                    SDS = "(Dayshift)"
                if i == 5 :
                    table = CPtable("Transfer to PPT","A47:J53",powerpoint_slide,1)
                    graph = CopyPaste("OEE_Daily_Month2 NS","KOPCHART",powerpoint_slide,1)
                    title.TextFrame.TextRange.Text = 'Part 1: Daily OEE Trend ('+M2text+', Nightshift)'
                    monthtext = M2text
                    AVGOEE = xlWorkbook.Worksheets("Transfer to PPT").Range("C50").text
                    FLUOEE = xlWorkbook.Worksheets("Transfer to PPT").Range("C53").text
                    SDS = "(Nightshift)"
                table.LockAspectRatio = False
                graph.LockAspectRatio = False
                PS(table,inch(0.1),inch(4.02),inch(9.79),inch(2))
                Sizetext(title,20)
                title.TextFrame.TextRange.font.bold = True
                title.Left = 15
                title.Top = 5
                PS(graph,10,45,inch(9.7),inch(3.3))

                textbelow = powerpoint_slide.Shapes.AddTextbox(1, inch(0.1), inch(6.15), inch(9.85),inch(0.57))
                textbelow.TextFrame.TextRange.Text = 'The average OEE of '+monthtext+' '+SDS +' was '+AVGOEE+', and the OEE fluctuation was '+FLUOEE+'.'
                colortext(textbelow,0,0,255)
                Sizetext(textbelow,14)
                textbelow.fill.forecolor.rgb = colorcode(255,255,0)
                textbelow.TextFrame.TextRange.font.bold = True

                i = i+1
                page = page+1


        LineName = xlWorkbook2.Worksheets("Calculation Sheet").Range("B1").text
        month = int(month)
        Target = int(Target)
        powerpoint_slide = PPTPresentation.Slides.Add(page,12) # Position of Slide, Layout
        #Title
        if month < 1.5 :
            DNS = xlWorkbook.Worksheets("Transfer to PPT").Range("X4").text
            DS =  xlWorkbook.Worksheets("Transfer to PPT").Range("AA4").text
            NS =  xlWorkbook.Worksheets("Transfer to PPT").Range("AD4").text
            if DNS > DS:
                if NS > DNS:MaxtoMin = DS + " - " + NS
                if DNS > NS:
                    if NS > DS: MaxtoMin = DS + " - " + DNS
                    else: MaxtoMin = NS + " - " + DNS
            else:
                if NS > DS : MaxtoMin = DNS + " - " + NS
                if DS > NS : 
                    if NS > DNS: MaxtoMin = DNS + " - " + DS
                    else: MaxtoMin = NS + " - " + DS
                

            title = powerpoint_slide.Shapes.AddTextbox(1, 0, 0, 724, 100) 
            title.TextFrame.TextRange.Text = 'Part 1: Overall Productivity Trend Summary'
            Sizetext(title,20)
            title.TextFrame.TextRange.font.bold = True
            title.Left = 20
            title.Top = 5
            
            #Table
            table1 = CPtable("Transfer to PPT","W3:AD7",powerpoint_slide,1)
            table1.LockAspectRatio = False
            PS(table1,inch(0.45),inch(3.82),inch(9.15),inch(2.24))
            #graph page 1 part 1
            graph = CopyPaste("Transfer to PPT","Chart 54",powerpoint_slide,1)
            graph.LockAspectRatio = False
            PS(graph,inch(0.65),inch(0.63),inch(4.57),inch(2.66))
            textpage1 = LineName+ ' has productivity at ' + MaxtoMin +'OEE. The dayshift operation tends to be more productive than in the nightshift on average.​'

        else:
            MaxtoMin = "..."
            M1text = xlWorkbook.Worksheets("Transfer to PPT").Range("L3").text
            M2text = xlWorkbook.Worksheets("Transfer to PPT").Range("M3").text
            M1 = xlWorkbook.Worksheets("Transfer to PPT").Range("L4").text
            M2 = xlWorkbook.Worksheets("Transfer to PPT").Range("M4").text
            if M1 > M2 : MtoM = M2 + " - " + M1
            else: MtoM = M1 + " - " + M2
            FM1 = xlWorkbook.Worksheets("Transfer to PPT").Range("L7").text
            FM2 = xlWorkbook.Worksheets("Transfer to PPT").Range("M7").text
            if FM1 > FM2 :
                Ftext = "The fluctuation of "+M1text+" (σ = "+FM1+") was worse than "+M2text+" (σ = "+FM2+")."
            else : Ftext = "The fluctuation of "+M2text+" (σ = "+FM2+") was worse than "+M1text+" (σ = "+FM1+")."
            title = powerpoint_slide.Shapes.AddTextbox(1, 0, 0, 724, 100) 
            title.TextFrame.TextRange.Text = 'Part 1: Overall Productivity Trend Summary'
            Sizetext(title,20)
            title.TextFrame.TextRange.font.bold = True
            title.Left = 20
            title.Top = 5

            Picture1 = CPtable("Transfer to PPT","K3:U7",powerpoint_slide,1)
            Picture1.LockAspectRatio = False
            PS(Picture1,inch(0.45),inch(3.82),inch(9.15),inch(2.24))  #graph page 1 part 1
            graph = CopyPaste("Transfer to PPT","Chart 56",powerpoint_slide,1)
            graph.LockAspectRatio = False
            PS(graph,inch(0.45),inch(0.63),inch(9.15),inch(2.68))

            textpage1 = "The average OEE is around " +MtoM+ ". The average OEE of "+M1text+" ("+M1+") was better than that of "+M2text+" ("+M2+"). " + Ftext


        title = powerpoint_slide.Shapes.AddTextbox(1,inch(0.4),inch(3.29),inch(2.89),inch(0.57))
        title.TextFrame.TextRange.Text = 'OEE Statistic (Dayshift+Nightshift)'
        Sizetext(title,14)
        title.TextFrame.TextRange.font.bold = True
        title.Texteffect.Alignment = 2

        title = powerpoint_slide.Shapes.AddTextbox(1,inch(3.7),inch(3.43),inch(2.67),inch(0.34))
        title.TextFrame.TextRange.Text = 'OEE Statistic (Dayshift)'
        Sizetext(title,14)
        title.TextFrame.TextRange.font.bold = True
        title.Texteffect.Alignment = 2

        title = powerpoint_slide.Shapes.AddTextbox(1,inch(6.77),inch(3.44),inch(2.84),inch(0.34))
        title.TextFrame.TextRange.Text = 'OEE Statistic (Nightshift)'
        Sizetext(title,14)
        title.TextFrame.TextRange.font.bold = True
        title.Texteffect.Alignment = 2


        textbelow = powerpoint_slide.Shapes.AddTextbox(1, inch(0.1), inch(6.17), inch(9.85),inch(0.57))
        textbelow.TextFrame.TextRange.Text = textpage1
        colortext(textbelow,0,0,255)
        Sizetext(textbelow,14)
        textbelow.fill.forecolor.rgb = colorcode(255,255,0)
        textbelow.TextFrame.TextRange.font.bold = True



        #Part 2 
        #From Current Productivity Trend Summary to Loss Analysis

        page=page+2
        #Analysis Target Date
        powerpoint_slide = PPTPresentation.Slides.Add(page,12)
        title = powerpoint_slide.Shapes.AddTextbox(1, 0, 0, 724, 100) 
        title.TextFrame.TextRange.Text = 'Part 2: Analysis Target Date'

        if month < 1.5 :
            table = CPtable("Transfer to PPT","A11:J17",powerpoint_slide,1)
            graph = CopyPaste("OEE_Daily_Month1 DS+NS","KOPCHART",powerpoint_slide,1)
            
            textbelow = powerpoint_slide.Shapes.AddTextbox(1, inch(0.1), inch(6.25), inch(9.8),inch(1))
            textbelow.TextFrame.TextRange.Text = 'Selected Loss Analysis Date: '+M1text+', OEE:'+M1+' (Closest to AVG OEE of the Month)'
            colortext(textbelow,0,0,255)
            Sizetext(textbelow,14)
            textbelow.fill.forecolor.rgb = colorcode(255,255,0)
            textbelow.TextFrame.TextRange.font.bold = True
            textbelow.TextEffect.Alignment = 2

        else:
            if Target < 1.5 :
                table = CPtable("Transfer to PPT","A11:J17",powerpoint_slide,1)
                graph = CopyPaste("OEE_Daily_Month1 DS+NS","KOPCHART",powerpoint_slide,1)

                textbelow = powerpoint_slide.Shapes.AddTextbox(1, inch(0.1), inch(6.25), inch(9.8),inch(1))
                textbelow.TextFrame.TextRange.Text = 'Selected Loss Analysis Date: '+M1text+', OEE:'+M1+' (Closest to AVG OEE of the Month)'
                colortext(textbelow,0,0,255)
                Sizetext(textbelow,14)
                textbelow.fill.forecolor.rgb = colorcode(255,255,0)
                textbelow.TextFrame.TextRange.font.bold = True
                textbelow.TextEffect.Alignment = 2

            else:
                table = CPtable("Transfer to PPT","A33:J39",powerpoint_slide,1)
                graph = CopyPaste("OEE_Daily_Month2 DS+NS","KOPCHART",powerpoint_slide,1)

                textbelow = powerpoint_slide.Shapes.AddTextbox(1, inch(0.1), inch(6.25), inch(9.8),inch(1))
                textbelow.TextFrame.TextRange.Text = 'Selected Loss Analysis Date: '+M2text+', OEE:'+M2+' (Closest to AVG OEE of the Month)'
                colortext(textbelow,0,0,255)
                Sizetext(textbelow,14)
                textbelow.fill.forecolor.rgb = colorcode(255,255,0)
                textbelow.TextFrame.TextRange.font.bold = True
                textbelow.TextEffect.Alignment = 2


        table.LockAspectRatio = False 
        graph.LockAspectRatio = False
                    
        PS(table,inch(0.1),inch(4.02),inch(9.79),inch(2))   
        Sizetext(title,20)
        title.TextFrame.TextRange.font.bold = True
        title.Left = 15
        title.Top = 5
        PS(graph,10,45,inch(9.7),inch(3.3))

        #text
        title = powerpoint_slide.Shapes.AddTextbox(1, inch(1.77), inch(0.95), inch(2.04),inch(0.3))
        Texttitle(title,'Target Analysis date')
        colortext(title,0,0,255)
        Sizetext(title,10)
        Fillcolor(title,191,240,255)
        title.TextFrame.TextRange.font.bold = True 
        
        title = powerpoint_slide.Shapes.AddTextbox(1, inch(8.32), inch(1.71), inch(1.5),inch(0.31))
        Texttitle(title,'AVG OEE %')
        colortext(title,253,45,0)
        Sizetext(title,10)
        Fillcolor(title,191,240,255)
        title.TextFrame.TextRange.font.bold = True 


        #xlWorkbook = ExcelApp.Workbooks.Open(Location_excel2)
        xlWorkbook = ExcelApp.Workbooks.Open(lossfile)

        #CT Analysis (4 conditions)
        if xlWorkbook.Worksheets("CT Histogram").Range("E14").text == "Don't have Mode":
            Staticwork = xlWorkbook.Worksheets("CT Histogram")
            STATCT = ["AVG","STDEV","MODE","MIN","P25","P50","P75","MAX"]
            for x in range(8):
                STATCT[x] = xlWorkbook.Worksheets("CT Histogram").Range("E"+str(x+4)).text

            powerpoint_slide = PPTPresentation.Slides[page]
            Table = CopyPaste("CT Histogram","Chart 3",powerpoint_slide,0)
            PS(Table,inch(0.36),inch(3.56),inch(5.59),inch(2.63))
            Table = CopyPaste("OEE Loss Analysis Dashboard","Chart 13",powerpoint_slide,0)
            PS(Table,inch(0.44),inch(0.63),inch(8.86),inch(2.8))

            Table = powerpoint_slide.shapes.AddTable(9,2,inch(6.03),inch(3.52),inch(3.62),inch(2.7))
            row = 1
            while row < 10:
                col = 1
                while col <3:
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.font.size = 12
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                    if row == 1:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(0,67,134)
                    if row == 2 or row == 4 or row == 6 or row == 8:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(203,207,217)
                    col = col + 1
                row = row + 1

            Texttable(Table,1,1,"Statistic")
            Texttable(Table,1,2,"Value​")
            Texttable(Table,2,1,"AVG​")
            Texttable(Table,2,2,STATCT[0] + " sec/pc")
            Texttable(Table,3,1,"STDEV​")
            Texttable(Table,3,2,STATCT[1] + " sec/pc")
            Texttable(Table,4,1,"Mode")
            Texttable(Table,4,2,STATCT[2] + " sec/pc")
            Texttable(Table,5,1,"Min")
            Texttable(Table,5,2,STATCT[3] + " sec/pc")
            Texttable(Table,6,1,"P25")
            Texttable(Table,6,2,STATCT[4] + " sec/pc")
            Texttable(Table,7,1,"P50")
            Texttable(Table,7,2,STATCT[5] + " sec/pc")
            Texttable(Table,8,1,"P75")
            Texttable(Table,8,2,STATCT[6] + " sec/pc")
            Texttable(Table,9,1,"Max")
            Texttable(Table,9,2,STATCT[7] + " sec/pc")

            textbelow = powerpoint_slide.Shapes.AddTextbox(1, inch(0), inch(6.29), inch(10.01),inch(0.34))
            textbelow.TextFrame.TextRange.Text = "Approximately " + xlWorkbook.Worksheets("CT Histogram").Range("AI2").text + " of cycles of operation had CT value lower than or equal to the standard CT. But mode CT could not be found because there is no duplicate data." 
            colortext(textbelow,255,0,0)
            Sizetext(textbelow,14)
            textbelow.fill.forecolor.rgb = colorcode(255,255,0)
            textbelow.TextFrame.TextRange.font.bold = True
            textbelow.Texteffect.Alignment = 2

            title = powerpoint_slide.Shapes.AddTextbox(1, inch(5.75), inch(1.25), inch(2.39),inch(0.51))
            title.TextFrame.TextRange.Text = "Mode CT = "  + xlWorkbook.Worksheets("CT Histogram").Range("E6").text + " sec\n" + "Standard CT =" + xlWorkbook.Worksheets("CT Histogram").Range("E1").text +" sec"
            colortext(title,0,0,255)
            Sizetext(title,12)
            Fillcolor(title,191,240,255)
            title.TextFrame.TextRange.font.bold = True

        elif xlWorkbook.Worksheets("CT Histogram").Range("E14").text == "Single Mode":
            Staticwork = xlWorkbook.Worksheets("CT Histogram")
            STATCT = ["AVG","STDEV","MODE","MIN","P25","P50","P75","MAX"]
            for x in range(8):
                STATCT[x] = xlWorkbook.Worksheets("CT Histogram").Range("E"+str(x+4)).text

            powerpoint_slide = PPTPresentation.Slides[page]
            Table = CopyPaste("CT Histogram","Chart 3",powerpoint_slide,0)
            PS(Table,inch(0.36),inch(3.56),inch(5.59),inch(2.63))
            Table = CopyPaste("OEE Loss Analysis Dashboard","Chart 13",powerpoint_slide,0)
            PS(Table,inch(0.44),inch(0.63),inch(8.86),inch(2.8))

            Table = powerpoint_slide.shapes.AddTable(9,2,inch(6.03),inch(3.52),inch(3.62),inch(2.7))
            row = 1
            while row < 10:
                col = 1
                while col <3:
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.font.size = 12
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                    if row == 1:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(0,67,134)
                    if row == 2 or row == 4 or row == 6 or row == 8:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(203,207,217)
                    col = col + 1
                row = row + 1

            Texttable(Table,1,1,"Statistic")
            Texttable(Table,1,2,"Value​")
            Texttable(Table,2,1,"AVG​")
            Texttable(Table,2,2,STATCT[0] + " sec/pc")
            Texttable(Table,3,1,"STDEV​")
            Texttable(Table,3,2,STATCT[1] + " sec/pc")
            Texttable(Table,4,1,"Mode")
            Texttable(Table,4,2,STATCT[2] + " sec/pc")
            Texttable(Table,5,1,"Min")
            Texttable(Table,5,2,STATCT[3] + " sec/pc")
            Texttable(Table,6,1,"P25")
            Texttable(Table,6,2,STATCT[4] + " sec/pc")
            Texttable(Table,7,1,"P50")
            Texttable(Table,7,2,STATCT[5] + " sec/pc")
            Texttable(Table,8,1,"P75")
            Texttable(Table,8,2,STATCT[6] + " sec/pc")
            Texttable(Table,9,1,"Max")
            Texttable(Table,9,2,STATCT[7] + " sec/pc")

            textbelow = powerpoint_slide.Shapes.AddTextbox(1, inch(0), inch(6.29), inch(10.01),inch(0.34))
            textbelow.TextFrame.TextRange.Text = "Approximately " + xlWorkbook.Worksheets("CT Histogram").Range("AI2").text + " of cycles of operation had CT value lower than or equal to the standard CT. " 
            colortext(textbelow,255,0,0)
            Sizetext(textbelow,14)
            textbelow.fill.forecolor.rgb = colorcode(255,255,0)
            textbelow.TextFrame.TextRange.font.bold = True
            textbelow.Texteffect.Alignment = 2

            title = powerpoint_slide.Shapes.AddTextbox(1, inch(5.75), inch(1.25), inch(2.39),inch(0.51))
            title.TextFrame.TextRange.Text = "Mode CT = "  + xlWorkbook.Worksheets("CT Histogram").Range("E6").text + " sec\n" + "Standard CT =" + xlWorkbook.Worksheets("CT Histogram").Range("E1").text +" sec"
            colortext(title,0,0,255)
            Sizetext(title,12)
            Fillcolor(title,191,240,255)
            title.TextFrame.TextRange.font.bold = True

        elif xlWorkbook.Worksheets("CT Histogram").Range("E14").text == "Multiple Mode" and xlWorkbook.Worksheets("CT Histogram").Range("H14").text == "Single Mode":
            Staticwork = xlWorkbook.Worksheets("CT Histogram")
            STATCT = ["AVG","STDEV","MODE","MIN","P25","P50","P75","MAX"]
            for x in range(8):
                STATCT[x] = xlWorkbook.Worksheets("CT Histogram").Range("E"+str(x+4)).text

            powerpoint_slide = PPTPresentation.Slides[page]
            Table = CopyPaste("CT Histogram","Chart 1",powerpoint_slide,0)
            PS(Table,inch(0.36),inch(3.56),inch(5.59),inch(2.63))
            Table = CopyPaste("OEE Loss Analysis Dashboard","Chart 13",powerpoint_slide,0)
            PS(Table,inch(0.44),inch(0.63),inch(8.86),inch(2.8))

            Table = powerpoint_slide.shapes.AddTable(9,2,inch(6.03),inch(3.52),inch(3.62),inch(2.7))
            row = 1
            while row < 10:
                col = 1
                while col <3:
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.font.size = 12
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                    if row == 1:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(0,67,134)
                    if row == 2 or row == 4 or row == 6 or row == 8:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(203,207,217)
                    col = col + 1
                row = row + 1

            Texttable(Table,1,1,"Statistic")
            Texttable(Table,1,2,"Value​")
            Texttable(Table,2,1,"AVG​")
            Texttable(Table,2,2,STATCT[0] + " sec/pc")
            Texttable(Table,3,1,"STDEV​")
            Texttable(Table,3,2,STATCT[1] + " sec/pc")
            Texttable(Table,4,1,"Mode")
            Texttable(Table,4,2,STATCT[2] + " sec/pc")
            Texttable(Table,5,1,"Min")
            Texttable(Table,5,2,STATCT[3] + " sec/pc")
            Texttable(Table,6,1,"P25")
            Texttable(Table,6,2,STATCT[4] + " sec/pc")
            Texttable(Table,7,1,"P50")
            Texttable(Table,7,2,STATCT[5] + " sec/pc")
            Texttable(Table,8,1,"P75")
            Texttable(Table,8,2,STATCT[6] + " sec/pc")
            Texttable(Table,9,1,"Max")
            Texttable(Table,9,2,STATCT[7] + " sec/pc")

            textbelow = powerpoint_slide.Shapes.AddTextbox(1, inch(0), inch(6.29), inch(10.01),inch(0.34))
            textbelow.TextFrame.TextRange.Text = "Approximately " + xlWorkbook.Worksheets("CT Histogram").Range("AI2").text + " of cycles of operation had CT value lower than or equal to the standard CT. " 
            colortext(textbelow,255,0,0)
            Sizetext(textbelow,14)
            textbelow.fill.forecolor.rgb = colorcode(255,255,0)
            textbelow.TextFrame.TextRange.font.bold = True
            textbelow.Texteffect.Alignment = 2

            title = powerpoint_slide.Shapes.AddTextbox(1, inch(5.75), inch(1.25), inch(2.39),inch(0.51))
            title.TextFrame.TextRange.Text = "Mode CT = "  + xlWorkbook.Worksheets("CT Histogram").Range("E6").text + " sec\n" + "Standard CT =" + xlWorkbook.Worksheets("CT Histogram").Range("E1").text +" sec"
            colortext(title,0,0,255)
            Sizetext(title,12)
            Fillcolor(title,191,240,255)
            title.TextFrame.TextRange.font.bold = True

        elif xlWorkbook.Worksheets("CT Histogram").Range("E14").text == "Multiple Mode" and xlWorkbook.Worksheets("CT Histogram").Range("H14").text == "Multiple Mode":
            Staticwork = xlWorkbook.Worksheets("CT Histogram")
            STATCT = ["AVG","STDEV","MODE","MIN","P25","P50","P75","MAX"]
            for x in range(8):
                STATCT[x] = xlWorkbook.Worksheets("CT Histogram").Range("E"+str(x+4)).text

            powerpoint_slide = PPTPresentation.Slides[page]
            Table = CopyPaste("CT Histogram","Chart 1",powerpoint_slide,0)
            PS(Table,inch(0.36),inch(3.56),inch(5.59),inch(2.63))
            Table = CopyPaste("OEE Loss Analysis Dashboard","Chart 13",powerpoint_slide,0)
            PS(Table,inch(0.44),inch(0.63),inch(8.86),inch(2.8))

            Table = powerpoint_slide.shapes.AddTable(9,2,inch(6.03),inch(3.52),inch(3.62),inch(2.7))
            row = 1
            while row < 10:
                col = 1
                while col <3:
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.font.size = 12
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                    if row == 1:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(0,67,134)
                    if row == 2 or row == 4 or row == 6 or row == 8:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(203,207,217)
                    col = col + 1
                row = row + 1

            Texttable(Table,1,1,"Statistic")
            Texttable(Table,1,2,"Value​")
            Texttable(Table,2,1,"AVG​")
            Texttable(Table,2,2,STATCT[0] + " sec/pc")
            Texttable(Table,3,1,"STDEV​")
            Texttable(Table,3,2,STATCT[1] + " sec/pc")
            Texttable(Table,4,1,"Mode")
            Texttable(Table,4,2,STATCT[2] + " sec/pc")
            Texttable(Table,5,1,"Min")
            Texttable(Table,5,2,STATCT[3] + " sec/pc")
            Texttable(Table,6,1,"P25")
            Texttable(Table,6,2,STATCT[4] + " sec/pc")
            Texttable(Table,7,1,"P50")
            Texttable(Table,7,2,STATCT[5] + " sec/pc")
            Texttable(Table,8,1,"P75")
            Texttable(Table,8,2,STATCT[6] + " sec/pc")
            Texttable(Table,9,1,"Max")
            Texttable(Table,9,2,STATCT[7] + " sec/pc")

            textbelow = powerpoint_slide.Shapes.AddTextbox(1, inch(0), inch(6.29), inch(10.01),inch(0.34))
            textbelow.TextFrame.TextRange.Text = "Approximately " + xlWorkbook.Worksheets("CT Histogram").Range("AI2").text + " of cycles of operation had CT value lower than or equal to the standard CT and Mode CT had multiple data. " 
            colortext(textbelow,255,0,0)
            Sizetext(textbelow,14)
            textbelow.fill.forecolor.rgb = colorcode(255,255,0)
            textbelow.TextFrame.TextRange.font.bold = True
            textbelow.Texteffect.Alignment = 2

            title = powerpoint_slide.Shapes.AddTextbox(1, inch(5.75), inch(1.25), inch(2.39),inch(0.51))
            title.TextFrame.TextRange.Text = "Mode CT = "  + xlWorkbook.Worksheets("CT Histogram").Range("E6").text + " sec\n" + "Standard CT =" + xlWorkbook.Worksheets("CT Histogram").Range("E1").text +" sec"
            colortext(title,0,0,255)
            Sizetext(title,12)
            Fillcolor(title,191,240,255)
            title.TextFrame.TextRange.font.bold = True
            
        page = page + 1

        #OEE Summary chart
        powerpoint_slide = PPTPresentation.Slides[page]
        Table = CopyPaste("OEE Loss Analysis Dashboard","OEE_SUMMARY_CHART",powerpoint_slide,0)
        PS(Table,inch(0.1),inch(0.69),inch(9.7),inch(4.7))

        textbelow = powerpoint_slide.Shapes.AddTextbox(1, inch(0.1), inch(5.48), inch(9.8),inch(1))
        textbelow.TextFrame.TextRange.Text = "- Excessive time from the cycles with CT > 3 standard deviations of the standard CT contributes to " + xlWorkbook.Worksheets("Pivot Table").Range("D6").text + " of OEE Loss on " + xlWorkbook.Worksheets("Calculation Sheet").Range("A1").text + ".\n- Since " + xlWorkbook.Worksheets("Calculation Sheet").Range("A1").text + " had OEE that was the most recent and the closest value to the average OEE during " + xlWorkbook.Worksheets("Calculation Sheet").Range("I2").text + " - " + xlWorkbook.Worksheets("Calculation Sheet").Range("I3").text + " period, we estimate that " + xlWorkbook.Worksheets("Pivot Table").Range("D18").text + " OEE loss structure can be described by the graph above."
        colortext(textbelow,255,0,0)
        Sizetext(textbelow,14)
        textbelow.fill.forecolor.rgb = colorcode(255,255,0)
        textbelow.TextFrame.TextRange.font.bold = True
        textbelow.Texteffect.Alignment = 1

        page = page + 1
        #Pareto chart
        powerpoint_slide = PPTPresentation.Slides[page]
        Table = CopyPaste("OEE Loss Analysis Dashboard","Chart 3",powerpoint_slide,0)
        PS(Table,inch(0.04),inch(0.52),inch(9.86),inch(2.61))

        Picture = CopyPaste("OEE Loss Analysis Dashboard","Chart 6",powerpoint_slide,1)
        Picture.LockAspectRatio = False
        PS(Picture,inch(0),inch(4.04),inch(9.99),inch(2.41))

        #Text
        title = powerpoint_slide.Shapes.AddTextbox(1, inch(4.93), inch(0.98), inch(4.98),inch(0.51))
        Texttitle(title,'[This is the Cumulative OEE Loss in '+ xlWorkbook.Worksheets("Calculation Sheet").Range("B1").text +' due to Each TPM Loss Category on ' + xlWorkbook.Worksheets("Calculation Sheet").Range("A1").text + ']')
        colortext(title,0,0,255)
        Sizetext(title,12)
        Fillcolor(title,191,240,255)
        title.TextFrame.TextRange.font.bold = True

        title = powerpoint_slide.Shapes.AddTextbox(1, inch(5.55), inch(4.3), inch(4.17),inch(0.71))
        Texttitle(title,'[This is the Cumulative OEE Loss due to Each of These Actual Loss Phenomena of the Estimated Root Cause visible by Human Eye]')
        colortext(title,0,0,255)
        Sizetext(title,12)
        Fillcolor(title,191,240,255)
        title.TextFrame.TextRange.font.bold = True

        textbelow = powerpoint_slide.Shapes.AddTextbox(1, inch(0), inch(6.48), inch(10.06),inch(0.3))
        textbelow.TextFrame.TextRange.Text = "The 5 major loss phenomena caused OEE to drop by " + xlWorkbook.Worksheets("Calculation Sheet").Range("I4").text + " ,which is equivalent to " + xlWorkbook.Worksheets("Calculation Sheet").Range("I5").text + " of total OEE loss." 
        colortext(textbelow,255,0,0)
        Sizetext(textbelow,12)
        textbelow.fill.forecolor.rgb = colorcode(255,255,0)
        textbelow.TextFrame.TextRange.font.bold = True
        textbelow.Texteffect.Alignment = 2


        page = page + 1
        #Major loss
        Num = 1
        POR = 7
        LYN = ("B","C","D","E","F")
        Collecttext= ["Statistic","Time Loss (OEE Loss)","No. of Occurrences",10,"Max. Time Loss per Occurrence",10,"Min. Time Loss per Occurrence",10,"Average Time Loss",10,"Standard Deviation of     Time Loss",10]
        for x in range(5):

            for A in [3,5,7,9,11]:
                if A == 3 :
                    Collecttext[A] = xlWorkbook.Worksheets("Calculation Sheet").Range("O"+str(POR)).text + " times"
                else :
                    Collecttext[A] = xlWorkbook.Worksheets("Calculation Sheet").Range("O"+str(POR)).text + " sec\n" + "(= " + xlWorkbook.Worksheets("Calculation Sheet").Range("P"+str(POR)).text + " OEE Loss)"
                POR = POR + 1

            powerpoint_slide = PPTPresentation.Slides[page]
            title = powerpoint_slide.Shapes.AddTextbox(1, 0, 0, 724, 100)
            Texttitle(title,'Part 2: [Major Loss#'+ str(Num) + '] '+ xlWorkbook.Worksheets("Calculation Sheet").Range(LYN[x]+"3").text)
            Sizetext(title,20)
            title.TextFrame.TextRange.font.bold = True
            Picture = CopyPaste("OEE Loss Analysis Dashboard","Chart 6",powerpoint_slide,1)
            Picture.LockAspectRatio = False
            PS(Picture,inch(0),inch(0.8),inch(9.99),inch(2.41))

            title = powerpoint_slide.Shapes.AddTextbox(1, inch(5.53), inch(1.16), inch(3.95),inch(0.71))
            Texttitle(title,'[Major Root Cause Phenomenon#' + str(Num) + ']\n'+ xlWorkbook.Worksheets("Calculation Sheet").Range(LYN[x]+"3").text)
            colortext(title,0,0,255)
            Sizetext(title,12)
            Fillcolor(title,191,240,255)
            title.TextFrame.TextRange.font.bold = True
            title.Texteffect.Alignment = 2
        
            textbelow = powerpoint_slide.Shapes.AddTextbox(1, inch(0.1), inch(6.25), inch(9.8),inch(1))
            textbelow.TextFrame.TextRange.Text = xlWorkbook.Worksheets("Calculation Sheet").Range(LYN[x]+"5").text + " of total time loss on " + xlWorkbook.Worksheets("Calculation Sheet").Range("A1").text + " was from " + xlWorkbook.Worksheets("Calculation Sheet").Range(LYN[x]+"3").text + " (equal to " + xlWorkbook.Worksheets("Calculation Sheet").Range(LYN[x]+"4").text + " of OEE Loss )."
            colortext(textbelow,255,0,0)
            Sizetext(textbelow,14)
            textbelow.fill.forecolor.rgb = colorcode(255,255,0)
            textbelow.TextFrame.TextRange.font.bold = True
            textbelow.Texteffect.Alignment = 2

            Table = powerpoint_slide.shapes.AddTable(6,2,inch(0.41),inch(3.34),inch(5.1),inch(2.53))
            row = 1
            Collectnum = 0
            while row < 7:
                col = 1
                while col <3:
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.font.size = 12.5
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                    if row == 1:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(0,67,134)
                    if row == 2 or row == 4 or row == 6 :
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(203,207,217)
                    Texttable(Table,row,col,Collecttext[Collectnum])
                    Collectnum = Collectnum + 1
                    col = col + 1
                row = row + 1

            page = page + 1
            Num = Num+1

        #CT Chart (Production stability in Normal page1)
        Staticwork = xlWorkbook.Worksheets("CT Histogram")
        STATCT = ["AVG","STDEV","MODE","MIN","P25","P50","P75","MAX"]
        for x in range(8):
            STATCT[x] = xlWorkbook.Worksheets("CT Histogram").Range("E"+str(x+4)).text

        powerpoint_slide = PPTPresentation.Slides[page]
        Table = CopyPaste("OEE Loss Analysis Dashboard","Chart 13",powerpoint_slide,0)
        PS(Table,inch(0.13),inch(0.99),inch(5.94),inch(2.44))

        Table = powerpoint_slide.shapes.AddTable(9,2,inch(6.24),inch(0.7),inch(3.64),inch(2.72))
        row = 1
        while row < 10:
            col = 1
            while col <3:
                Table.Table.Cell(row,col).Shape.TextFrame.TextRange.font.size = 12
                Table.Table.Cell(row,col).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                if row == 1:
                    Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(0,67,134)
                if row == 2 or row == 4 or row == 6 or row == 8:
                    Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(203,207,217)
                col = col + 1
            row = row + 1

        Texttable(Table,1,1,"Statistic")
        Texttable(Table,1,2,"Value​")
        Texttable(Table,2,1,"AVG​")
        Texttable(Table,2,2,STATCT[0] + " sec/pc")
        Texttable(Table,3,1,"STDEV​")
        Texttable(Table,3,2,STATCT[1] + " sec/pc")
        Texttable(Table,4,1,"Mode")
        Texttable(Table,4,2,STATCT[2] + " sec/pc")
        Texttable(Table,5,1,"Min")
        Texttable(Table,5,2,STATCT[3] + " sec/pc")
        Texttable(Table,6,1,"P25")
        Texttable(Table,6,2,STATCT[4] + " sec/pc")
        Texttable(Table,7,1,"P50")
        Texttable(Table,7,2,STATCT[5] + " sec/pc")
        Texttable(Table,8,1,"P75")
        Texttable(Table,8,2,STATCT[6] + " sec/pc")
        Texttable(Table,9,1,"Max")
        Texttable(Table,9,2,STATCT[7] + " sec/pc")

        Staticwork = xlWorkbook.Worksheets("CT Histogram")
        STATCT = ["AVG","STDEV","MODE","MIN","P25","P50","P75","MAX"]
        for x in range(8):
            STATCT[x] = xlWorkbook.Worksheets("CT Histogram").Range("P"+str(x+4)).text

        powerpoint_slide = PPTPresentation.Slides[page]
        Table = CopyPaste("OEE Loss Analysis Dashboard","Chart 20",powerpoint_slide,0)
        PS(Table,inch(0.13),inch(4.02),inch(5.94),inch(2.3))

        Table = powerpoint_slide.shapes.AddTable(9,2,inch(6.24),inch(3.6),inch(3.64),inch(2.72))
        row = 1
        while row < 10:
            col = 1
            while col <3:
                Table.Table.Cell(row,col).Shape.TextFrame.TextRange.font.size = 12
                Table.Table.Cell(row,col).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                if row == 1:
                    Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(0,67,134)
                if row == 2 or row == 4 or row == 6 or row == 8:
                    Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(203,207,217)
                col = col + 1
            row = row + 1

        Texttable(Table,1,1,"Statistic")
        Texttable(Table,1,2,"Value​")
        Texttable(Table,2,1,"AVG​")
        Texttable(Table,2,2,STATCT[0] + " sec/pc")
        Texttable(Table,3,1,"STDEV​")
        Texttable(Table,3,2,STATCT[1] + " sec/pc")
        Texttable(Table,4,1,"Mode")
        Texttable(Table,4,2,STATCT[2] + " sec/pc")
        Texttable(Table,5,1,"Min")
        Texttable(Table,5,2,STATCT[3] + " sec/pc")
        Texttable(Table,6,1,"P25")
        Texttable(Table,6,2,STATCT[4] + " sec/pc")
        Texttable(Table,7,1,"P50")
        Texttable(Table,7,2,STATCT[5] + " sec/pc")
        Texttable(Table,8,1,"P75")
        Texttable(Table,8,2,STATCT[6] + " sec/pc")
        Texttable(Table,9,1,"Max")
        Texttable(Table,9,2,STATCT[7] + " sec/pc")

        #text 
        textbelow = powerpoint_slide.Shapes.AddTextbox(1, inch(0.1), inch(6.25), inch(9.8),inch(1))
        STDCT = xlWorkbook.Worksheets("Analysis Table").Range("BL6").text
        AVG = xlWorkbook.Worksheets("CT Histogram").Range("P4").text
        Mode = xlWorkbook.Worksheets("CT Histogram").Range("P6").text
        Mode = float(Mode)
        AVG = float(AVG)
        STDCT = float(STDCT)

        word = ["","is meet","is slightly lower than","is below","is slightly greater than","is considerably greater than"]

        ShowAVG = ""
        r1 = STDCT+(0.05*STDCT)
        l1 = STDCT-(0.05*STDCT)
        l2 = STDCT-(0.2*STDCT)

        if(AVG==STDCT) :
            ShowAVG = word[1]
        elif(l2<=AVG<STDCT) :
            ShowAVG = word[2]
        elif(AVG<l2) :
            ShowAVG = word[3]
        elif(STDCT<AVG<=r1) :
            ShowAVG = word[4]
        elif (AVG>r1) :
            ShowAVG = word[5]

        ShowMode = ""
        if(Mode==STDCT) :
            ShowMode = word[1]
        elif(l2<=Mode<STDCT) :
            ShowMode = word[2]
        elif(Mode<l2) :
            ShowMode = word[3]
        elif(STDCT<Mode<=r1) :
            ShowMode = word[4]
        elif (Mode>r1) :
            ShowMode = word[5]


        textbelow.TextFrame.TextRange.Text = "In normal condition, the average CT " + ShowAVG + " the standard CT (standard CT = " + xlWorkbook.Worksheets("Analysis Table").Range("BI9").text + " sec) , the mode CT " + ShowMode + " the standard CT, and the standard deviation is " + xlWorkbook.Worksheets("CT Histogram").Range("P5").text + " sec/pc (equivalent to " + xlWorkbook.Worksheets("CT Histogram").Range("P12").text + " of the standard CT of " + xlWorkbook.Worksheets("Analysis Table").Range("BL6").text + " sec)."
        colortext(textbelow,255,0,0)
        Sizetext(textbelow,12)
        textbelow.fill.forecolor.rgb = colorcode(255,255,0)
        textbelow.TextFrame.TextRange.font.bold = True
        textbelow.Texteffect.Alignment = 2

        title = powerpoint_slide.Shapes.AddTextbox(1, inch(0.47), inch(0.68), inch(4.53),inch(0.3))
        Texttitle(title,'Actual CT data on ' + xlWorkbook.Worksheets("Calculation Sheet").Range("A1").text + ' (1 CT:1 cycle)')
        colortext(title,0,0,0)
        Sizetext(title,12)
        title.TextFrame.TextRange.font.bold = True

        title = powerpoint_slide.Shapes.AddTextbox(1, inch(0.23), inch(3.53), inch(5.27),inch(0.51))
        Texttitle(title,'Exclude All Loss Phenomena according to Loss Pareto, Include Only Normal Condition and Sum of Small Losses')
        colortext(title,0,0,0)
        Sizetext(title,12)
        title.TextFrame.TextRange.font.bold = True


        page = page + 1


        #CT Histogram (Production stability in Normal condition page2) (4 conditions)
        if xlWorkbook.Worksheets("CT Histogram").Range("P14").text == "Don't have Mode":
            Staticwork = xlWorkbook.Worksheets("CT Histogram")
            STATCT = ["AVG","STDEV","MODE","MIN","P25","P50","P75","MAX"]
            for x in range(8):
                STATCT[x] = xlWorkbook.Worksheets("CT Histogram").Range("E"+str(x+4)).text
            powerpoint_slide = PPTPresentation.Slides[page]
            Table = CopyPaste("CT Histogram","Chart 3",powerpoint_slide,0)
            PS(Table,inch(0.13),inch(0.99),inch(5.94),inch(2.44))

            Table = powerpoint_slide.shapes.AddTable(9,2,inch(6.24),inch(0.7),inch(3.64),inch(2.72))
            row = 1
            while row < 10:
                col = 1
                while col <3:
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.font.size = 12
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                    if row == 1:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(0,67,134)
                    if row == 2 or row == 4 or row == 6 or row == 8:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(203,207,217)
                    col = col + 1
                row = row + 1

            Texttable(Table,1,1,"Statistic")
            Texttable(Table,1,2,"Value​")
            Texttable(Table,2,1,"AVG​")
            Texttable(Table,2,2,STATCT[0] + " sec/pc")
            Texttable(Table,3,1,"STDEV​")
            Texttable(Table,3,2,STATCT[1] + " sec/pc")
            Texttable(Table,4,1,"Mode")
            Texttable(Table,4,2,STATCT[2] + " sec/pc")
            Texttable(Table,5,1,"Min")
            Texttable(Table,5,2,STATCT[3] + " sec/pc")
            Texttable(Table,6,1,"P25")
            Texttable(Table,6,2,STATCT[4] + " sec/pc")
            Texttable(Table,7,1,"P50")
            Texttable(Table,7,2,STATCT[5] + " sec/pc")
            Texttable(Table,8,1,"P75")
            Texttable(Table,8,2,STATCT[6] + " sec/pc")
            Texttable(Table,9,1,"Max")
            Texttable(Table,9,2,STATCT[7] + " sec/pc")

            Staticwork = xlWorkbook.Worksheets("CT Histogram")
            STATCT = ["AVG","STDEV","MODE","MIN","P25","P50","P75","MAX"]
            for x in range(8):
                STATCT[x] = xlWorkbook.Worksheets("CT Histogram").Range("P"+str(x+4)).text
            powerpoint_slide = PPTPresentation.Slides[page]
            Table = CopyPaste("CT Histogram","Chart 4",powerpoint_slide,0)
            PS(Table,inch(0.13),inch(4.02),inch(5.94),inch(2.3))

            Table = powerpoint_slide.shapes.AddTable(9,2,inch(6.24),inch(3.6),inch(3.64),inch(2.72))
            row = 1
            while row < 10:
                col = 1
                while col <3:
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.font.size = 12
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                    if row == 1:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(0,67,134)
                    if row == 2 or row == 4 or row == 6 or row == 8:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(203,207,217)
                    col = col + 1
                row = row + 1

            Texttable(Table,1,1,"Statistic")
            Texttable(Table,1,2,"Value​")
            Texttable(Table,2,1,"AVG​")
            Texttable(Table,2,2,STATCT[0] + " sec/pc")
            Texttable(Table,3,1,"STDEV​")
            Texttable(Table,3,2,STATCT[1] + " sec/pc")
            Texttable(Table,4,1,"Mode")
            Texttable(Table,4,2,STATCT[2] + " sec/pc")
            Texttable(Table,5,1,"Min")
            Texttable(Table,5,2,STATCT[3] + " sec/pc")
            Texttable(Table,6,1,"P25")
            Texttable(Table,6,2,STATCT[4] + " sec/pc")
            Texttable(Table,7,1,"P50")
            Texttable(Table,7,2,STATCT[5] + " sec/pc")
            Texttable(Table,8,1,"P75")
            Texttable(Table,8,2,STATCT[6] + " sec/pc")
            Texttable(Table,9,1,"Max")
            Texttable(Table,9,2,STATCT[7] + " sec/pc")

            #text 
            textbelow = powerpoint_slide.Shapes.AddTextbox(1, inch(0.1), inch(6.25), inch(9.8),inch(1))
            STDCT = xlWorkbook.Worksheets("Analysis Table").Range("BL6").text
            AVG = xlWorkbook.Worksheets("CT Histogram").Range("P4").text
            Mode = xlWorkbook.Worksheets("CT Histogram").Range("P6").text
            Mode = float(Mode)
            AVG = float(AVG)
            STDCT = float(STDCT)

            word = ["","is meet","is slightly lower than","is below","is slightly greater than","is considerably greater than"]

            ShowAVG = ""
            r1 = STDCT+(0.05*STDCT)
            l1 = STDCT-(0.05*STDCT)
            l2 = STDCT-(0.2*STDCT)

            if(AVG==STDCT) :
                ShowAVG = word[1]
            elif(l2<=AVG<STDCT) :
                ShowAVG = word[2]
            elif(AVG<l2) :
                ShowAVG = word[3]
            elif(STDCT<AVG<=r1) :
                ShowAVG = word[4]
            elif (AVG>r1) :
                ShowAVG = word[5]

            ShowMode = ""
            if(Mode==STDCT) :
                ShowMode = word[1]
            elif(l2<=Mode<STDCT) :
                ShowMode = word[2]
            elif(Mode<l2) :
                ShowMode = word[3]
            elif(STDCT<Mode<=r1) :
                ShowMode = word[4]
            elif (Mode>r1) :
                ShowMode = word[5]

            textbelow.TextFrame.TextRange.Text = "In normal condition, the average CT " + ShowAVG + " the standard CT (standard CT = " + xlWorkbook.Worksheets("Analysis Table").Range("BI9").text + " sec) , the mode CT " + ShowMode + " the standard CT, and the standard deviation is " + xlWorkbook.Worksheets("CT Histogram").Range("P5").text + " sec/pc (equivalent to " + xlWorkbook.Worksheets("CT Histogram").Range("P12").text + " of the standard CT of " + xlWorkbook.Worksheets("Analysis Table").Range("BL6").text + " sec). But mode CT could not be found because there is no duplicate data."
            colortext(textbelow,255,0,0)
            Sizetext(textbelow,12)
            textbelow.fill.forecolor.rgb = colorcode(255,255,0)
            textbelow.TextFrame.TextRange.font.bold = True
            textbelow.Texteffect.Alignment = 2

            title = powerpoint_slide.Shapes.AddTextbox(1, inch(0.47), inch(0.68), inch(4.53),inch(0.3))
            Texttitle(title,'Actual CT data on ' + xlWorkbook.Worksheets("Calculation Sheet").Range("A1").text + ' (1 CT:1 cycle)')
            colortext(title,0,0,0)
            Sizetext(title,12)
            title.TextFrame.TextRange.font.bold = True

            title = powerpoint_slide.Shapes.AddTextbox(1, inch(0.23), inch(3.53), inch(5.27),inch(0.51))
            Texttitle(title,'Exclude All Loss Phenomena according to Loss Pareto, Include Only Normal Condition and Sum of Small Losses')
            colortext(title,0,0,0)
            Sizetext(title,12)
            title.TextFrame.TextRange.font.bold = True

        elif xlWorkbook.Worksheets("CT Histogram").Range("P14").text == "Single Mode":
            Staticwork = xlWorkbook.Worksheets("CT Histogram")
            STATCT = ["AVG","STDEV","MODE","MIN","P25","P50","P75","MAX"]
            for x in range(8):
                STATCT[x] = xlWorkbook.Worksheets("CT Histogram").Range("E"+str(x+4)).text
            powerpoint_slide = PPTPresentation.Slides[page]
            Table = CopyPaste("CT Histogram","Chart 3",powerpoint_slide,0)
            PS(Table,inch(0.13),inch(0.99),inch(5.94),inch(2.44))

            Table = powerpoint_slide.shapes.AddTable(9,2,inch(6.24),inch(0.7),inch(3.64),inch(2.72))
            row = 1
            while row < 10:
                col = 1
                while col <3:
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.font.size = 12
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                    if row == 1:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(0,67,134)
                    if row == 2 or row == 4 or row == 6 or row == 8:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(203,207,217)
                    col = col + 1
                row = row + 1

            Texttable(Table,1,1,"Statistic")
            Texttable(Table,1,2,"Value​")
            Texttable(Table,2,1,"AVG​")
            Texttable(Table,2,2,STATCT[0] + " sec/pc")
            Texttable(Table,3,1,"STDEV​")
            Texttable(Table,3,2,STATCT[1] + " sec/pc")
            Texttable(Table,4,1,"Mode")
            Texttable(Table,4,2,STATCT[2] + " sec/pc")
            Texttable(Table,5,1,"Min")
            Texttable(Table,5,2,STATCT[3] + " sec/pc")
            Texttable(Table,6,1,"P25")
            Texttable(Table,6,2,STATCT[4] + " sec/pc")
            Texttable(Table,7,1,"P50")
            Texttable(Table,7,2,STATCT[5] + " sec/pc")
            Texttable(Table,8,1,"P75")
            Texttable(Table,8,2,STATCT[6] + " sec/pc")
            Texttable(Table,9,1,"Max")
            Texttable(Table,9,2,STATCT[7] + " sec/pc")

            Staticwork = xlWorkbook.Worksheets("CT Histogram")
            STATCT = ["AVG","STDEV","MODE","MIN","P25","P50","P75","MAX"]
            for x in range(8):
                STATCT[x] = xlWorkbook.Worksheets("CT Histogram").Range("P"+str(x+4)).text
            powerpoint_slide = PPTPresentation.Slides[page]
            Table = CopyPaste("CT Histogram","Chart 4",powerpoint_slide,0)
            PS(Table,inch(0.13),inch(4.02),inch(5.94),inch(2.3))

            Table = powerpoint_slide.shapes.AddTable(9,2,inch(6.24),inch(3.6),inch(3.64),inch(2.72))
            row = 1
            while row < 10:
                col = 1
                while col <3:
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.font.size = 12
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                    if row == 1:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(0,67,134)
                    if row == 2 or row == 4 or row == 6 or row == 8:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(203,207,217)
                    col = col + 1
                row = row + 1

            Texttable(Table,1,1,"Statistic")
            Texttable(Table,1,2,"Value​")
            Texttable(Table,2,1,"AVG​")
            Texttable(Table,2,2,STATCT[0] + " sec/pc")
            Texttable(Table,3,1,"STDEV​")
            Texttable(Table,3,2,STATCT[1] + " sec/pc")
            Texttable(Table,4,1,"Mode")
            Texttable(Table,4,2,STATCT[2] + " sec/pc")
            Texttable(Table,5,1,"Min")
            Texttable(Table,5,2,STATCT[3] + " sec/pc")
            Texttable(Table,6,1,"P25")
            Texttable(Table,6,2,STATCT[4] + " sec/pc")
            Texttable(Table,7,1,"P50")
            Texttable(Table,7,2,STATCT[5] + " sec/pc")
            Texttable(Table,8,1,"P75")
            Texttable(Table,8,2,STATCT[6] + " sec/pc")
            Texttable(Table,9,1,"Max")
            Texttable(Table,9,2,STATCT[7] + " sec/pc")

            #text 
            textbelow = powerpoint_slide.Shapes.AddTextbox(1, inch(0.1), inch(6.25), inch(9.8),inch(1))
            STDCT = xlWorkbook.Worksheets("Analysis Table").Range("BL6").text
            AVG = xlWorkbook.Worksheets("CT Histogram").Range("P4").text
            Mode = xlWorkbook.Worksheets("CT Histogram").Range("P6").text
            Mode = float(Mode)
            AVG = float(AVG)
            STDCT = float(STDCT)

            word = ["","is meet","is slightly lower than","is below","is slightly greater than","is considerably greater than"]

            ShowAVG = ""
            r1 = STDCT+(0.05*STDCT)
            l1 = STDCT-(0.05*STDCT)
            l2 = STDCT-(0.2*STDCT)

            if(AVG==STDCT) :
                ShowAVG = word[1]
            elif(l2<=AVG<STDCT) :
                ShowAVG = word[2]
            elif(AVG<l2) :
                ShowAVG = word[3]
            elif(STDCT<AVG<=r1) :
                ShowAVG = word[4]
            elif (AVG>r1) :
                ShowAVG = word[5]

            ShowMode = ""
            if(Mode==STDCT) :
                ShowMode = word[1]
            elif(l2<=Mode<STDCT) :
                ShowMode = word[2]
            elif(Mode<l2) :
                ShowMode = word[3]
            elif(STDCT<Mode<=r1) :
                ShowMode = word[4]
            elif (Mode>r1) :
                ShowMode = word[5]

            textbelow.TextFrame.TextRange.Text = "In normal condition, the average CT " + ShowAVG + " the standard CT (standard CT = " + xlWorkbook.Worksheets("Analysis Table").Range("BI9").text + " sec) , the mode CT " + ShowMode + " the standard CT, and the standard deviation is " + xlWorkbook.Worksheets("CT Histogram").Range("P5").text + " sec/pc (equivalent to " + xlWorkbook.Worksheets("CT Histogram").Range("P12").text + " of the standard CT of " + xlWorkbook.Worksheets("Analysis Table").Range("BL6").text + " sec)."
            colortext(textbelow,255,0,0)
            Sizetext(textbelow,12)
            textbelow.fill.forecolor.rgb = colorcode(255,255,0)
            textbelow.TextFrame.TextRange.font.bold = True
            textbelow.Texteffect.Alignment = 2

            title = powerpoint_slide.Shapes.AddTextbox(1, inch(0.47), inch(0.68), inch(4.53),inch(0.3))
            Texttitle(title,'Actual CT data on ' + xlWorkbook.Worksheets("Calculation Sheet").Range("A1").text + ' (1 CT:1 cycle)')
            colortext(title,0,0,0)
            Sizetext(title,12)
            title.TextFrame.TextRange.font.bold = True

            title = powerpoint_slide.Shapes.AddTextbox(1, inch(0.23), inch(3.53), inch(5.27),inch(0.51))
            Texttitle(title,'Exclude All Loss Phenomena according to Loss Pareto, Include Only Normal Condition and Sum of Small Losses')
            colortext(title,0,0,0)
            Sizetext(title,12)
            title.TextFrame.TextRange.font.bold = True

        elif xlWorkbook.Worksheets("CT Histogram").Range("P14").text == "Multiple Mode" and xlWorkbook.Worksheets("CT Histogram").Range("S14").text == "Single Mode":
            Staticwork = xlWorkbook.Worksheets("CT Histogram")
            STATCT = ["AVG","STDEV","MODE","MIN","P25","P50","P75","MAX"]
            for x in range(8):
                STATCT[x] = xlWorkbook.Worksheets("CT Histogram").Range("E"+str(x+4)).text
            powerpoint_slide = PPTPresentation.Slides[page]
            Table = CopyPaste("CT Histogram","Chart 1",powerpoint_slide,0)
            PS(Table,inch(0.13),inch(0.99),inch(5.94),inch(2.44))

            Table = powerpoint_slide.shapes.AddTable(9,2,inch(6.24),inch(0.7),inch(3.64),inch(2.72))
            row = 1
            while row < 10:
                col = 1
                while col <3:
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.font.size = 12
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                    if row == 1:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(0,67,134)
                    if row == 2 or row == 4 or row == 6 or row == 8:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(203,207,217)
                    col = col + 1
                row = row + 1

            Texttable(Table,1,1,"Statistic")
            Texttable(Table,1,2,"Value​")
            Texttable(Table,2,1,"AVG​")
            Texttable(Table,2,2,STATCT[0] + " sec/pc")
            Texttable(Table,3,1,"STDEV​")
            Texttable(Table,3,2,STATCT[1] + " sec/pc")
            Texttable(Table,4,1,"Mode")
            Texttable(Table,4,2,STATCT[2] + " sec/pc")
            Texttable(Table,5,1,"Min")
            Texttable(Table,5,2,STATCT[3] + " sec/pc")
            Texttable(Table,6,1,"P25")
            Texttable(Table,6,2,STATCT[4] + " sec/pc")
            Texttable(Table,7,1,"P50")
            Texttable(Table,7,2,STATCT[5] + " sec/pc")
            Texttable(Table,8,1,"P75")
            Texttable(Table,8,2,STATCT[6] + " sec/pc")
            Texttable(Table,9,1,"Max")
            Texttable(Table,9,2,STATCT[7] + " sec/pc")

            Staticwork = xlWorkbook.Worksheets("CT Histogram")
            STATCT = ["AVG","STDEV","MODE","MIN","P25","P50","P75","MAX"]
            for x in range(8):
                STATCT[x] = xlWorkbook.Worksheets("CT Histogram").Range("P"+str(x+4)).text
            powerpoint_slide = PPTPresentation.Slides[page]
            Table = CopyPaste("CT Histogram","Chart 2",powerpoint_slide,0)
            PS(Table,inch(0.13),inch(4.02),inch(5.94),inch(2.3))

            Table = powerpoint_slide.shapes.AddTable(9,2,inch(6.24),inch(3.6),inch(3.64),inch(2.72))
            row = 1
            while row < 10:
                col = 1
                while col <3:
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.font.size = 12
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                    if row == 1:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(0,67,134)
                    if row == 2 or row == 4 or row == 6 or row == 8:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(203,207,217)
                    col = col + 1
                row = row + 1

            Texttable(Table,1,1,"Statistic")
            Texttable(Table,1,2,"Value​")
            Texttable(Table,2,1,"AVG​")
            Texttable(Table,2,2,STATCT[0] + " sec/pc")
            Texttable(Table,3,1,"STDEV​")
            Texttable(Table,3,2,STATCT[1] + " sec/pc")
            Texttable(Table,4,1,"Mode")
            Texttable(Table,4,2,STATCT[2] + " sec/pc")
            Texttable(Table,5,1,"Min")
            Texttable(Table,5,2,STATCT[3] + " sec/pc")
            Texttable(Table,6,1,"P25")
            Texttable(Table,6,2,STATCT[4] + " sec/pc")
            Texttable(Table,7,1,"P50")
            Texttable(Table,7,2,STATCT[5] + " sec/pc")
            Texttable(Table,8,1,"P75")
            Texttable(Table,8,2,STATCT[6] + " sec/pc")
            Texttable(Table,9,1,"Max")
            Texttable(Table,9,2,STATCT[7] + " sec/pc")

            #text 
            textbelow = powerpoint_slide.Shapes.AddTextbox(1, inch(0.1), inch(6.25), inch(9.8),inch(1))
            STDCT = xlWorkbook.Worksheets("Analysis Table").Range("BL6").text
            AVG = xlWorkbook.Worksheets("CT Histogram").Range("P4").text
            Mode = xlWorkbook.Worksheets("CT Histogram").Range("P6").text
            Mode = float(Mode)
            AVG = float(AVG)
            STDCT = float(STDCT)

            word = ["","is meet","is slightly lower than","is below","is slightly greater than","is considerably greater than"]

            ShowAVG = ""
            r1 = STDCT+(0.05*STDCT)
            l1 = STDCT-(0.05*STDCT)
            l2 = STDCT-(0.2*STDCT)

            if(AVG==STDCT) :
                ShowAVG = word[1]
            elif(l2<=AVG<STDCT) :
                ShowAVG = word[2]
            elif(AVG<l2) :
                ShowAVG = word[3]
            elif(STDCT<AVG<=r1) :
                ShowAVG = word[4]
            elif (AVG>r1) :
                ShowAVG = word[5]

            ShowMode = ""
            if(Mode==STDCT) :
                ShowMode = word[1]
            elif(l2<=Mode<STDCT) :
                ShowMode = word[2]
            elif(Mode<l2) :
                ShowMode = word[3]
            elif(STDCT<Mode<=r1) :
                ShowMode = word[4]
            elif (Mode>r1) :
                ShowMode = word[5]

            textbelow.TextFrame.TextRange.Text = "In normal condition, the average CT " + ShowAVG + " the standard CT (standard CT = " + xlWorkbook.Worksheets("Analysis Table").Range("BI9").text + " sec) , the mode CT " + ShowMode + " the standard CT, and the standard deviation is " + xlWorkbook.Worksheets("CT Histogram").Range("P5").text + " sec/pc (equivalent to " + xlWorkbook.Worksheets("CT Histogram").Range("P12").text + " of the standard CT of " + xlWorkbook.Worksheets("Analysis Table").Range("BL6").text + " sec)."
            colortext(textbelow,255,0,0)
            Sizetext(textbelow,12)
            textbelow.fill.forecolor.rgb = colorcode(255,255,0)
            textbelow.TextFrame.TextRange.font.bold = True
            textbelow.Texteffect.Alignment = 2

            title = powerpoint_slide.Shapes.AddTextbox(1, inch(0.47), inch(0.68), inch(4.53),inch(0.3))
            Texttitle(title,'Actual CT data on ' + xlWorkbook.Worksheets("Calculation Sheet").Range("A1").text + ' (1 CT:1 cycle)')
            colortext(title,0,0,0)
            Sizetext(title,12)
            title.TextFrame.TextRange.font.bold = True

            title = powerpoint_slide.Shapes.AddTextbox(1, inch(0.23), inch(3.53), inch(5.27),inch(0.51))
            Texttitle(title,'Exclude All Loss Phenomena according to Loss Pareto, Include Only Normal Condition and Sum of Small Losses')
            colortext(title,0,0,0)
            Sizetext(title,12)
            title.TextFrame.TextRange.font.bold = True

        elif xlWorkbook.Worksheets("CT Histogram").Range("P14").text == "Multiple Mode" and xlWorkbook.Worksheets("CT Histogram").Range("S14").text == "Multiple Mode":
            Staticwork = xlWorkbook.Worksheets("CT Histogram")
            STATCT = ["AVG","STDEV","MODE","MIN","P25","P50","P75","MAX"]
            for x in range(8):
                STATCT[x] = xlWorkbook.Worksheets("CT Histogram").Range("E"+str(x+4)).text
            powerpoint_slide = PPTPresentation.Slides[page]
            Table = CopyPaste("CT Histogram","Chart 1",powerpoint_slide,0)
            PS(Table,inch(0.13),inch(0.99),inch(5.94),inch(2.44))

            Table = powerpoint_slide.shapes.AddTable(9,2,inch(6.24),inch(0.7),inch(3.64),inch(2.72))
            row = 1
            while row < 10:
                col = 1
                while col <3:
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.font.size = 12
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                    if row == 1:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(0,67,134)
                    if row == 2 or row == 4 or row == 6 or row == 8:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(203,207,217)
                    col = col + 1
                row = row + 1

            Texttable(Table,1,1,"Statistic")
            Texttable(Table,1,2,"Value​")
            Texttable(Table,2,1,"AVG​")
            Texttable(Table,2,2,STATCT[0] + " sec/pc")
            Texttable(Table,3,1,"STDEV​")
            Texttable(Table,3,2,STATCT[1] + " sec/pc")
            Texttable(Table,4,1,"Mode")
            Texttable(Table,4,2,STATCT[2] + " sec/pc")
            Texttable(Table,5,1,"Min")
            Texttable(Table,5,2,STATCT[3] + " sec/pc")
            Texttable(Table,6,1,"P25")
            Texttable(Table,6,2,STATCT[4] + " sec/pc")
            Texttable(Table,7,1,"P50")
            Texttable(Table,7,2,STATCT[5] + " sec/pc")
            Texttable(Table,8,1,"P75")
            Texttable(Table,8,2,STATCT[6] + " sec/pc")
            Texttable(Table,9,1,"Max")
            Texttable(Table,9,2,STATCT[7] + " sec/pc")

            Staticwork = xlWorkbook.Worksheets("CT Histogram")
            STATCT = ["AVG","STDEV","MODE","MIN","P25","P50","P75","MAX"]
            for x in range(8):
                STATCT[x] = xlWorkbook.Worksheets("CT Histogram").Range("P"+str(x+4)).text
            powerpoint_slide = PPTPresentation.Slides[page]
            Table = CopyPaste("CT Histogram","Chart 2",powerpoint_slide,0)
            PS(Table,inch(0.13),inch(4.02),inch(5.94),inch(2.3))

            Table = powerpoint_slide.shapes.AddTable(9,2,inch(6.24),inch(3.6),inch(3.64),inch(2.72))
            row = 1
            while row < 10:
                col = 1
                while col <3:
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.font.size = 12
                    Table.Table.Cell(row,col).Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 2
                    if row == 1:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(0,67,134)
                    if row == 2 or row == 4 or row == 6 or row == 8:
                        Table.Table.Cell(row,col).Shape.Fill.forecolor.rgb = RGB(203,207,217)
                    col = col + 1
                row = row + 1

            Texttable(Table,1,1,"Statistic")
            Texttable(Table,1,2,"Value​")
            Texttable(Table,2,1,"AVG​")
            Texttable(Table,2,2,STATCT[0] + " sec/pc")
            Texttable(Table,3,1,"STDEV​")
            Texttable(Table,3,2,STATCT[1] + " sec/pc")
            Texttable(Table,4,1,"Mode")
            Texttable(Table,4,2,STATCT[2] + " sec/pc")
            Texttable(Table,5,1,"Min")
            Texttable(Table,5,2,STATCT[3] + " sec/pc")
            Texttable(Table,6,1,"P25")
            Texttable(Table,6,2,STATCT[4] + " sec/pc")
            Texttable(Table,7,1,"P50")
            Texttable(Table,7,2,STATCT[5] + " sec/pc")
            Texttable(Table,8,1,"P75")
            Texttable(Table,8,2,STATCT[6] + " sec/pc")
            Texttable(Table,9,1,"Max")
            Texttable(Table,9,2,STATCT[7] + " sec/pc")

            #text 
            textbelow = powerpoint_slide.Shapes.AddTextbox(1, inch(0.1), inch(6.25), inch(9.8),inch(1))
            STDCT = xlWorkbook.Worksheets("Analysis Table").Range("BL6").text
            AVG = xlWorkbook.Worksheets("CT Histogram").Range("P4").text
            Mode = xlWorkbook.Worksheets("CT Histogram").Range("P6").text
            Mode = float(Mode)
            AVG = float(AVG)
            STDCT = float(STDCT)

            word = ["","is meet","is slightly lower than","is below","is slightly greater than","is considerably greater than"]

            ShowAVG = ""
            r1 = STDCT+(0.05*STDCT)
            l1 = STDCT-(0.05*STDCT)
            l2 = STDCT-(0.2*STDCT)

            if(AVG==STDCT) :
                ShowAVG = word[1]
            elif(l2<=AVG<STDCT) :
                ShowAVG = word[2]
            elif(AVG<l2) :
                ShowAVG = word[3]
            elif(STDCT<AVG<=r1) :
                ShowAVG = word[4]
            elif (AVG>r1) :
                ShowAVG = word[5]

            ShowMode = ""
            if(Mode==STDCT) :
                ShowMode = word[1]
            elif(l2<=Mode<STDCT) :
                ShowMode = word[2]
            elif(Mode<l2) :
                ShowMode = word[3]
            elif(STDCT<Mode<=r1) :
                ShowMode = word[4]
            elif (Mode>r1) :
                ShowMode = word[5]

            title = powerpoint_slide.Shapes.AddTextbox(1, inch(0.47), inch(0.68), inch(4.53),inch(0.3))
            Texttitle(title,'Actual CT data on ' + xlWorkbook.Worksheets("Calculation Sheet").Range("A1").text + ' (1 CT:1 cycle)')
            colortext(title,0,0,0)
            Sizetext(title,12)
            title.TextFrame.TextRange.font.bold = True

            title = powerpoint_slide.Shapes.AddTextbox(1, inch(0.23), inch(3.53), inch(5.27),inch(0.51))
            Texttitle(title,'Exclude All Loss Phenomena according to Loss Pareto, Include Only Normal Condition and Sum of Small Losses')
            colortext(title,0,0,0)
            Sizetext(title,12)
            title.TextFrame.TextRange.font.bold = True
            textbelow.TextFrame.TextRange.Text = "The CT frequency distribution has multiple mode CT values (Mode A, Mode B, Mode C, ..., Mode n) which signifies that there is a considerable amount of fluctuation in this production line."
            colortext(textbelow,255,0,0)
            Sizetext(textbelow,12)
            textbelow.fill.forecolor.rgb = colorcode(255,255,0)
            textbelow.TextFrame.TextRange.font.bold = True
            textbelow.Texteffect.Alignment = 2

        page = page + 1
        '''
        # Save as the presentation
        PPTPresentation.SaveAs(r"")
        '''