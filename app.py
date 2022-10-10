import os, json, locale
from flask import Flask,redirect,url_for, render_template, request, send_file
import os.path, razorpay
import time, string,urllib.parse
from decimal import Decimal
from datetime import date
from reportlab.pdfgen import canvas
from reportlab.lib.enums import TA_JUSTIFY
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.graphics import renderPM
from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import Table, TableStyle
from reportlab.graphics.shapes import*
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.rl_config import defaultPageSize
from PIL import Image
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy.sql import text
from google.oauth2 import service_account


locale.setlocale(locale.LC_ALL, 'en_IN')
app=Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///estimator.db' 
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = True


@app.route("/", methods=['GET','POST'])
def home():	
	if request.method=='POST':
		cname = request.form['work']
		gatno = request.form['gno']
		address=gatno + ", " + request.form['add']
		project=request.form['project']
		length=request.form['length']
		mobile=request.form['mobile']
		area=request.form['barea']
		kitchen=request.form['kitchen']
		bath=request.form['bath']
		floors =request.form['floors']
		mailid=request.form['mailid']
		


		SERVICE_ACCOUNT_FILE = 'keys.json'

		SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

		cred = None
		cred = service_account.Credentials.from_service_account_file(
			SERVICE_ACCOUNT_FILE, scopes=SCOPES)
		SAMPLE_SPREADSHEET_ID = '1uogWguNxoRXTY6sKVH6on2DHXwkriFc9RtkoqusF9dc'

		service = build('sheets', 'v4', credentials=cred)

		# Call the Sheets API
		sheet = service.spreadsheets()
		

		val=[[cname]]
		
		
		databank=[{'range':'ABSTRACT_SHEET!C8:K8','values':[[cname.upper()]]},
				  {'range':'ABSTRACT_SHEET!C6:K6','values':[[address.upper()]]},
				  {'range':'ABSTRACT_SHEET!O5:Q5','values':[[project.upper()]]},
				  {'range':'ABSTRACT_SHEET!O13:Q13','values':[[length]]},
				  {'range':'ABSTRACT_SHEET!Q7','values':[[area]]},
				  {'range':'ABSTRACT_SHEET!O8:P8','values':[[kitchen]]},
				  {'range':'ABSTRACT_SHEET!O9:P9','values':[[bath]]},
				  {'range':'ABSTRACT_SHEET!O10:Q10','values':[[floors]]}]
		bodydata={'valueInputOption':'USER_ENTERED','data':databank}
		#print(databank)
		#print(bodydata)
		batch_update=sheet.values().batchUpdate(spreadsheetId=SAMPLE_SPREADSHEET_ID,body=bodydata).execute()
		result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,range="ABSTRACT_SHEET!K71").execute()
		print(result)
		amt=result['values'][0][0]
		#abstractdata = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,range="ABSTRACT_SHEET!B13:K71").execute()
		
		#print(address.upper())
		#link=render_template("viewdetails.html",custname=cname.upper(),add=address.upper(),project=project.upper(),length=length,width=width,area=area,kitchen=kitchen,bath=bath,floors=floors.upper(),amt=amt,mailid=mailid.lower())
		#print(link)
		keyid="rzp_test_OzPBvsohQVxA7z"
		keysecret="0a46y7ZCCqnkDnouoMnf6u40"
		rzpclient = razorpay.Client(auth=(keyid,keysecret))
		genpamt=500
		order_amount = genpamt*100
		order_currency = 'INR'
		order_receipt = 'order_rcptid' + str(int(time.time()))
		payment = rzpclient.order.create({'amount':order_amount, 'currency':order_currency, 'receipt':order_receipt})
		
		return 	render_template('viewdetails.html',custname=cname.upper(),add=address.upper(),project=project.upper(),length=length,area=area,kitchen=kitchen,bath=bath,floors=floors.upper(),amt=amt,mailid=mailid.lower(),mobile=mobile,orderid=order_receipt,pamt=order_amount,payment=payment)
	
	return render_template("bsindex.html")

@app.route("/view/<path:custname>/<path:add>/<path:project>/<length>/<area>/<kitchen>/<bath>/<floors>/<amt>/<mailid>/<mobile>/<orderid>/<pamt>",methods=["GET", "POST"])
def view(custname,add,project,length,area,kitchen,bath,floors,amt,mailid,mobile,orderid,pamt):
#########Generate Razorpay Order
	print(pamt)
	print(orderid)
	return 	render_template('viewdetails.html',custname=custname.upper(),add=add.upper(),project=project.upper(),length=length,area=area,kitchen=kitchen,bath=bath,floors=floors.upper(),amt=amt,mailid=mailid.lower(),mobile=mobile,orderid=orderid,pamt=pamt)
	
	#if custname==str(result['values'][0][0]):
	
		
		#print(add)
		#print(amt)
		#amount = float(amt.replace(',',''))
		#print("Amount :"+str(amount))
		#print(bath)
		#clength=("{:.2f}".format(float(length)))
		#cwidth=("{:.2f}".format(float(width)))
		#carea=("{:.2f}".format(float(area)))
		#cbath=("{:.2f}".format(float(bath)))
		#ckitchen=("{:.2f}".format(float(kitchen)))
		#camt=float(("{:.2f}".format(float(amt.replace(',', '')))))
		#print(float(camt))
		#locale.setlocale(locale.LC_NUMERIC, 'en_IN')
		#print( locale.format("%d", camt, grouping=True))
		#camt = "Rs."+ locale.format("%d", camt, grouping=True)+".00/-"
		#print(cwidth)
		#print(carea)
		#print(cbath)
		#print(ckitchen)
		#print(floors)
		#print(camt)
		#print(mailid)
		#print(bath)
	
	#return render_template("viewdetails.html",custname=custname.upper(),add=add.upper(),project=project.upper(),length=length,width=width,area=area,kitchen=kitchen,bath=bath,floors=floors.upper(),amt=amt,mailid=mailid.lower())
	#else:
	#	return redirect(url_for('home'))
	
	#if request.method=='POST':
	#	print(request.method['btncon'])

@app.route("/success/<path:custname>/<path:add>/<path:project>/<length>/<area>/<kitchen>/<bath>/<path:floors>/<mailid>/<mobile>",methods=["GET", "POST"])
def success(custname,add,project,length,area,kitchen,bath,floors,amt,mailid):
	return redirect(url_for("confirm",custname=custname.upper(),add=add.upper(),project=project.upper(),length=length,width=width,area=area,kitchen=kitchen,bath=bath,floors=floors.upper(),amt=amt,mailid=mailid.lower()))


@app.route("/confirm/<path:custname>/<path:add>/<path:project>/<length>/<area>/<kitchen>/<bath>/<path:floors>/<mailid>/<mobile>/<orderid>",methods=["GET", "POST"])
def confirm(custname,add,project,length,area,kitchen,bath,floors,mailid,mobile,orderid):
	if request.method=='POST':
		if request.form.get("btnprocess"):
			
			address = add
			project = project
			length = length
			
			area = area
			kitchen = kitchen
			bath = bath
			floors = floors
			mailid = mailid
			cname = request.form.get('cwork')
			cadd = request.form.get('cadd')
			cproject = request.form.get('cproject')
			clength = request.form.get('clength')
			cmobile = request.form.get('cmobile')
			carea = request.form.get('carea')
			cbath = request.form.get('cbath')
			ckitchen = request.form.get('ckitchen')
			cfloors = request.form.get('cfloors')
			cmailid = request.form.get('cmailid')
			oid = request.form.get('oid')
			pid = request.form.get('pid')
			rs=request.form.get('rs')
			print(rs)
			

			keyid="rzp_test_OzPBvsohQVxA7z"
			keysecret="0a46y7ZCCqnkDnouoMnf6u40"
			client = razorpay.Client(auth=(keyid,keysecret))

			params_dict = {
				'razorpay_order_id': oid,
				'razorpay_payment_id': pid,
				'razorpay_signature': rs
			}
			client.utility.verify_payment_signature(params_dict)
			print(res)

			
			#update_Data(custname=custname,add=address,project=project,length=length,width=width,area=area,kitchen=kitchen,bath=bath,floors=floors,mailid=mailid)
			SERVICE_ACCOUNT_FILE = 'keys.json'

			SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

			cred = None
			cred = service_account.Credentials.from_service_account_file(
				SERVICE_ACCOUNT_FILE, scopes=SCOPES)
			SAMPLE_SPREADSHEET_ID = '1uogWguNxoRXTY6sKVH6on2DHXwkriFc9RtkoqusF9dc'

			service = build('sheets', 'v4', credentials=cred)

			# Call the Sheets API
			sheet = service.spreadsheets()
			

			
			databank=[{'range':'ABSTRACT_SHEET!C8:K8','values':[[cname.upper()]]},
						{'range':'ABSTRACT_SHEET!C6:K6','values':[[cadd.upper()]]},
						{'range':'ABSTRACT_SHEET!O5:Q5','values':[[cproject.upper()]]},
						{'range':'ABSTRACT_SHEET!O13:Q13','values':[[clength]]},
						{'range':'ABSTRACT_SHEET!Q7','values':[[carea]]},
						{'range':'ABSTRACT_SHEET!O8:P8','values':[[ckitchen]]},
						{'range':'ABSTRACT_SHEET!O9:P9','values':[[cbath]]},
						{'range':'ABSTRACT_SHEET!O10:Q10','values':[[cfloors]]}]
			bodydata={'valueInputOption':'USER_ENTERED','data':databank}
			#print(databank)
			#print(bodydata)
			batch_update=sheet.values().batchUpdate(spreadsheetId=SAMPLE_SPREADSHEET_ID,body=bodydata).execute()
			batch_update=sheet.values().batchUpdate(spreadsheetId=SAMPLE_SPREADSHEET_ID,body=bodydata).execute()
			USER_SPREADSHEET_ID = '1WArCvUNrjekWioburcvXS8l1Y0GmM0eY7ij97bVk6QI'

			service1 = build('sheets', 'v4', credentials=cred)

			# Call the Sheets API
			userdata=[[cname.upper(),cproject.upper(),cadd.upper(),clength,carea,ckitchen,cbath,cfloors,cmailid,mobile]]
			usersheet = service1.spreadsheets()
			update_users=usersheet.values().append(spreadsheetId=USER_SPREADSHEET_ID,range="USERDATA!B2", valueInputOption="USER_ENTERED",body={"values":userdata}).execute()
					
					
			gen_pdf(cname.upper(),cadd.upper(),cproject.upper(),clength,carea,ckitchen,cbath,cfloors.upper(),cmailid)
			filename="ESTIMATE FOR " +cname.upper()+ ".pdf"
			return send_file('Estimate.pdf', attachment_filename=filename, as_attachment=False)
			
		elif request.form.get("btnback"):
			return redirect('/')

def update_Data(custname,add,project,length,width,area,kitchen,bath,floors,mailid):
	
	SERVICE_ACCOUNT_FILE = 'keys.json'

	SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

	cred = None
	cred = service_account.Credentials.from_service_account_file(
		SERVICE_ACCOUNT_FILE, scopes=SCOPES)
	SAMPLE_SPREADSHEET_ID = '1uogWguNxoRXTY6sKVH6on2DHXwkriFc9RtkoqusF9dc'

	service = build('sheets', 'v4', credentials=cred)

	# Call the Sheets API
	sheet = service.spreadsheets()
	

	
	databank=[{'range':'ABSTRACT_SHEET!C8:K8','values':[[custname.upper()]]},
				  {'range':'ABSTRACT_SHEET!C6:K6','values':[[add.upper()]]},
				  {'range':'ABSTRACT_SHEET!O5:Q5','values':[[project.upper()]]},
				  {'range':'ABSTRACT_SHEET!O7','values':[[length]]},
				  {'range':'ABSTRACT_SHEET!P7','values':[[width]]},
				  {'range':'ABSTRACT_SHEET!Q7','values':[[area]]},
				  {'range':'ABSTRACT_SHEET!O8:P8','values':[[kitchen]]},
				  {'range':'ABSTRACT_SHEET!O9:P9','values':[[bath]]},
				  {'range':'ABSTRACT_SHEET!O10:Q10','values':[[floors]]}]
	bodydata={'valueInputOption':'USER_ENTERED','data':databank}
	print(databank)
	print(bodydata)
	batch_update=sheet.values().batchUpdate(spreadsheetId=SAMPLE_SPREADSHEET_ID,body=bodydata).execute()





def gen_pdf(custname,add,project,length,area,kitchen,bath,floors,mailid):
	Title = "Hello world"
	pageinfo = "platypus example"
	print(length)
	print(area)
	def sstable(width,height):
		#pdf.setFont("Roboto-Regular", 12)
		enggname="""SAGAR VIJAYRAO KATORE""" + "<br/><font size=12>CONSULTING ENGINEER, STRUCTURAL DESIGNER & GOVT. REGISTERED VALUER</font>"""
		enggadd = """At : Nimgaon Korhale, Post : Laxmiwadi, Tal : Rahata,Ahmednagar - 423 109 <br/>+91- 9960 867 555, 9859 121 121 """
		parastyle = ParagraphStyle('para1')
		parastyle.fontSize = 14
		parastyle.fontName='Roboto-Bold'
		parastyle.alignment=TA_CENTER
		parastyle.leading= 20
		engname=Paragraph(enggname,parastyle)
		parastyle = ParagraphStyle('para1')
		parastyle.alignment=TA_CENTER
		parastyle.fontSize = 12
		parastyle.leading= 20
		engadd =Paragraph(enggadd,parastyle)
		ERNAME=[engname,engadd]
		parastyle = ParagraphStyle('para1')
		parastyle.alignment=TA_CENTER
		parastyle.fontName='Roboto-Regular'
		parastyle.fontSize = 12
		parastyle.leading= 20
		namount =Paragraph("Rs."+str(namt['values'][0][0])+"/-",parastyle)
		netamt=[namount]
		parastyle = ParagraphStyle('para1')
		parastyle.alignment=TA_CENTER
		parastyle.fontName='Roboto-Regular'
		parastyle.fontSize = 12
		parastyle.leading= 20
		eleamt =Paragraph("Rs."+str(eamt['values'][0][0])+"/-",parastyle)
		eleamount=[eleamt]
		parastyle = ParagraphStyle('para1')
		parastyle.alignment=TA_CENTER
		parastyle.fontName='Roboto-Regular'
		parastyle.fontSize = 12
		parastyle.leading= 20
		conamt =Paragraph("Rs."+str(camt['values'][0][0])+"/-",parastyle)
		camount=[conamt]
		parastyle = ParagraphStyle('para1')
		parastyle.alignment=TA_CENTER
		parastyle.fontName='Roboto-Bold'
		parastyle.fontSize = 12
		parastyle.leading= 20
		tamount =Paragraph("Rs."+str(amt['values'][0][0])+"/-",parastyle)
		tamt=[tamount]
		parastyle = ParagraphStyle('para1')
		parastyle.alignment=TA_CENTER
		parastyle.fontName='Roboto-Bold'
		parastyle.fontSize = 12
		parastyle.leading= 20
		inwrds =Paragraph(inwords['values'][0][0],parastyle)
		iwrds=[inwrds]




		widthList=[width*0.35,width*0.65,]
		heightList=[height*0.18,height*0.18,height*0.18,height*0.18,height*0.26,]

		res=Table([
			['Estimate Value',netamt],
			['Add 2.50% for Contingencies',conamt],
			['Add 2.50% for Electrification',eleamount],
			['Total Estimate Value',tamt],
			['In Words',iwrds],        
			
		],colWidths=widthList,rowHeights=heightList)


		res.setStyle([
			('GRID', (0,0), (-1,-1),1,'black'),
			('ALIGN',(0,0), (-1,-1),'CENTER'),
			('VALIGN',(0,0), (-1,5),'MIDDLE'),
			
			('FONT',(0,0),(0,0),'Roboto-Bold'),
			('FONTSIZE',(0,0),(0,0),12),
			('BOTTOMPADDING',(0,0),(0,0),20),
			('FONT',(0,1),(0,1),'Roboto-Bold'),
			('FONTSIZE',(0,1),(0,1),12),
			('FONT',(0,2),(0,2),'Roboto-Bold'),
			('FONTSIZE',(0,2),(0,2),12),
			('FONT',(0,3),(1,3),'Roboto-Bold'),
			('FONTSIZE',(0,3),(1,3),12),
			('FONT',(0,4),(0,4),'Roboto-Bold'),
			('FONTSIZE',(0,4),(0,4),12),
			('FONT',(1,3),(1,4),'Roboto-Bold'),
			('FONTSIZE',(1,3),(1,4),12),
			

		])
		return res

	def sftable(swidth,sheight,customername,work,area,floors):
		#pdf.setFont("Roboto-Regular", 12)
		enggname="""SAGAR VIJAYRAO KATORE""" + "<br/><font size=12>CONSULTING ENGINEER, STRUCTURAL DESIGNER & GOVT. REGISTERED VALUER</font>"""
		enggadd = """At : Nimgaon Korhale, Post : Laxmiwadi, Tal : Rahata,Ahmednagar - 423 109 <br/>+91- 9960 867 555, 9859 121 121 """
		parastyle = ParagraphStyle('para1')
		parastyle.fontSize = 14
		parastyle.fontName='Roboto-Bold'
		parastyle.alignment=TA_CENTER
		parastyle.leading= 20
		engname=Paragraph(enggname,parastyle)
		parastyle = ParagraphStyle('para1')
		parastyle.alignment=TA_CENTER
		parastyle.fontName='Roboto-Regular'
		parastyle.fontSize = 12
		parastyle.leading= 20
		work =Paragraph(work,parastyle)
		workdetail=[work]
		parastyle = ParagraphStyle('para1')
		parastyle.alignment=TA_CENTER
		parastyle.fontName='Roboto-Regular'
		parastyle.fontSize = 12
		parastyle.leading= 20
		cliname =Paragraph(customername,parastyle)
		clname=[cliname]
		cliarea =Paragraph(area+".00 SQ.FT.",parastyle)
		clarea=[cliarea]
		clifloors =Paragraph(floors,parastyle)
		clfloors=[clifloors]

		widthList=[swidth*0.35,swidth*0.65,]
		heightList=[sheight*0.55,sheight*0.15,sheight*0.15,sheight*0.15,]

		res=Table([
			['Name of Work',workdetail],
			['Name of Client',clname],
			['Total Builtup Area',clarea],
			['No of Floors',clfloors],
			
		],colWidths=widthList,rowHeights=heightList)


		res.setStyle([
			('GRID', (0,0), (-1,-1),1,'black'),
			('ALIGN',(0,0), (-1,-1),'CENTER'),
			('VALIGN',(0,0), (0,0),'MIDDLE'),
			('VALIGN',(0,0), (-1,-1),'MIDDLE'),
			('FONT',(0,0),(0,0),'Roboto-Bold'),
			('FONTSIZE',(0,0),(0,0),14),
			('BOTTOMPADDING',(0,0),(0,0),20),
			('FONT',(0,1),(0,1),'Roboto-Bold'),
			('FONTSIZE',(0,1),(0,1),14),
			('FONT',(0,3),(0,3),'Roboto-Bold'),
			('FONTSIZE',(0,3),(0,3),14),
			('FONT',(0,2),(0,2),'Roboto-Bold'),
			('FONTSIZE',(0,2),(0,2),14),

		])
		return res
	

	def tftable(width,height,work):
		#pdf.setFont("Roboto-Regular", 10)
		parastyle = ParagraphStyle('para1')
		parastyle.alignment=TA_LEFT
		parastyle.fontName='Roboto-Regular'
		parastyle.fontSize = 12
		parastyle.leading= 15
		work =Paragraph(work,parastyle)
		workdetail=[work]
		widthList=[width*0.2,width*0.80,]
		heightList=[height*1,]

		res=Table([
			['WORK :',workdetail],
			
			
		],colWidths=widthList,rowHeights=heightList)


		res.setStyle([
			('GRID', (0,0), (-1,-1),1,'black'),
			('ALIGN',(0,0), (-1,-1),'CENTER'),
			('VALIGN',(0,0), (0,0),'MIDDLE'),
			('VALIGN',(0,0), (-1,-1),'MIDDLE'),
			('FONT',(0,0),(0,0),'Roboto-Bold'),
			('FONTSIZE',(0,0),(0,0),12),
			('BOTTOMPADDING',(0,0),(0,0),20),
			('FONT',(0,1),(0,1),'Roboto-Bold'),
			('FONTSIZE',(0,1),(0,1),14),
			('FONT',(0,3),(0,3),'Roboto-Bold'),
			('FONTSIZE',(0,3),(0,3),12),
			

		])
		return res
	
	def tstable(width,height,work):
		#pdf.setFont("Roboto-Regular", 10)
		parastyle = ParagraphStyle('para1')
		parastyle.alignment=TA_LEFT
		parastyle.fontName='Roboto-Regular'
		parastyle.fontSize = 12
		parastyle.leading= 15
		work =Paragraph(work,parastyle)
		workdetail=[work]
		widthList=[width*0.2,width*0.80,]
		heightList=[height*1,]

		res=Table([
			['CLIENT :',workdetail],
			
			
		],colWidths=widthList,rowHeights=heightList)


		res.setStyle([
			('GRID', (0,0), (-1,-1),1,'black'),
			('ALIGN',(0,0), (-1,-1),'CENTER'),
			('VALIGN',(0,0), (-1,-1),'MIDDLE'),
			
			('FONT',(0,0),(0,0),'Roboto-Bold'),
			('FONTSIZE',(0,0),(0,0),12),
			
			('FONT',(0,1),(0,1),'Roboto-Bold'),
			('FONTSIZE',(0,1),(0,1),14),
			('FONT',(0,3),(0,3),'Roboto-Bold'),
			('FONTSIZE',(0,3),(0,3),12),

		])
		return res
	def tttable(width,height):
		#pdf.setFont("Roboto-Regular", 10)
		widthList=[width*1,]
		heightList=[height*1,]
		parastyle = ParagraphStyle('para1')
		parastyle.alignment=TA_LEFT
		parastyle.fontName='Roboto-Regular'
		parastyle.fontSize = 12
		parastyle.leading= 15
		res=Table([
			['ABSTRACT SHEET'],
			
			
		],colWidths=widthList,rowHeights=heightList)


		res.setStyle([
			('GRID', (0,0), (-1,-1),1,'black'),
			('ALIGN',(0,0), (-1,-1),'CENTER'),
			('VALIGN',(0,0), (-1,-1),'MIDDLE'),
			('FONT',(0,0),(0,0),'Roboto-Bold'),
			('FONTSIZE',(0,0),(0,0),12),
			])
		return res
	
	def tsheet(width,height,td):
		#pdf.setFont("Roboto-Regular", 10)
		parastyle = ParagraphStyle('para1')
		parastyle.alignment=TA_LEFT
		parastyle.fontName='Roboto-Regular'
		parastyle.fontSize = 12
		parastyle.leading= 15
		#cname =Paragraph(cname,parastyle)
		#workdetail=[cname]
		widthList=[width*0.06,width*0.44,width*0.11,width*0.07,width*0.13,width*0.18,]
		heightList=[height*1,]

		res=Table(td,widthList)


		res.setStyle([
			('GRID', (0,0), (-1,-1),1,'black'),
			('ALIGN',(0,0), (-1,-1),'CENTER'),
			('VALIGN',(0,0), (-1,-1),'MIDDLE'),
			('FONT',(0,0),(0,0),'Roboto-Bold'),
			('FONTSIZE',(0,0),(0,0),12),
			
			('FONT',(0,-1),(0,-1),'Roboto-Bold'),
			('FONTSIZE',(0,-1),(0,-1),12),
			('FONT',(0,55),(5,55),'Roboto-Bold'),
			('FONTSIZE',(0,55),(5,55),12),
			('FONT',(0,59),(5,59),'Roboto-Bold'),
			('FONTSIZE',(0,59),(5,59),12),
			('SPAN',(2,59),(4,59)),
			

		])
		return res


	def myFirstPage(canvas, doc):
		canvas.saveState()
		canvas.setTitle('Estimate')
		canvas.setFont("Roboto-Regular", 12)
		canvas.rect(0.5*inch, 0.5*inch, 7.27*inch, 10.69*inch, fill=0)
		#fpageTable.wrapOn(canvas,0,0)
		#fpageTable.drawOn(canvas,0,0)
		#canvas.drawCentredString(PAGE_WIDTH/2.0, PAGE_HEIGHT-108, Title)
		#canvas.drawString(inch, 0.75 * inch, "First Page / %s" % pageinfo)
		canvas.restoreState()
		#canvas.showPage()

	def myLaterPages(canvas, doc):
		canvas.saveState()
		fn = os.path.join(os.path.dirname(os.path.abspath(__file__)),"static","images",'LH.png')
		img = Image.open(fn)
		logo = ImageReader(fn)
		
		#canvas.drawImage(fn,0,0,*A4)

		###############################
		vfn = os.path.join(os.path.dirname(os.path.abspath(__file__)),"static","images",'LHH.png')
		vimg = Image.open(fn)
		vlogo = ImageReader(vfn)
		canvas.drawImage(vfn,0*inch,10.05*inch,8.27*inch,1.6*inch)
		canvas.setFont("Roboto-Bold", 24)
		#canvas.drawCentredString(4.15*inch, 11.2 * inch, "VKON CONSULTANTS")
		canvas.setFont("Roboto-Regular", 12.5)
		#canvas.drawCentredString(4.15*inch, 10.90 * inch, "Nagar Manmad Highway, Near IBP Petrol Pump,Nimgaon Korhale")
		#canvas.drawCentredString(4.15*inch, 10.60 * inch, "Rahata, Ahmednagar - 423 109")
		#canvas.drawCentredString(4.15*inch, 10.30 * inch, "PHONE : +91-9960867555, 9859121121 EMAIL : vkonconsultants@gmail.com")

		
		#canvas.setFillColorRGB(242,138,27)
		canvas.setFillColorRGB(0.949,0.54,0.105)
		canvas.rect(0*inch, 10.05*inch, 8.27*inch, 0.15*inch ,stroke=0,fill=1)
		canvas.rect(0*inch, 0*inch, 8.27*inch, 0.15*inch ,stroke=0,fill=1)



		canvas.setFont("Roboto-Regular", 12)
		#canvas.drawString(inch, 0.75 * inch, "Page %d %s" % (doc.page, pageinfo))
		canvas.restoreState()



	Title = "Hello world"
	pageinfo = "platypus example"
	
	cname = custname
	address = add
	project = project
	length = length
	
	area = area
	kitchen = kitchen
	bath = bath
	floors = floors
	mailid = mailid
	
	


	SERVICE_ACCOUNT_FILE = 'keys.json'
	SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

	cred = None
	cred = service_account.Credentials.from_service_account_file(
		SERVICE_ACCOUNT_FILE, scopes=SCOPES)
	SAMPLE_SPREADSHEET_ID = '1uogWguNxoRXTY6sKVH6on2DHXwkriFc9RtkoqusF9dc'

	service = build('sheets', 'v4', credentials=cred)

	# Call the Sheets API
	sheet = service.spreadsheets()
	
	databank=[{'range':'ABSTRACT_SHEET!C8:K8','values':[[custname.upper()]]},
				{'range':'ABSTRACT_SHEET!C6:K6','values':[[address.upper()]]},
				{'range':'ABSTRACT_SHEET!O5:Q5','values':[[project.upper()]]},
				{'range':'ABSTRACT_SHEET!O7','values':[[length]]},
				
				{'range':'ABSTRACT_SHEET!Q7','values':[[area]]},
				{'range':'ABSTRACT_SHEET!O8:P8','values':[[kitchen]]},
				{'range':'ABSTRACT_SHEET!O9:P9','values':[[bath]]},
				{'range':'ABSTRACT_SHEET!O10:Q10','values':[[floors.upper()]]}]
	bodydata={'valueInputOption':'USER_ENTERED','data':databank}
	#print(databank)
	#print(bodydata)
	

	#batch_update=sheet.values().batchUpdate(spreadsheetId=SAMPLE_SPREADSHEET_ID,body=bodydata).execute()
	namt = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,range="ABSTRACT_SHEET!K67").execute()
	camt = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,range="ABSTRACT_SHEET!K68").execute()
	eamt = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,range="ABSTRACT_SHEET!K69").execute()
	amt = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,range="ABSTRACT_SHEET!K71").execute()
	inwords = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,range="ABSTRACT_SHEET!C73:K73").execute()
	



	asheet = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
								range="ABSTRACT_SHEET!C8:K8").execute()
	PAGE_HEIGHT=defaultPageSize[1]; PAGE_WIDTH=defaultPageSize[0]
	styles = getSampleStyleSheet()
	styles.add(ParagraphStyle(name='Right', alignment=TA_RIGHT,fontName="Roboto-Regular",fontSize=12))
	styles.add(ParagraphStyle(name='Justify', alignment=TA_JUSTIFY,fontName="Roboto-Regular",fontSize=12))
	styles.add(ParagraphStyle(name='BoldR', alignment=TA_RIGHT,fontName="Roboto-Bold",fontSize=12))
	styles.add(ParagraphStyle(name='BoldD', alignment=TA_RIGHT,fontName="Roboto-Bold",fontSize=12))
	styles.add(ParagraphStyle(name='BoldL', alignment=TA_LEFT,fontName="Roboto-Bold",fontSize=12))


	pdfmetrics.registerFont(TTFont('Roboto-Regular', 'Fonts/Roboto-Regular.ttf'))
	pdfmetrics.registerFont(TTFont('Roboto-Bold', 'Fonts/Roboto-Bold.ttf'))


	pagesize = (8.27 * inch, 11.70 * inch)
	#output = cStringIO.StringIO()
	doc = SimpleDocTemplate("Estimate.pdf",pagesize=A4,rightMargin=.25*inch,leftMargin=0.25*inch,topMargin=1.65*inch,bottomMargin=0.25*inch)
	Story=[]

	width, height = A4
	widthList=[width*0.1,width*.8,width*0.1,]
	heightList=[height*0.12,height*0.05,height*0.05,height*0.05,height*0.15,height*0.05,height*0.05,height*0.28,]
	#pdf = canvas.Canvas("file.pdf",pagesize=A4)
	#canvas.setTitle('Estimate')
	#canvas.setFont("Roboto-Regular", 14)
	#pdf.rect(0.5*inch, 0.5*inch, 7.27*inch, 10.69*inch, fill=0)
	enggname="""SAGAR VIJAYRAO KATORE""" + "<br/><font size=14>CONSULTING ENGINEER, STRUCTURAL DESIGNER <br/>& GOVT. REGISTERED VALUER</font>"""
	enggadd = """At : Nimgaon Korhale, Post : Laxmiwadi, Tal : Rahata,<br/>Ahmednagar - 423 109 <br/>+91- 9960 867 555, 9859 121 121 """
	parastyle = ParagraphStyle('para1')
	parastyle.fontSize = 14
	parastyle.fontName='Roboto-Bold'
	parastyle.alignment=TA_CENTER
	parastyle.leading= 20
	engname=Paragraph(enggname,parastyle)
	parastyle = ParagraphStyle('para1')
	parastyle.alignment=TA_CENTER
	parastyle.fontSize = 14
	parastyle.leading= 20
	engadd =Paragraph(enggadd,parastyle)
	ERNAME=[engname,engadd]
	parastyle = ParagraphStyle('para1')
	parastyle.alignment=TA_CENTER
	parastyle.fontName='Roboto-Bold'
	parastyle.fontSize = 14
	parastyle.leading= 20
	clientname =Paragraph(custname,parastyle)
	CNAME=[clientname]
	parastyle = ParagraphStyle('para1')
	parastyle.alignment=TA_CENTER
	parastyle.fontName='Roboto-Regular'
	parastyle.fontSize = 14
	parastyle.leading= 20
	addrs =Paragraph(address,parastyle)
	CADDRESS=[addrs]
	pwork= "PROPOSED "+ project 
	work= "PROPOSED "+ project + " IN "+ address
	#print(work)



	fpageTable=Table([
		['','ESTIMATE',''],
		['','PROJECT',''],
		['',pwork,''],
		['','IN',''],['',CADDRESS,''],
		['','OF',''],['',CNAME,''],
		['',ERNAME,'']
	],colWidths=widthList,rowHeights=heightList)


	fpageTable.setStyle([
		#('GRID', (0,0), (-1,-1),1,'red'),
		('ALIGN',(0,0), (-1,-1),'CENTER'),
		('VALIGN',(1,0), (1,6),'MIDDLE'),
		('FONT',(1,0),(1,0),'Roboto-Bold'),
		('FONTSIZE',(1,0),(1,0),16),
		('FONT',(1,1),(1,1),'Roboto-Bold'),
		('FONTSIZE',(1,1),(1,1),16),
		('FONT',(1,2),(1,3),'Roboto-Regular'),
		('FONTSIZE',(1,2),(1,3),14),
		('FONT',(1,5),(1,5),'Roboto-Regular'),
		('FONTSIZE',(1,5),(1,5),14),
		('FONT',(1,6),(1,6),'Roboto-Bold'),
		('FONTSIZE',(1,6),(1,6),16),
		('FONT',(1,7),(1,7),'Roboto-Bold'),
		('VALIGN',(1,7), (1,7),'BOTTOM'),
		('FONTSIZE',(1,7),(1,7),12),
		('ALIGN',(1,7), (1,7),'CENTER'),
		('BOTTOMPADDING',(1,7),(1,7),15),

	])
	

	#fpageTable.wrapOn(pdf,0,0)
	#fpageTable.drawOn(pdf,0,0)
	Story.append(fpageTable)


	#########################First Page End
	"""
	file.rect(0.5*inch, 0.5*inch, 7.27*inch, 10.69*inch, fill=0)
	file.setFont("Helvetica-Bold", 20)
	file.drawString(3.00*inch,8.70*inch,"ESTIMATE SUMMARY")
	file.setFont("Helvetica-Bold", 12)
	file.drawCentredString(3*inch,9.70*inch,"NAME OF CLENT")
	file.showPage()
	file.drawImage("img.png",0,0,*A4,kind='proportional' )
	file.setFont("Helvetica-Bold", 20)
	file.drawString(3.00*inch,8.70*inch,"ESTIMATE SUMMARY")
	"""
	#pdf.showPage()
	#pdf.drawImage("LH.png",0,0,*A4)

	#canvas.setFont("Roboto-Regular", 12)
	enggname="""SAGAR VIJAYRAO KATORE""" + "<br/><font size=12>CONSULTING ENGINEER, STRUCTURAL DESIGNER & GOVT. REGISTERED VALUER</font>"""
	enggadd = """At : Nimgaon Korhale, Post : Laxmiwadi, Tal : Rahata,Ahmednagar - 423 109 <br/>+91- 9960 867 555, 9859 121 121 """
	parastyle = ParagraphStyle('para1')
	parastyle.fontSize = 14
	parastyle.fontName='Roboto-Bold'
	parastyle.alignment=TA_CENTER
	parastyle.leading= 20
	engname=Paragraph(enggname,parastyle)
	parastyle = ParagraphStyle('para1')
	parastyle.alignment=TA_CENTER
	parastyle.fontSize = 12
	parastyle.leading= 20
	engadd =Paragraph(enggadd,parastyle)
	ERNAME=[engname,engadd]

	#pdf.setFont("Roboto-Regular", 12)
	widthList=[width*0.10,width*0.80,width*0.10]
	heightList=[height*0.05,height*0.30,height*0.05,height*0.25,height*0.17,]
	spageTable=Table([
		['','ESTIMATE SUMMARY',''],
		['',sftable(widthList[1],heightList[1],custname,work,area,floors),''],
		['','',''],
		['',sstable(widthList[1],heightList[3]),''],
		['','',''],
	],colWidths=widthList,rowHeights=heightList)
	spageTable.setStyle([
		#('GRID', (0,0), (-1,-1),1,'red'),
		('ALIGN',(0,0), (-1,-1),'CENTER'),
		('VALIGN',(0,0), (0,2),'MIDDLE'),
		('VALIGN',(0,0), (0,0),'BOTTOM'),
		('FONT',(1,0),(1,0),'Roboto-Bold'),
		('FONTSIZE',(1,0),(1,0),14),
		('BOTTOMPADDING',(1,0),(1,0),20),
		('FONT',(0,1),(0,1),'Roboto-Bold'),
		('FONTSIZE',(1,1),(1,1),14),
		('FONT',(1,3),(1,3),'Roboto-Bold'),
		('FONTSIZE',(1,3),(1,3),12),

	])
	Story.append(spageTable)

	#spageTable.wrapOn(pdf,0,0)
	#spageTable.drawOn(pdf,0,0)
	#pdf.showPage()
	#################### Third Page
	#pdf.drawImage("LH.png",0,0,*A4)
	#pdf.setFont("Roboto-Regular", 12)
	widthList=[width*0.05,width*0.9,width*0.05]
	heightList=[height*0.01,height*0.07,height*0.01,height*0.025,height*0.01,height*0.025,]
	#height*0.62,height*0.03,
	abstractdata = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,range="ABSTRACT_SHEET!B13:K71").execute()
	
	
	styleH=styles['Normal']
	styleH.fontName="Roboto-Bold"
	styleH.fontSize=12
	styleH.alignment=TA_CENTER
	
	
	
	SNOH=Paragraph('Item no',styleH)
	DESCH=Paragraph('Description',styleH)
	QTYH=Paragraph('Qty',styleH)
	UNITH=Paragraph('Unit',styleH)
	RATEH=Paragraph('Rate',styleH)
	AMTTH=Paragraph('Amount',styleH)


	td=[[SNOH,DESCH,QTYH,UNITH,RATEH,AMTTH]]
	

	dataset= list(abstractdata.values())
	dataset.pop(0)
	dataset.pop(0)
	r=0
	newdata=[]
	data=[]
	for row in dataset:
		newdata.append(row)
		
	styleN=styles['BodyText']
	styleN.alignment=TA_LEFT
	#styleN.alignment=TA_CENTER
	styleN.fontName="Roboto-Regular"
	styleN.fontSize=12

	##########
	#styleR=styles['BodyText']
	#styleR.alignment=TA_LEFT
	#styleR.alignment=TA_RIGHT
	#styleR.fontName="Roboto-Regular"
	#styleR.fontSize=12

	############
	


	for row in newdata:
		for r in row:
			if len(r)==2:
				SNO=Paragraph(r[0],styleN)
				DESC=Paragraph(r[1],styleN)
				

				data=[SNO,DESC,"","","",""]
			elif r[1]=='Total':
				print(r[1])
				SNO=Paragraph(r[0],styleN)
				DESC=Paragraph(r[1],styles['BoldR'])
				QTY=Paragraph(r[6],styleN)
				UNIT=Paragraph(r[7],styleN)
				RATE=Paragraph(r[8],styleN)
				AMTT=Paragraph(r[9], styles['BoldR'])
				data=[SNO,DESC,QTY,UNIT,RATE,AMTT]
			else:
				SNO=Paragraph(r[0],styleN)
				DESC=Paragraph(r[1],styles['Justify'])
				QTY=Paragraph(r[6],styleN)
				UNIT=Paragraph(r[7],styleN)
				RATE=Paragraph(r[8],styleN)
				AMTT=Paragraph(r[9], styles['Right'])
				data=[SNO,DESC,QTY,UNIT,RATE,AMTT]
			td.append(data)
	
	tpageTable=Table([
		['','',''],
		['',tftable(widthList[1],heightList[1],work),''],
		['','',''],
		['',tstable(widthList[1],heightList[3],custname),''],
		['','',''],
		['',tttable(widthList[1],heightList[5]),''],
		#['',tsheet(widthList[1],heightList[9],td),''],
		#['','','']
	],colWidths=widthList,rowHeights=heightList)
	tpageTable.setStyle([
		
		#('GRID', (0,0), (-1,-1),1,'red'),
		('ALIGN',(0,0), (-1,-1),'CENTER'),
		('VALIGN',(0,0), (-1,-1),'MIDDLE'),
		('VALIGN',(0,0), (0,0),'BOTTOM'),
		('FONT',(1,0),(1,0),'Roboto-Bold'),
		('FONTSIZE',(1,0),(1,0),12),
		('FONT',(0,1),(0,1),'Roboto-Bold'),
		('FONTSIZE',(1,1),(1,1),12),
		('FONT',(1,5),(1,5),'Roboto-Bold'),
		('FONTSIZE',(1,5),(1,5),12),
		('FONT',(1,3),(1,3),'Roboto-Bold'),
		('FONTSIZE',(1,3),(1,3),12),

	])
	tbld=Table(td)
	ts=TableStyle([('GRID',(0,0),(-1,-1),2,"red")])
	tbld.setStyle(ts)
	today=date.today()
	dt = today.strftime("%d/%m/%Y")
	#Story.append(Paragraph('REF NO: ',styles['BoldL']))
	Story.append(Paragraph('DATE: '+ str(dt),styles['BoldR']))
	
	Story.append(tpageTable)
	Story.append(Spacer(1,0.05*inch))
	Story.append(tsheet(widthList[1],height*.62,td))
	#print(td)

	#tpageTable.wrapOn(pdf,0,0)
	#tpageTable.drawOn(pdf,0,0)
	#pdf.showPage()



	############################
	#pdf1=SimpleDocTemplate("Binder.pdf")
	#Story.append(tpageTable)
	#pdf1.build(pdf1)
	doc.build(Story, onFirstPage=myFirstPage, onLaterPages=myLaterPages)
	#pdf_out = output.getvalue()
	#output.close()
	#response = make_response(pdf_out)
	#response.headers['Content-Disposition'] = "attachment; filename='test.pdf"
	#response.mimetype = 'application/pdf'
	filename="ESTIMATE FOR " +cname.upper()+ ".pdf"
	return send_file('Estimate.pdf', attachment_filename=filename, as_attachment=False)


	#pdf.save()



	

	


	


if __name__ =="__main__":
	app.run(debug=True)