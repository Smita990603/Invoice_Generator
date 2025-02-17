    
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle,Spacer,Paragraph,Frame,PageTemplate
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle

try:
    
    company_name='Tech Solutions Inc.'
    Company_address='123 Main Street, Anytown, CA 91234'
    #writing to Excel file
    # data = {'Cust_id': [1,1,1,2,2,3,3],
    #         'Product_Id': [1,2,3,1,4,2,3],
    #         'Product Name': ['Laptop', 'Keyboard', 'Mouse', 'Laptop','T-shirt','Keyboard','Mouse'],
    #         'Quantity':[1,2,3,1,2,1,1],
    #         'Unit_price':[1200,75,25,1200,20,75,25]
    #         }
    # df=pd.DataFrame(data)

    # df.to_excel(r"C:\Users\Star Computer\Desktop\Python programs\invoice_data.xlsx",sheet_name='Sheet1',index=False)

    # reading Excel file
    df=pd.read_excel(r"C:\Users\Star Computer\Desktop\Python programs\invoice_data.xlsx",engine='openpyxl')
    df['Total_price']=df['Quantity']*df['Unit_price']#calculating Total price group by cust_id
    df['Tax']=df['Total_price']*(5/100)#calculating Tax group by cust_id
    df['Grand_Total']=df['Total_price']+df['Tax']#calculating Grand Total
    #adding calculations to Excel File
    df.to_excel(r"C:\Users\Star Computer\Desktop\Python programs\invoice_data.xlsx",sheet_name='Sheet1',index=False)
    #group by cust_id
    df1=df.groupby('Cust_id')
    cnt=1#counter for giving invoice number
    for cust_id,group in df1:
        Invoice_number='INV-'+str(pd.to_datetime("today").strftime("%Y-%m-%d"))+'-00'+str(cnt)
        cnt+=1
        Invoice_Date=pd.to_datetime("today").strftime("%Y-%m-%d")
        df2=group[['Product_Id', 'Product Name','Quantity','Unit_price','Total_price']]
        Subtotal=group['Total_price'].sum()
        Tax=group['Tax'].sum()
        Grand_Total=group['Grand_Total'].sum()
        outputpdf='Output'+str(cust_id)
        doc = SimpleDocTemplate(rf"C:\Users\Star Computer\Desktop\Python programs\Invoice_generated\{outputpdf}.pdf", pagesize=letter)
        elements=[]
        # Margins and dimensions (using mm for example, adjust as needed):
        margin = 20 * mm
        width, height = letter
        content_width = width - 2 * margin
        content_height = height - 2 * margin
        frame = Frame(margin, margin, content_width, content_height,  # x, y, width, height
                  leftPadding=4, bottomPadding=0, rightPadding=0, topPadding=0,
                  showBoundary=1)  # Show the border
        page_template = PageTemplate(frames=[frame])
        doc.addPageTemplates(page_template)
        styles = getSampleStyleSheet() #Get the sample stylesheet
        centered_style = ParagraphStyle(name='CenteredStyle',parent=styles['Normal'],alignment=1,fontName='Helvetica-Bold',fontSize=17)
        left_style = ParagraphStyle(name='LeftStyle',parent=styles['Normal'],alignment=0,fontName='Helvetica',fontSize=12)
        elements.append(Paragraph("Invoice",centered_style)) #Add heading to the pdf
        elements.append(Spacer(1,20))
        elements.append(Paragraph(f"<b>Company Name: </b>{company_name}",left_style)) #Add company name
        elements.append(Spacer(1,10))
        elements.append(Paragraph(f"<b>Company Address: </b>{Company_address}",left_style)) #Add company address
        elements.append(Spacer(1,10))
        elements.append(Paragraph(f"<b>Invoice Number: </b>{Invoice_number}", left_style)) #Add invoice number
        elements.append(Spacer(1,10))
        elements.append(Paragraph(f"<b>Invoice Date: </b>{Invoice_Date}",left_style)) #Add invoice date
        elements.append(Spacer(1,10))
        elements.append(Paragraph(f"<b>Customer ID: </b>{str(cust_id)}", left_style)) #Add customer id
        elements.append(Spacer(1,10))
        elements.append(Paragraph("<b><font size='12'>Product Details:</font></b>", centered_style))  # Product details label
        
        
        style = TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),  # Header background
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),  # Header text color
    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),  # Center header text
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),  # Bold header font
    ('FONTSIZE', (0, 0), (-1, 0), 12),  # Header font size
    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),  # Header bottom padding
    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),  # Data row background
    ('GRID', (0, 0), (-1, -1), 1, colors.black), # Table border
    ('ALIGN', (0, 1), (-1, -1), 'LEFT'), # Align data to the left
    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),  # Data font
    ('FONTSIZE', (0, 1), (-1, -1), 10),  # Data font size
    ('BOTTOMPADDING', (0, 1), (-1, -1), 6),  # Data bottom padding
    ('TOPPADDING', (0, 1), (-1, -1), 6),  # Data top padding
])
        list_of_lists = [df2.columns.tolist()]+df2.values.tolist()
        table=Table(list_of_lists)
        table.setStyle(style)
        elements.append(Spacer(1,40))
        elements.append(table)
        elements.append(Spacer(1,40))
        elements.append(Paragraph(f"<b>Subtotal: </b>{str(Subtotal)}" ,left_style)) 
        elements.append(Spacer(1,10))
        elements.append(Paragraph(f"<b>Tax (5%): </b>{str(Tax)}", left_style))
        elements.append(Spacer(1,10))
        elements.append(Paragraph(f"<b>Grand Total: </b>{str(Grand_Total)}", left_style))
        doc.build(elements)
        
    print("Check pdf in respected folder:)")
except Exception as e:
    print(e)


