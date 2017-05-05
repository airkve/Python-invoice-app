#! python3
# invoiceSaladino.py - Module that generate an invoice with ReportLab.
# Author: Richard Jimenez

import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# Tamaño de la pagina tamaño carta en puntos 612x792

def createInvoice(date, client, items, payment, ship):
    """Function to generate the invoice."""
    invoice = canvas.Canvas('Factura' + str(datetime.datetime.now().strftime("%d-%m-%Y %H.%M.%S")) + '.pdf', pagesize=(612, 396))  # Tamaño media pagina
    pdfmetrics.registerFont(TTFont('Calibri', 'calibri.ttf'))
    pdfmetrics.registerFont(TTFont('CalibriB', 'calibrib.ttf'))
    formDate   = date  # datetime.date.today().strftime("%d   %m     %Y")
    clientName = client[0]
    clientAddr = client[1]
    clientTelf = client[2]
    clientRIF  = client[3]
    condOfPay  = payment[1]
    guideShip  = ship
    sortItems  = items[0][0:len(items[0])]
    cash       = ''
    check      = ''
    cardCredit = ''
    cardDebit  = ''
    if 'Efectivo' in payment:
        cash = 'X'
    elif 'Cheque' in payment:
        check = 'X'
    elif 'Tarjeta de credito' in payment:
        cardCredit = 'X'
    elif 'tarjeta de debito' in payment:
        cardDebit  = 'X'

    checkNum   = payment[3][0]
    checkBank  = payment[3][1]
    qty        = list(map(int, items[1]))
    priceUnit  = list(map(int, items[2]))
    bolivares  = []
    # Calculo de la factura
    for p, q in zip(priceUnit, qty):
        resultBs = float("{0:.2f}".format(p)) * float("{0:.2f}".format(q))
        bolivares.append(resultBs)
    subTotal   = sum(bolivares)
    if type(payment[2]) == int:
        tax = payment[2] / 100
    else:
        tax = float(0.12)

    taxTotal   = subTotal * tax
    total      = subTotal + taxTotal


    # Fecha
    invoice.setFont('Calibri', 20)
    invoice.drawString(485, 317, formDate)
    # Datos del cliente
    invoice.setFont('Calibri', 10)
    # Nombre del cliente
    invoice.drawString(104, 265, clientName)
    # Direccion del cliente
    invoice.drawString(100, 245, clientAddr)
    # Telefono del cliente
    invoice.drawString(68, 225, clientTelf)
    # RIF del cliente
    invoice.drawString(207, 225, clientRIF)
    # Condicion del cliente
    invoice.drawString(365, 225, condOfPay)
    # Guia de despacho
    invoice.drawString(546, 225, guideShip)
    # Items
    invoice.setFont('Calibri', 10)

    # Fill the invoice with the items and amounts
    posY = 182
    lineH = 15
    y = 0
    for lines in range(0, len(sortItems)):
        if y < posY:
            y = posY
            invoice.drawString(31, y, sortItems[lines])
            invoice.drawCentredString(417, y, str(qty[lines]))
            priceUnit_formated = "{0:.2f}".format(priceUnit[lines])
            invoice.drawRightString(507, y,
                                    priceUnit_formated.replace('.', ','))
            bs_formated = "{0:.2f}".format(bolivares[lines])
            invoice.drawRightString(580, y, bs_formated.replace('.', ','))
        else:
            y -= lineH
            invoice.drawString(31, y, sortItems[lines])
            invoice.drawCentredString(417, y, str(qty[lines]))
            priceUnit_formated = "{0:.2f}".format(priceUnit[lines])
            invoice.drawRightString(507, y,
                                    priceUnit_formated.replace('.', ','))
            bs_formated = "{0:.2f}".format(bolivares[lines])
            invoice.drawRightString(580, y, bs_formated.replace('.', ','))

    # Subtotal, iva y total
    invoice.setFont('CalibriB', 10)
    # Subtotal
    subtotal_formated = "{0:.2f}".format(subTotal)
    invoice.drawRightString(580, 41, subtotal_formated.replace('.', ','))
    # IVA
    iva_formated = "{0:.2f}".format(taxTotal)
    invoice.drawRightString(490, 25, '12')
    invoice.drawRightString(580, 25, iva_formated.replace('.', ','))
    # Total
    total_formated = "{0:.2f}".format(total)
    invoice.drawRightString(580, 9, total_formated.replace('.', ','))
    # Forma de pago
    invoice.setFont('Calibri', 9)
    # Numero de cheque
    invoice.drawString(80, 27, checkNum)
    # Banco
    invoice.drawString(61, 12, checkBank)
    # Marca de Efectivo
    invoice.drawString(111, 41, cash)
    # Marca de Cheque
    invoice.drawString(167, 41, check)
    # Marca de Tarjeta de Credito
    invoice.drawString(229, 41, cardCredit)
    # Marca de Tarjeta de Debito
    invoice.drawString(324, 41, cardDebit)

    invoice.showPage()
    invoice.save()

date = '12 08 1973'
client = ['Nombre', 'Ctra Vieja de Las Minas entre Callejón San Tomás y Sta Rosalia.Edif San Benito.PB.Local 1.Las Minas de Baruta', '555-5555', 'J-123456789-0']
items = [['Tomates pelados', 'pure de tomate'], ['10', '5'], ['7800', '6000'], ['78000', '30000']]
payment = ['Efectivo', 'De contado', '', ['', '']]
ship = '00000'
createInvoice(date, client, items, payment, ship)
