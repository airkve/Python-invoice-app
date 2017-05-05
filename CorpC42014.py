#! python3
# CorpC42014.py (V2.0)- Sistema de facturas para Corporación C4-2014, C.A.
# Autor: Richard Jimenez

import shelve
import invoiceSaladino
import os
import win32api
import win32print
import subprocess
import platform
import datetime
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import messagebox

client = []
items = [[], [], [], []]
payment = []


class MainMenu:
    """Doc."""

    def __init__(self, master):
        """Doc."""
        self.master = master
        master.title('Corporación C4-2014 - Menu Principal')
        # Icon
        if platform.system() == 'Linux':
            pass
        else:
            self.master.wm_iconbitmap('Tomato.ico')

        self.imgFrame = tk.Frame(self.master)
        self.imgFrame.grid(row=0)
        self.bannerImg = tk.PhotoImage(file='banner.gif')
        self.bannerLab = tk.Label(self.imgFrame, image=self.bannerImg)
        self.bannerLab.grid(row=0)
        self.frame = tk.Frame(self.master)
        self.frame.grid(row=1)
        self.var1 = tk.StringVar()
        self.labComp = ttk.Label(self.frame, text='Empresa:')
        self.labComp.grid(row=1, column=1, padx=5, pady=5)
        self.dropbox = ttk.OptionMenu(self.frame, self.var1,
                                      'Seleccione', 'Corporacion C4-2014',
                                      'Compañia 2', 'Compañia 3')
        self.dropbox.grid(row=2, column=1, padx=5, pady=5)
        self.button1 = ttk.Button(self.frame, text='Crear Factura',
                                  width=25, command=self.openInvoiceMenu)
        self.button1.grid(row=3, column=0, padx=5, pady=5)
        self.button2 = ttk.Button(self.frame, text='Imprimir Factura',
                                  width=25, command=self.openPrintMenu)
        self.button2.grid(row=3, column=1, padx=5, pady=5)
        self.button3 = ttk.Button(self.frame, text='Salir',
                                  width=25, command=self.master.destroy)
        self.button3.grid(row=3, column=2, padx=5, pady=5)



    def openInvoiceMenu(self):
        self.invoiceMenu = tk.Toplevel(self.master)
        self.app = InvoiceMenu(self.invoiceMenu)


    def openPrintMenu(self):
        self.printMenu = tk.Toplevel(self.master)
        self.app = PrintMenu(self.printMenu)

    def exitMenu(self):
        InvoiceMenu.dbClient.close()
        self.master.destroy()

class InvoiceMenu:
    """ Docstring for InvoiceMenu"""

    def __init__(self, master):
        self.clients = []
        self.dbClient = shelve.open('clients.db', writeback=True)
        self.clients = list(self.dbClient.values())
        self.master = master
        master.title('Corporación C4-2014 - Carga de Factura')
        # Icon
        if platform.system() == 'Linux':
            pass
        else:
            self.master.wm_iconbitmap('Tomato.ico')
        # First frame
        self.labFrame1 = ttk.LabelFrame(self.master, text='Datos del Cliente')
        self.labFrame1.grid(padx=5, pady=5)
        # Client information
        self.labDateDC = tk.Label(self.labFrame1, text='Fecha:')
        self.labDateDC.grid(row=0, column=0, sticky='E', padx=5, pady=5)
        self.varEntrDC = tk.StringVar()
        self.varEntrDC.set(datetime.date.today().strftime("%d-%m-%Y"))
        self.entDateDC = tk.Entry(self.labFrame1, width=10, state='readonly',
                                  textvariable=self.varEntrDC)
        self.entDateDC.grid(row=0, column=1, sticky='W', padx=5, pady=5)
        self.dashpos = [2, 5]
        self.entDateDC.bind('<Key>', self.formatDate)
        self.dateVar = tk.IntVar()
        self.dateBtnDC = ttk.Checkbutton(self.labFrame1, text='Fecha manual',
                                         variable=self.dateVar,
                                         command=self.udpEntDate)
        self.dateBtnDC.grid(row=0, column=2, sticky='W')
        self.labNameDC = ttk.Label(self.labFrame1,
                                   text='Nombre o Razón social:')
        self.labNameDC.grid(row=1, column=0, columnspan=2, sticky='E', padx=5,
                            pady=5)
        self.entNameDC = ttk.Combobox(self.labFrame1, width=57)
        self.entNameDC['values'] = list(x[0] for x in self.clients)
        self.entNameDC.bind('<<ComboboxSelected>>', self.udpClientFields)
        self.entNameDC.grid(row=1, column=1, sticky='W', padx=5, pady=5)

        self.labAddrDC = ttk.Label(self.labFrame1, text='Domicilio fiscal:')
        self.labAddrDC.grid(row=2, column=0, sticky='E', padx=5, pady=5)
        self.entAddrDC = ttk.Entry(self.labFrame1, width=60)
        self.entAddrDC.grid(row=2, column=1, sticky='W', padx=5, pady=5)

        self.labTelfDC = ttk.Label(self.labFrame1, text='Teléfono:')
        self.labTelfDC.grid(row=2, column=2, sticky='E', padx=5, pady=5)
        self.entTelfDC = ttk.Entry(self.labFrame1, width=15)
        self.entTelfDC.grid(row=2, column=3, sticky='W', padx=5, pady=5)

        self.labRif_DC = ttk.Label(self.labFrame1, text='RIF:')
        self.labRif_DC.grid(row=2, column=4, sticky='E', padx=5, pady=5)
        self.entRif_DC = ttk.Entry(self.labFrame1, width=13)
        self.entRif_DC.grid(row=2, column=5, sticky='W', padx=5, pady=5)

        self.labMailDC = ttk.Label(self.labFrame1, text='Correo:')
        self.labMailDC.grid(row=1, column=2, sticky='E', padx=5, pady=5)
        self.entMailDC = ttk.Entry(self.labFrame1, width=37)
        self.entMailDC.grid(row=1, column=3, columnspan=3, sticky='W',
                            padx=5, pady=5)
        # Client buttons
        self.btnSaveDC = ttk.Button(self.labFrame1, text='Guardar cliente',
                                    command=self.saveClient)
        self.btnSaveDC.grid(row=3, column=0, padx=5, pady=5)
        self.btnClr_DC = ttk.Button(self.labFrame1, text='Limpiar',
                                    command=self.clearClient)
        self.btnClr_DC.grid(row=3, column=1, padx=5, pady=5)
        self.btnClr_DC = ttk.Button(self.labFrame1, text='Eliminar',
                                    command=self.delClient)
        self.btnClr_DC.grid(row=3, column=2, padx=5, pady=5)
        # Second frame
        self.labFrame2 = ttk.LabelFrame(self.master,
                                        text='Descripción y Montos')
        self.labFrame2.grid(padx=5, pady=5)
        # Items and cost information
        self.labItemDM = ttk.Label(self.labFrame2,
                                   text='Comcepto o descripción:')
        self.labItemDM.grid(row=1, column=0, sticky='E', padx=5, pady=5)
        self.entItemDM = ttk.Entry(self.labFrame2, width=60)
        self.entItemDM.grid(row=1, column=1, sticky='W', padx=5, pady=5)

        self.labQntyDM = ttk.Label(self.labFrame2, text='Cantidad:')
        self.labQntyDM.grid(row=1, column=2, sticky='E', padx=5, pady=5)
        self.entQntyDM = ttk.Entry(self.labFrame2, width=4)
        self.entQntyDM.grid(row=1, column=3, sticky='W', padx=5, pady=5)

        self.labPricDM = ttk.Label(self.labFrame2, text='Precio unitario:')
        self.labPricDM.grid(row=1, column=4, sticky='E', padx=5, pady=5)
        self.entPricDM = ttk.Entry(self.labFrame2, width=12)
        self.entPricDM.grid(row=1, column=5, sticky='W', padx=5, pady=5)
        self.entPricDM.bind('<Return>', self.insertData)
        # Add items
        self.btnRegiDM = ttk.Button(self.labFrame2, text='Agregar',
                                    command=self.insertData)
        self.btnRegiDM.grid(row=2, column=0, padx=5, pady=5)

        self.separatDM = ttk.Separator()
        self.separatDM.grid(row=3)

        self.listView = ttk.Treeview(self.labFrame2)
        self.listView['show'] = 'headings'
        self.listView['columns'] = ('items', 'qty', 'unitPrice', 'bolivares')
        self.listView.column('items', width=312)
        self.listView.column('qty', width=70, anchor='center')
        self.listView.column('unitPrice', width=160, anchor='e')
        self.listView.column('bolivares', width=160, anchor='e')
        self.listView.heading('items', text='Concepto o Descripcíon')
        self.listView.heading('qty', text='Cantidad')
        self.listView.heading('unitPrice', text='Precio Unitario')
        self.listView.heading('bolivares', text='Bolívares')
        self.listView.grid(columnspan=10, padx=5, pady=5)
        ttk.Style().configure("Treeview.Heading", foreground="blue",
                              background='light gray')
        # Third frame
        self.labFrame3 = ttk.LabelFrame(self.master, text='Condiciones de pago')
        self.labFrame3.grid(padx=5, pady=5)
        # payment methods
        self.labPayCCP = ttk.Label(self.labFrame3, text='Forma de pago:')
        self.labPayCCP.grid(row=1, column=0, sticky='E', padx=5, pady=5)

        self.chkBx1 = tk.IntVar()
        self.chkBx1.set(1)
        self.chkBx2 = tk.IntVar()
        self.chkBx3 = tk.IntVar()
        self.chkBx4 = tk.IntVar()
        self.chkCashCP = ttk.Checkbutton(self.labFrame3,
                                         text='Efectivo',
                                         variable=self.chkBx1,
                                         command=self.udpCashBtn)
        self.chkCashCP.grid(row=1, column=1, padx=5, pady=5)
        self.chkChecCP = ttk.Checkbutton(self.labFrame3,
                                         text='Cheque',
                                         variable=self.chkBx2,
                                         command=self.udpChekBtn)
        self.chkChecCP.grid(row=1, column=2, padx=5, pady=5)
        self.chkCredCP = ttk.Checkbutton(self.labFrame3,
                                         text='Tarjeta de crédito',
                                         variable=self.chkBx3,
                                         command=self.udpCredBtn)
        self.chkCredCP.grid(row=1, column=3, padx=5, pady=5)
        self.chkDebtCP = ttk.Checkbutton(self.labFrame3,
                                         text='Tarjeta de débito',
                                         variable=self.chkBx4,
                                         command=self.udpDebtBtn)
        self.chkDebtCP.grid(row=1, column=4, padx=5, pady=5)

        self.labBankCP = ttk.Label(self.labFrame3, text='Banco:')
        self.labBankCP.grid(row=2, column=0, sticky='E', padx=5, pady=5)
        self.entBankCP = ttk.Entry(self.labFrame3, width=30, state='disabled')
        self.entBankCP.grid(row=2, column=1, columnspan=2, sticky='W',
                            padx=5, pady=5)

        self.labChkNCP = ttk.Label(self.labFrame3, text='Cheque No.')
        self.labChkNCP.grid(row=2, column=3, sticky='E', padx=5, pady=5)
        self.entChkNCP = ttk.Entry(self.labFrame3, state='disabled')
        self.entChkNCP.grid(row=2, column=4, sticky='W', padx=5, pady=5)

        self.labPayWCP = ttk.Label(self.labFrame3, text='Condición de pago:')
        self.labPayWCP.grid(row=3, column=0, sticky='E', padx=5, pady=5)
        self.entPayWCP = ttk.Entry(self.labFrame3, width=30)
        self.entPayWCP.grid(row=3, column=1, sticky='W', padx=5, pady=2)

        self.labGuideShip = ttk.Label(self.labFrame3, text='Guia de despacho:')
        self.labGuideShip.grid(row=3, column=2, sticky='E', padx=5, pady=5)
        self.entGuideShip = ttk.Entry(self.labFrame3, width=30)
        self.entGuideShip.grid(row=3, column=3, sticky='W', padx=5, pady=2)

        self.btnSaveFc = ttk.Button(self.labFrame3, text='Crear factura',
                                    width=20, command=self.saveInvoice)
        self.btnSaveFc.grid(row=3, column=6, padx=5, pady=5)

    def saveClient(self):
        self.clientCache = [self.entNameDC.get(), self.entAddrDC.get(),
                            self.entTelfDC.get(), self.entRif_DC.get(),
                            self.entMailDC.get()]
        self.dbClient[str(len(self.dbClient) + 1)] = self.clientCache
        self.clients = list(self.dbClient.values())
        self.entNameDC['values'] = list(x[0] for x in self.clients)
        self.btnSaveDC.state(['disabled'])
        self.dsblFieldsDC()
        tk.messagebox.showinfo(parent=self.master, title='Mensaje',
                               message='Cliente guardado.')

    def udpClientFields(self, event):
        self.enblFieldsDC()
        self.btnSaveDC.state(['disabled'])
        self.getComBox = self.entNameDC.get()
        self.clrFieldsDC()
        for value in self.dbClient.values():
            if self.getComBox in value:
                self.entAddrDC.insert(0, value[1])
                self.entTelfDC.insert(0, value[2])
                self.entRif_DC.insert(0, value[3])
                self.entMailDC.insert(0, value[4])
        self.dsblFieldsDC()
        self.entItemDM.focus()

    def clearClient(self):
        self.enblFieldsDC()
        self.entNameDC.delete(0, tk.END)
        self.btnSaveDC.state(['!disabled'])
        self.clrFieldsDC()
        self.entNameDC.focus()

    def delClient(self):
        self.getComBox = self.entNameDC.get()
        for keys, values in self.dbClient.items():
            if self.getComBox in values:
                    del self.dbClient[keys]
        self.clients = list(self.dbClient.values())
        self.enblFieldsDC()
        self.entNameDC['values'] = list(x[0] for x in self.clients)
        self.entNameDC.delete(0, tk.END)
        self.clrFieldsDC()
        self.btnSaveDC.state(['!disabled'])
        tk.messagebox.showinfo(parent=self.master, title='Mensaje',
                               message='Cliente Eliminado.')
        self.entNameDC.focus()

    def insertData(self, event):
        self.get1 = self.entItemDM.get()
        self.get2 = self.entQntyDM.get()
        self.get3 = self.entPricDM.get()
        self.get4 = int(self.get2) * int(self.get3)
        items[0].append(self.get1)
        items[1].append(self.get2)
        items[2].append(self.get3)
        items[3].append(str(self.get4))
        self.listView.insert("", 0, text="", values=(self.get1, self.get2,
                                                     self.get3, self.get4))
        self.entItemDM.delete(0, tk.END)
        self.entQntyDM.delete(0, tk.END)
        self.entPricDM.delete(0, tk.END)
        self.entItemDM.focus()

    def clrFieldsDC(self):
        self.entAddrDC.delete(0, tk.END)
        self.entTelfDC.delete(0, tk.END)
        self.entRif_DC.delete(0, tk.END)
        self.entMailDC.delete(0, tk.END)

    def enblFieldsDC(self):
        self.entAddrDC.configure(state='NORMAL')
        self.entTelfDC.configure(state='NORMAL')
        self.entRif_DC.configure(state='NORMAL')
        self.entMailDC.configure(state='NORMAL')

    def dsblFieldsDC(self):
        self.entAddrDC.configure(state='readonly')
        self.entTelfDC.configure(state='readonly')
        self.entRif_DC.configure(state='readonly')
        self.entMailDC.configure(state='readonly')

    def udpCashBtn(self):
        if self.chkBx1.get() == 1:
            self.chkChecCP.configure(state='disabled')
            self.chkCredCP.configure(state='disabled')
            self.chkDebtCP.configure(state='disabled')
        else:
            self.chkChecCP.configure(state='normal')
            self.chkCredCP.configure(state='normal')
            self.chkDebtCP.configure(state='normal')

    def udpChekBtn(self):
        if self.chkBx2.get() == 1:
            self.entBankCP.configure(state='normal')
            self.entChkNCP.configure(state='normal')
            self.chkCashCP.configure(state='disabled')
            self.chkCredCP.configure(state='disabled')
            self.chkDebtCP.configure(state='disabled')
        else:
            self.entBankCP.configure(state='disabled')
            self.entChkNCP.configure(state='disabled')
            self.chkCashCP.configure(state='normal')
            self.chkCredCP.configure(state='normal')
            self.chkDebtCP.configure(state='normal')

    def udpCredBtn(self):
        if self.chkBx3.get() == 1:
            self.chkChecCP.configure(state='disabled')
            self.chkCashCP.configure(state='disabled')
            self.chkDebtCP.configure(state='disabled')
        else:
            self.chkChecCP.configure(state='normal')
            self.chkCashCP.configure(state='normal')
            self.chkDebtCP.configure(state='normal')

    def udpDebtBtn(self):
        if self.chkBx4.get() == 1:
            self.chkChecCP.configure(state='disabled')
            self.chkCredCP.configure(state='disabled')
            self.chkCashCP.configure(state='disabled')
        else:
            self.chkChecCP.configure(state='normal')
            self.chkCredCP.configure(state='normal')
            self.chkCashCP.configure(state='normal')

    def udpEntDate(self):
        if self.dateVar.get() == 1:
            self.entDateDC.configure(state='normal')
            self.entDateDC.delete(0, tk.END)
            self.entDateDC.focus()
        else:
            self.entDateDC.delete(0, tk.END)
            self.varEntrDC.set(datetime.date.today().strftime("%d-%m-%Y"))
            self.entDateDC.configure(state='readonly')

    def formatDate(self, event):
        entrylist = [c for c in self.varEntrDC.get() if c != '-']
        for pos in self.dashpos:
            if len(entrylist) > pos:
                entrylist.insert(pos, '-')
        self.varEntrDC.set(''.join(entrylist))
        # Controlling cursor
        cursorpos = self.entDateDC.index(tk.INSERT)
        for pos in self.dashpos:
            if cursorpos == (pos + 1):
                cursorpos += 1
        if event.keysym not in ['BackSpace', 'Right', 'Left', 'Up', 'Down']:
            self.entDateDC.icursor(cursorpos)

    def saveInvoice(self):
        self.entTax_CP = ''
        client.append(self.entNameDC.get())
        client.append(self.entAddrDC.get())
        client.append(self.entTelfDC.get())
        client.append(self.entRif_DC.get())
        client.append(self.entMailDC.get())
        # ship = self.entGuideShip.get()
        if self.chkBx1.get() == 1:
            payment.append('Efectivo')
            payment.append(self.entPayWCP.get())
            payment.append(self.entGuideShip.get())
            payment.append([self.entChkNCP.get(), self.entBankCP.get()])
        elif self.chkBx2.get() == 1:
            payment.append('Cheque')
            payment.append(self.entPayWCP.get())
            payment.appendself.entGuideShip.get())
            payment.append([self.entChkNCP.get(), self.entBankCP.get()])
        elif self.chkBx3.get() == 1:
            payment.append('Tarjeta de credito')
            payment.append(self.entPayWCP.get())
            payment.append(self.entGuideShip.get())
            payment.append([self.entChkNCP.get(), self.entBankCP.get()])
        elif self.chkBx4.get() == 1:
            payment.append('Tarjeta de debito')
            payment.append(self.entPayWCP.get())
            payment.append(self.entGuideShip.get())
            payment.append([self.entChkNCP.get(), self.entBankCP.get()])
        self.entDateDC.configure(state='readonly')
        self.manualDate = self.entDateDC.get()
        date = '{day:5}{month:7}{year:4}'.format(day=self.manualDate[0:2],
                                                 month=self.manualDate[3:5],
                                                 year=self.manualDate[6:10])
        invoiceSaladino.createInvoice(date, client, items, payment)
        self.master.destroy()

class PrintMenu:

    def __init__(self, master):
        self.pdfFiles = []
        self.master = master
        master.title('Corporación C4-2014 - Menu de Impresión')
        # Icon
        if platform.system() == 'Linux':
            pass
        else:
            self.master.wm_iconbitmap('Tomato.ico')

        for file in os.listdir():
            if file.endswith('.pdf'):
                self.pdfFiles.append(file)

        self.labFramePrt = ttk.LabelFrame(self.master,
                                          text='Lista de Facturas')
        self.labFramePrt.pack(side='top', padx=5, pady=5)

        scrollBar = ttk.Scrollbar(self.labFramePrt)
        scrollBar.pack(side='right', fill='y')
        self.listBox = tk.Listbox(self.labFramePrt, selectmode='single')
        self.listBox.pack(side='left', fill='both')
        self.listBox.config(width=60)
        scrollBar.config(command=self.listBox.yview)
        self.listBox.config(yscrollcommand=scrollBar.set)
        self.pdfFiles.sort()
        for item in self.pdfFiles:
            self.listBox.insert(tk.END, item)

        self.btnPrint = ttk.Button(self.master, text='Imprimir',
                                   command=self.printInvoice)
        self.btnPrint.pack(side='right', padx=5, pady=5)

        self.btnOpen = ttk.Button(self.master, text='Abrir',
                                  command=self.openInvoice)
        self.btnOpen.pack(side='left', padx=5, pady=5)

    def printInvoice(self):
        self.getPdf = self.listBox.get('active')
        win32api.ShellExecute(
            0,
            'print',
            self.getPdf,
            #
            '/d:"%s"' % win32print.GetDefaultPrinter(),
            '.',
            5
            )

    def openInvoice(self):
        self.getPdf = self.listBox.get('active')
        self.openPdf = subprocess.Popen(self.getPdf, shell=True)

def main():
    window = tk.Tk()
    app = MainMenu(window)
    window.mainloop()

if __name__ == '__main__':
    main()
