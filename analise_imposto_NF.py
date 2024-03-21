import sys
from PySide6.QtWidgets import QApplication, QWidget, QGridLayout, QPushButton, QFileDialog, QTextEdit, QMessageBox, QTableWidget, QTableWidgetItem, QGroupBox
from PySide6.QtCore import Qt
from PySide6.QtPrintSupport import QPrinter, QPrintDialog
from PySide6.QtGui import QTextCursor, QIcon
import xml.etree.ElementTree as ET
from xml.dom import minidom
import os
import pandas as pd
import numpy as np
import subprocess


try:
    from ctypes import windll
    myappid = 'mycompany.myproduct.subproduct.version'
    windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
except ImportError:
    pass

class XMLAnalyzer(QWidget):
    def __init__(self):
        super().__init__()
        
        self.initUI()
        



    def initUI(self):
        layout = QGridLayout()
        self.setLayout(layout)

        # Group Box for buttons
        group_box = QGroupBox("MENU")
        group_box.setStyleSheet("QGroupBox:title {"
                                "subcontrol-origin: margin;"
                                "subcontrol-position: top center;"
                                "padding-left: 10px;"
                                "padding-right: 10px; }")
        bpx_btns = QGridLayout(group_box)

        # Create and add widgets to the layout
        self.btn_select_xml = QPushButton('Selecionar Arquivo XML')
        self.btn_select_xml.clicked.connect(self.select_xml_file)
        bpx_btns.addWidget(self.btn_select_xml, 0, 0)

        self.btn_run_analysis = QPushButton('Executar Análise')
        self.btn_run_analysis.clicked.connect(self.run_analysis)
        bpx_btns.addWidget(self.btn_run_analysis, 1, 0)

        self.btn_edit_table = QPushButton('Editar Tabela NCM')
        self.btn_edit_table.clicked.connect(self.edit_ncm_table)
        bpx_btns.addWidget(self.btn_edit_table, 6, 0)

        self.print_button = QPushButton("Imprimir relatorio", self)
        self.print_button.clicked.connect(self.printTextEdit)
        bpx_btns.addWidget(self.print_button, 3, 0)

        self.excelButton_tb = QPushButton("Abrir no EXCEL")
        self.excelButton_tb.clicked.connect(self.open_in_excel)
        bpx_btns.addWidget(self.excelButton_tb, 4, 0)

        group_box.setLayout(bpx_btns)  # Set the layout for the group box

        # Add group box and other widgets to the main layout
        layout.addWidget(group_box, 0, 0)
        self.tableWidget = QTableWidget()
        layout.addWidget(self.tableWidget, 2, 1, 1, 6)

        self.tableWidget2 = QTableWidget()
        layout.addWidget(self.tableWidget2, 1, 1, 1, 6)

        self.tableWidget3 = QTableWidget()
        layout.addWidget(self.tableWidget3, 0, 1, 1, 6)

        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)
        #layout.addWidget(self.output_text, 2, 1, 1, 2)

        self.setWindowTitle("Fisco XML")
        janela_icon = QIcon('ICO_FISCO.ico')
        self.setWindowIcon(janela_icon)
        self.showMaximized()

    def printTextEdit(self):
        printer = QPrinter(QPrinter.HighResolution)
        dialog = QPrintDialog(printer, self)
        if dialog.exec() == QPrintDialog.Accepted:
            self.output_text.print_(printer)

    def select_xml_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo XML", "", "XML Files (*.xml)")
        if file_name:
            self.xml_file_path = file_name

    def print_to_output(self, text):
        self.output_text.moveCursor(QTextCursor.End)
        self.output_text.insertPlainText(text)

    def open_in_excel(self):
        try:
            df1 = self.df_forn
            df2 = self.df_ncms
            df3 = self.df_products
            name = self.forn_name + '_nf_'+ self.number_nf 


            file_path, _ = QFileDialog.getSaveFileName(self, 'Salvar Excel', name, 'Excel Files (*.xlsx)')
            

            if file_path:
                with pd.ExcelWriter(file_path) as writer:
                    df1.to_excel(writer, sheet_name='Info NFe', index=False)
                    df2.to_excel(writer, sheet_name='NCMs compiladas', index=False)
                    df3.to_excel(writer, sheet_name='Produtos', index=False)

                # Abrir o Excel automaticamente após o salvamento
                if os.path.isfile(file_path):
                    subprocess.Popen(['start', 'excel', file_path], shell=True)
        except:
            QMessageBox.warning(self, 'ATENÇÃO', 'Realize a análise antes.')


    def run_analysis(self):
        if hasattr(self, 'xml_file_path'):
            self.output_text.clear()
            self.analyze_xml(self.xml_file_path)
        else:
            QMessageBox.warning(self, 'Aviso', 'Por favor, selecione um arquivo XML primeiro.')

    def analyze_xml(self, xml_path):
        os.system('cls')  # Limpa a tela do console
        # Methods xml - math
        self.df_products = pd.DataFrame()
        self.df_products = pd.DataFrame(columns=['codigo', 'ncm', 'origem', 'cst', 'cfop', 'valor_total', 'valor_ipi', 'status', 'mva_uso', 'valor_st',]) #'difal'])
        def extract_data(xml_string, dad_tag, tags):
            result = []
            with open(xml_string, 'r') as f:
                dom = minidom.parse(f)
                items = dom.getElementsByTagName(dad_tag)

                for i, item in enumerate(items):
                    item_data = {'item': i + 1}
                    for tag in tags:
                        if item.getElementsByTagName(tag):
                            tag_value = item.getElementsByTagName(tag)[0].firstChild.data
                            item_data[tag] = tag_value
                            if tag == 'orig':
                                ori_prod = origem(tag_value)
                                item_data['desc_origem'] = ori_prod
                            if tag == 'CFOP':
                                status_cfop = statusCFOP(tag_value)
                                item_data['status_imp'] = status_cfop
                        else:
                            item_data[tag] = 0

                    result.append(item_data)

            return result

        def origem(orig):
            if orig in ['0', '4', '5']:
                ori_p = 'Nacional'
            elif orig in ['1', '2', '5', '6', '7']:
                ori_p = 'Importada'
            elif orig == '3':
                ori_p = 'Nacional+imp40'
            else:
                ori_p = 'Desconhecido'
            return ori_p

        def statusCFOP(cfop):
            if cfop in ['5101', '5102', '6101', '6102','6152']:
                status = 'Analise'  # cfop - analise de ncm
            elif cfop in ['5401', '5403', '6401', '6403', '6404']:
                status = 'Destacado'  # cfop imp. destacado
            elif cfop in ['5411', '5405']:
                status = 'Nao_paga'  # cfop não paga
            elif cfop in ['6202', '6411', '5202']:
                status = 'Devolucao'  # cfop devolucao nunca tem st
            else:
                status = 'Nao_Encontrado'  # cfop não encontrado na base
            return status

        def calcula_ST(valor_total_prod, ipi_prod, mva_orig, aliq_int_adq, aliq_int_vend):  # desenvolver
            total = float(valor_total_prod) + float(ipi_prod)
            xmva = float(mva_orig) * float(total)
            base_calc = xmva + total
            xaliq_intern_prod = float(base_calc) * float(aliq_int_adq)
            icms_op_prop = float(valor_total_prod) * float(aliq_int_vend)
            valor_st = xaliq_intern_prod - icms_op_prop
            return valor_st

        def read_table(file, table):
            filex = pd.read_excel(file, sheet_name=table, converters={'classif_fiscal_ncm': int})
            return filex

        def search_get_ncm_info(ncm, table):
            info = {}
            for i, x in enumerate(table['classif_fiscal_ncm']):
                if int(ncm) == x:
                    for col in table.columns:
                        info[col] = table.loc[i, col]
                    break
            return info

        def select_mva(forn_categ, cst, icms_ind, uf_forn):
            aliq_int_adqx = 0.195  # PR - Db Truck
            if forn_categ and (cst == 'Nacional' or cst == 'Nacional+imp40'):
                mva = 'MVA INTERNA'
                aliq_int_vendx = 0.12
                if uf_forn == 'PR':
                    aliq_int_vendx = 0.195
            elif not forn_categ and cst == 'Nacional':
                mva = 'MVA 12%'
                aliq_int_vendx = 0.12
                if uf_forn == 'PR':
                    aliq_int_vendx = 0.195
            elif cst == 'Nacional+imp40' and float(icms_ind) == 12:
                mva = 'MVA 12%'
                aliq_int_vendx = 0.12  # DUVIDA
                if uf_forn == 'PR':
                    aliq_int_vendx = 0.195
            elif (cst == 'Nacional+imp40' and float(icms_ind) == 4) or (cst == 'Importada'):
                mva = 'MVA 4% (IMPORTADO)'
                aliq_int_vendx = 0.04
            return mva, aliq_int_vendx, aliq_int_adqx
        
        def add_info_df(codigo, ncm, origem, cst, cfop, valor_total, valor_ipi, status, mva_uso, valor_st): #, difal):
            self.df_products = self.df_products._append({'codigo': codigo, 'ncm': ncm, 'origem': origem, 'cst': cst, 'cfop': cfop, 'valor_total': valor_total, 'valor_ipi': valor_ipi, 'status': status, 'mva_uso': mva_uso, 'valor_st': valor_st }, ignore_index=True) #'difal': difal

        def unify_by_ncm():
            grouped_df = self.df_products.groupby('ncm').agg({'mva_uso':'first', 'valor_total': 'sum', 'valor_ipi': 'sum', 'valor_st': 'sum'}).reset_index() #,'difal':'sum'
            return grouped_df[['ncm', 'mva_uso','valor_total', 'valor_ipi', 'valor_st',]] #'difal']]    
        
        def ncm_warning(ncm):
            ncm = str(ncm)
            text = f'NCM {ncm} não encontrada nos registros. \n Atualize a tabela e realiza a análise novamente'
            QMessageBox.warning(self, 'ATENÇÃO', text)


        # Main script
        data = xml_path
        base = read_table(r'TabelaNCM.xlsx', 'PR por NCM')


        # nf geral
        ftags = ['CNPJ', 'xNome', 'UF']  # vBC - base calculo ICMS
        forn = extract_data(data, dad_tag='infNFe', tags=ftags)
        forn[0].pop('item')

        nftags = ['nNF']
        nfn = extract_data(data, dad_tag='ide', tags=nftags)
        nfn[0].pop('item')

        bcitag = ['vBC']
        base_calc_icms = extract_data(data, dad_tag='ICMSTot', tags=bcitag)
        base_calc_icms[0].pop('item')

        # produtos
        ptags = ['cProd', 'NCM', 'orig', 'CST', 'CFOP', 'pICMS', 'vProd', 'vIPI']
        prods = extract_data(data, dad_tag='det', tags=ptags)

        # fornecedor do simples
        if float(base_calc_icms[0]['vBC']) != 0:
            forn_simples = False
        else:
            forn_simples = True

        

        self.print_to_output('_____________________________________________________________________________________________________________________________\n')    

        self.print_to_output('______________________________________________ANALISE DE IMPOSTO DE NF____________________________________________________\n')
                
        self.print_to_output('_____________________________________________________________________________________________________________________________\n')  
           

        self.print_to_output('NF: {}    |'.format(nfn[0]['nNF']))
        self.number_nf = str(nfn[0]['nNF'])
        self.print_to_output('   Fornecedor: {}\n'.format(forn[0]['xNome']))
        self.forn_name = str(forn[0]['xNome'])
        self.print_to_output('Estado: {}    |   '.format(forn[0]['UF']))
        self.print_to_output('CNPJ: {}\n'.format(forn[0]['CNPJ']))
        if forn_simples:
            self.print_to_output('Regime Trib: **  Simples Nacional  **\n')
            stats = 'Simples Nacional'
        else:
            self.print_to_output('Regime Trib: **  Normal  **\n')
            stats = 'Normal'


        st_total = 0
        #difal_total = 0
        ipi_total = 0

        for i in prods:
            self.print_to_output('_____________________________________________________________________________________________________________________________\n')
            self.print_to_output('Código      | {}\n'.format(i['cProd']))
            self.print_to_output('Valor total | R${}\n'.format(i['vProd']))
            self.print_to_output('CST            | {} - {}{} \n'.format(i['desc_origem'], i['orig'], i['CST']))
            self.print_to_output('IPI              | {}\n'.format(i['vIPI']))
            self.print_to_output('NCM         | {}\n'.format(i['NCM']))
            self.print_to_output('CFOP        | {} - {}\n'.format(i['CFOP'], i['status_imp']))
                  
            
      
            if i['status_imp'] == 'Analise':  
                analise = search_get_ncm_info(i['NCM'], base)
                ncm = str(i['NCM'])
                if analise == {}:
                    self.print_to_output('NCM não encontrada nos registros \n'
                                         'Realizar busca de NCM  e atualizar tabela\n')
                    ncm_warning(ncm)
                    return

                else:
                    interna = analise['MVA INTERNA']*100
                    imp4 = analise['MVA 4% (IMPORTADO)']*100
                    i7 = analise['MVA 7%']*100
                    i12 = analise['MVA 12%']*100
                    #self.print_to_output('Resultado|----MVA INT ({:.2f}%)---MVA 4 ({:.2f}%)---MVA 7 ({:.2f}%)---MVA 12  ({:.2f}%)---|\n'.format(interna, imp4, i7, i12))
                    if analise['MVA INTERNA'] in ['NÃO TEM ST', '']:
                        mvax = 0
                        #difal = 0
                        st=0
                        self.print_to_output('PRODUTO NÃO POSSUI ST\n') # OU DIFAL\n')
                        '''
                        if float(i['pICMS']) == 4:
                            self.print_to_output('ATENÇÃO!\n')
                            # CALCULA DIFAL
                            difal = float(i['vProd']) * 0.04
                            self.print_to_output('PRODUTO COM DIFAL, Valor calculado do DIFAL = R$ {:.3f}\n'.format(difal))
                            difal_total += difal
                        '''    
                    
                        
                    elif analise['MVA INTERNA'] == 'NCM INEXISTENTE':
                        mvax = 0
                        st = 0
                        #difal = 0
                        self.print_to_output('NCM inexistente, VERIFICAR\n')
                        #fazer janela de aviso

                    
                    else:
                        #difal=0
                        mva_tag, aliq_int_vendx, aliq_int_adqx = select_mva(forn_categ=forn_simples,
                                                                            cst=i['desc_origem'],
                                                                            icms_ind=i['pICMS'],
                                                                            uf_forn=forn[0]['UF'])
                        mvax = analise[mva_tag]
                        self.print_to_output('MVA Utilizado: {} %\n'.format(mvax * 100))
                        st = calcula_ST(valor_total_prod=i['vProd'], ipi_prod=i['vIPI'], mva_orig=mvax,
                                         aliq_int_adq=aliq_int_adqx, aliq_int_vend=aliq_int_vendx)
                        self.print_to_output('Valor calculado do ST = R$ {:.3f}\n'.format(st))
                        st_total += st
                       
            else:
                mvax = 0
                #difal = 0
                st = 0
                pass

            ipi_total += float(i['vIPI'])   
    
            #update df
            add_info_df( codigo=i['cProd'],ncm=i['NCM'],origem=i['desc_origem'],cst=(str(i['orig'])+str(i['CST'])),cfop=i['CFOP'],valor_total=float(i['vProd']),valor_ipi=float(i['vIPI']),status=i['status_imp'],mva_uso='{:.2f}%'.format(mvax*100), valor_st=st) #,difal=difal)

        row_define_forn = ['NF', 'Fornecedor','Estado','CNPJ','Regime Tributario','Valor total ST calculado', 'Valor total IPI'] #'Valor total Difal calculado',]
        data_forn = [nfn[0]['nNF'],forn[0]['xNome'],forn[0]['UF'], forn[0]['CNPJ'], stats,st_total,ipi_total,]#difal_total]
        data_forn_xl = {'Campos': row_define_forn, 'Dados': data_forn}
        self.df_forn = pd.DataFrame(data_forn_xl)    

        #plotting table 1
        self.tableWidget.setRowCount(self.df_products.shape[0])
        self.tableWidget.setColumnCount(self.df_products.shape[1])
        header_labels = [column.upper() for column in self.df_products.columns]
        self.tableWidget.setHorizontalHeaderLabels(header_labels)
        for i, row in self.df_products.iterrows():
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                if self.df_products.columns[j] in ['valor_total', 'valor_ipi', 'valor_st',]: # 'difal']:  
                    item.setData(0, value)  # Store the original value
                    item.setText(f"R$ {value:.2f}")  # Format the value to display two decimal places
                self.tableWidget.setItem(i, j, item)
        self.tableWidget.horizontalHeader().setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)

        #plotting table 2
        self.df_ncms = unify_by_ncm()
        self.tableWidget2.setRowCount(self.df_ncms.shape[0])
        self.tableWidget2.setColumnCount(self.df_ncms.shape[1])
        header_labels2 = ('NCM', 'MVA', 'TOTAL', 'TOTAL_IPI', 'TOTAL ST',) #'TOTAL_DIFAL')
        self.tableWidget2.setHorizontalHeaderLabels(header_labels2)
        for i, row in self.df_ncms.iterrows():
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                if self.df_ncms.columns[j] in ['valor_total', 'valor_ipi', 'valor_st',]: # 'difal']:  
                    item.setData(0, value)  # valor original
                    item.setText(f"R$ {value:.2f}")  # 2 decimais
                self.tableWidget2.setItem(i, j, item)
        self.tableWidget2.horizontalHeader().setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)        

        #plotting table 3
        self.tableWidget3.setRowCount(self.df_forn.shape[0])
        self.tableWidget3.setColumnCount(self.df_forn.shape[1])
        header_labels = [column.upper() for column in self.df_forn.columns]
        self.tableWidget3.setHorizontalHeaderLabels(header_labels)
        for i, row in self.df_forn.iterrows():
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                if type(value) == float:
                    item.setData(0, value)  # valor original
                    item.setText(f"R$ {value:.2f}")  # 2 decimais
                self.tableWidget3.setItem(i, j, item)     
        self.tableWidget3.horizontalHeader().setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        self.tableWidget3.verticalHeader().setVisible(False)
        self.tableWidget3.resizeColumnsToContents()    

        
    


    #result print    
            
        self.print_to_output('_____________________________________________________________________________________________________________________________\n')
        self.print_to_output('_____________________________________________________________________________________________________________________________\n')
        if st_total > 0:
            self.print_to_output('Valor total ST - R$ {}\n\n'.format(st_total))
        else:
            self.print_to_output('NF sem valores de ST a pagar.\n\n')
        '''
        if difal_total > 0:
            self.print_to_output('Valor total Difal - R$ {}\n'.format(difal_total))
        else:
            self.print_to_output('NF sem valores de DIFAL a pagar.\n')
        '''
        self.print_to_output('_____________________________________________________________________________________________________________________________\n')    


    def edit_ncm_table(self):
        try:
            subprocess.Popen([r'TabelaNCM.xlsx'], shell=True)
        except Exception as e:
            QMessageBox.critical(self, 'Erro', f'Erro ao abrir o arquivo: {str(e)}')        

        
    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon('ICO_FISCO.ico'))
    window = XMLAnalyzer()
    window.show()
    sys.exit(app.exec())
