#-*- coding: utf-8 -*-
# ---------------------------------------------------------------------------
# Geração de Polígonos de Thiessen Variáveis  - Plugin QGIS3
# Cláudio Bielenki Júnior
# Especialista em Geoprocessamento -  SUM  -  ANA
# Março 2022     03032022
# ---------------------------------------------------------------------------
import numpy as np
import rasterio.shutil
import fiona

import pandas as pd
import geopandas as gpd
from pyproj import CRS
from pyproj.database import query_utm_crs_info
from pyproj.aoi import AreaOfInterest
import win32com, sys, string, os
from win32com.client import Dispatch
import shutil
from osgeo import ogr, gdal
from qgis.PyQt.QtCore import QSettings, QTranslator, QCoreApplication, QVariant
from qgis.PyQt.QtGui import QIcon
from qgis.PyQt.QtWidgets import QAction, QFileDialog, QLabel
import qgis.core
from qgis.core import QgsMapLayerProxyModel
from qgis.core import QgsMessageLog
from qgis.core import QgsVectorLayer, QgsProject

from qgis.gui import QgsMapLayerComboBox, QgsFieldComboBox, QgsFeaturePickerWidget
from qgis.core import QgsFieldProxyModel, QgsProcessingContext, QgsProcessingFeedback, QgsProcessingParameters, QgsVectorFileWriter, QgsExpression, QgsSpatialIndex, QgsProcessing, QgsProcessingMultiStepFeedback
from qgis.core import QgsProcessingUtils, Qgis, QgsFeatureRequest, QgsProcessingAlgorithm, QgsProcessingFeatureSource, QgsWkbTypes, QgsRectangle, QgsField, QgsFeature

from qgis.analysis import *
import processing

from .pMediaT_dialog import pMediaTDialog
from .pMediaIDW_dialog import pMediaIDWDialog
from .pFill_dialog import pFillDialog

from .algoritmos import thiessenClip
# Initialize Qt resources from file resources.py
from .resources import *



# Importando os módulos
import xlrd


def bbox_to_pixel_offsets(gt, bbox):
     originX = gt[0]
     originY = gt[3]
     pixel_width = gt[1]
     pixel_height = gt[5]
     x1 = int((bbox[0] - originX) / pixel_width)
     x2 = int((bbox[1] - originX) / pixel_width) + 1
     y1 = int((bbox[3] - originY) / pixel_height)
     y2 = int((bbox[2] - originY) / pixel_height) + 1
     xsize = x2 - x1
     ysize = y2 - y1
     return (x1, y1, xsize, ysize)
def zonal_stats(vector_path, raster_path, banda, nodata_value=None, global_src_extent=False):
     rds = gdal.Open(raster_path)
     assert(rds)
     rb = rds.GetRasterBand(banda)
     rgt = rds.GetGeoTransform()
     if nodata_value:
          nodata_value = float(nodata_value)
          rb.SetNoDataValue(nodata_value)
     vds = ogr.Open(vector_path)
     assert(vds)
     vlyr = vds.GetLayer(0)
     if global_src_extent:
          src_offset = bbox_to_pixel_offsets(rgt, vlyr.GetExtent())
          src_array = rb.ReadAsArray(*src_offset)
          new_gt = ((rgt[0] + (src_offset[0] * rgt[1])), rgt[1], 0.0, (rgt[3] + (src_offset[1] * rgt[5])), 0.0, rgt[5] )
     mem_drv = ogr.GetDriverByName('Memory')
     driver = gdal.GetDriverByName('MEM')
     stats = []
     ArrayMasked=[]
     feat = vlyr.GetNextFeature()
     cont=0
     while feat is not None:
          if not global_src_extent:
               src_offset = bbox_to_pixel_offsets(rgt, feat.geometry().GetEnvelope())
               src_array = rb.ReadAsArray(*src_offset)
               new_gt = ((rgt[0] + (src_offset[0] * rgt[1])), rgt[1], 0.0, (rgt[3] + (src_offset[1] * rgt[5])), 0.0, rgt[5])
               mem_ds = mem_drv.CreateDataSource('out')
          mem_layer = mem_ds.CreateLayer('poly', None, ogr.wkbPolygon)
          mem_layer.CreateFeature(feat.Clone())
          rvds = driver.Create('', src_offset[2], src_offset[3], 1, gdal.GDT_Byte)
          rvds.SetGeoTransform(new_gt)
          gdal.RasterizeLayer(rvds, [1], mem_layer, burn_values=[1])
          rv_array = rvds.ReadAsArray()
          masked = np.ma.MaskedArray( src_array, mask=np.logical_or(src_array == nodata_value, np.logical_not(rv_array) ) )
          ArrayMasked.append(masked)
          feature_stats = {'min': float(masked.min()),'mean': float(masked.mean()),'max': float(masked.max()),'std': float(masked.std()),'sum': float(masked.sum()),'count': int(masked.count()),'fid': int(feat.GetFID())}
          stats.append(feature_stats)
          rvds = None
          mem_ds = None
          feat = vlyr.GetNextFeature()
          cont=cont+1
     ds = None
     ds = None
     return stats, ArrayMasked, cont




class pMedia:
    def __init__(self, iface):

        # Save reference to the QGIS interface
        self.iface = iface
        # initialize plugin directory
        self.plugin_dir = os.path.dirname(__file__)
        # initialize locale
        locale = QSettings().value('locale/userLocale')[0:2]
        locale_path = os.path.join(
            self.plugin_dir,
            'i18n',
            'pMedia_{}.qm'.format(locale))

        if os.path.exists(locale_path):
            self.translator = QTranslator()
            self.translator.load(locale_path)
            QCoreApplication.installTranslator(self.translator)

        # Declare instance attributes
        self.actions = []
        #self.menu = self.tr(u'&CompPlex')

        # Check if plugin was started the first time in current QGIS session
        # Must be set in initGui() to survive plugin reloads
        self.first_start = None

    # noinspection PyMethodMayBeStatic
    def tr(self, message):

        # noinspection PyTypeChecker,PyArgumentList,PyCallByClass
        return QCoreApplication.translate('pMedia', message)


    def add_action(
        self,
        icon_path,
        text,
        callback,
        enabled_flag=True,
        add_to_menu=False,
        add_to_toolbar=True,
        status_tip=None,
        whats_this=None,
        parent=None):


        icon = QIcon(icon_path)
        action = QAction(icon, text, parent)
        action.triggered.connect(callback)
        action.setEnabled(enabled_flag)

        if status_tip is not None:
            action.setStatusTip(status_tip)

        if whats_this is not None:
            action.setWhatsThis(whats_this)

        if add_to_toolbar:
            # Adds plugin icon to Plugins toolbar
            self.toolBar.addAction(action)

        if add_to_menu:
            self.iface.addPluginToMenu(
                self.menu,
                action)

        self.actions.append(action)

        return action

    def hello(self):
        self.iface.messageBar().pushMessage(u'Welcome to pMedia Tools', level=Qgis.Info, duration=3)

    def initGui(self):
        """Create the menu entries and toolbar icons inside the QGIS GUI."""

        self.toolBar = self.iface.addToolBar("pMedia Tools")
        self.toolBar.setObjectName("pMediaTools")

        icon_path4 = ':/plugins/pyPmedia/fig1.png'
        self.add_action(
            icon_path4,
            text=self.tr(u'pMedia Tools'),
            callback=self.hello,
            parent=self.iface.mainWindow())


        pMediaLabel=QLabel(self.toolBar)
        pMediaLabel.setText("pMedia Tools: ")
        self.toolBar.addWidget(pMediaLabel)


        icon_path = ':/plugins/pyPmedia/thiessen.png'
        self.add_action(
            icon_path,
            text=self.tr(u'pMedia Thiessen'),
            callback=self.runThiessen,
            parent=self.iface.mainWindow())

        icon_path2 = ':/plugins/pyPmedia/idw.png'
        self.add_action(
            icon_path2,
            text=self.tr(u'pMedia IDW'),
            callback=self.runIDW,
            parent=self.iface.mainWindow())
        # will be set False in run()
        pFillLabel=QLabel(self.toolBar)
        pFillLabel.setText("Fill Gap Tools: ")
        self.toolBar.addWidget(pFillLabel)


        icon_path5 = ':/plugins/pyPmedia/fill.PNG'
        self.add_action(
            icon_path5,
            text=self.tr(u'Fill Gap Tools'),
            callback=self.runFill,
            parent=self.iface.mainWindow())


    def unload(self):
        """Removes the plugin menu item and icon from QGIS GUI."""
        for action in self.actions:
            self.iface.removePluginMenu(
                self.tr(u'&pMedia'),
                action)
            self.iface.removeToolBarIcon(action)

    def selecionar_xls(self, projectPath):
        arquivoCaminho = QFileDialog.getOpenFileName(self.dlg, "Select Rainfall Data File Input: ", self.projectPath, "*.xls")
        self.dlg.pathRainfallData.setText(arquivoCaminho[0])



    def rainfallChange(self):
        self.dlg.cbCodeField.setLayer(self.dlg.cbLayerRainfall.currentLayer())

    def watershedChange(self):
        self.dlg.cbAreaField.setLayer(self.dlg.cbLayerWatershed.currentLayer())

    def watershedChange2(self):
        self.dlg.cbIDField.setLayer(self.dlg.cbLayerWatershed.currentLayer())

    def runThiessen(self):
        """Run method that performs all the real work"""
        self.dlg = pMediaTDialog()
        self.projectPath = QgsProject.instance().homePath()
        # Create the dialog with elements (after translation) and keep reference
        # Only create GUI ONCE in callback, so that it will only load when the plugin is started
        if self.first_start == True:
            self.first_start = False

        self.dlg.cbLayerRainfall.setFilters(QgsMapLayerProxyModel.PointLayer)
        self.dlg.cbLayerWatershed.setFilters(QgsMapLayerProxyModel.PolygonLayer)

        self.dlg.cbCodeField.setLayer(self.dlg.cbLayerRainfall.currentLayer())
        #self.dlg.cbCodeField.setFilters(QgsFieldProxyModel.Numeric)

        self.dlg.cbAreaField.setLayer(self.dlg.cbLayerWatershed.currentLayer())
        self.dlg.cbAreaField.setFilters(QgsFieldProxyModel.Numeric)

        self.dlg.cbLayerRainfall.layerChanged.connect(self.rainfallChange)
        self.dlg.cbLayerWatershed.layerChanged.connect(self.watershedChange)

        self.dlg.pathRainfallData.clear()
        #self.dlg.pathRainfallOutput.clear()
        self.dlg.pbRainfallData.clicked.connect(self.selecionar_xls)
        #self.dlg.pbRainfalOutput.clicked.connect(self.selecionar_output)
        # show the dialog
        self.dlg.show()
        # Run the dialog event loop
        result = self.dlg.exec_()
        # See if OK was pressed
        if result:
            # Recebendo as variaveis de entrada
            idGage = self.dlg.cbCodeField.currentField()
            areaField = self.dlg.cbAreaField.currentField()
            #print("field = " + areaField)
            fc_estacao = self.dlg.cbLayerRainfall.currentLayer().dataProvider().dataSourceUri().split('|')[0] # definindo a feature class das estacoes
            fc_bacia = self.dlg.cbLayerWatershed.currentLayer().dataProvider().dataSourceUri().split('|')[0] # definindo a feature class da área da bacia
            file = self.dlg.pathRainfallData.text()# arquivo de entrada de dados de pluviometria

            shutil.copy(file, file.replace('.xls','_output.xls'))

            saida_xls = file.replace('.xls','_output.xls') #self.dlg.pathRainfallOutput.text() # arquivo de saida de dados de pluviometria com ponderadores
            layerGages = QgsVectorLayer(fc_estacao, "Gages", "ogr")
            layerWatershed = QgsVectorLayer(fc_bacia, "Watersheds", "ogr")

            extGages = layerGages.extent()
            extWatersheds = layerWatershed.extent()

            #define o extent para os poligonos de thiessen
            ext_x=[]
            ext_y=[]

            ext_x.append(extGages.xMinimum())
            ext_x.append(extGages.xMaximum())
            ext_x.append(extWatersheds.xMaximum())
            ext_x.append(extWatersheds.xMinimum())
            ext_y.append(extGages.yMinimum())
            ext_y.append(extGages.yMaximum())
            ext_y.append(extWatersheds.yMinimum())
            ext_y.append(extWatersheds.yMaximum())
            x_maximo=max(ext_x)+10000
            x_minimo=min(ext_x)-10000
            y_maximo=max(ext_y)+10000
            y_minimo=min(ext_y)-10000
            #.Extent(x_minimo,y_minimo,x_maximo,y_maximo)


            dataP = layerGages.dataProvider()
            fieldsGages = dataP.fields()
            if 'area' not in fieldsGages.names():
                dataP.addAttributes([QgsField("area",QVariant.Double)])
                layerGages.updateFields()

            if 'pond' not in fieldsGages.names():
                dataP.addAttributes([QgsField("pond",QVariant.Double)])
                layerGages.updateFields()

            watershedFeat = QgsFeature()
            layerWatershed.getFeatures().nextFeature(watershedFeat)
            vArea = watershedFeat[str(areaField)]


            xls = xlrd.open_workbook(file) #abre o arquivo xls de dados de entrada
            dados=xls.sheets()[0] #variável dados aponta para a 1ª planilha do xls
            ncol=dados.ncols #número de colunas
            nrow=dados.nrows #número de linhas
            linhas=dados.row_values #vari?vel dados recebe os valores da planilha
            estacao=linhas(0) #vetor estacao recebe a 1ª linha da planilha
            linha=[0]*(nrow)#cria o vetor linha

            # looping para ler cada linha de dados
            for i in range(1,nrow):
            	linha[i]=linhas(i)

            combinacao=[0]*(nrow)#variável combinacao para receber o valor da combinação soma(2^posição) para cada data
            combinacao_sel=[0]#variável que armazena os valores de combinação sem repetição
            comb_est=[0]#variável que armazena a lista de estações para cada combinação selecao
            comb_pond=[0]#variável que armazena a lista de ponderadores para cada combinação
            comb_est_clip=[0]
            for j in range(1,nrow): #looping pelas linhas
            	combinacao[j]=0 #inicializa a variável combinacao
            	est_aux=[] #inicializa a variável est_aux para capturar as estações com dados em cada data
            	for k in range(1,ncol): #looping pelas colunas
            		if linha[j][k]!= '': #verifica se existe dado medido para a estação numa data
            			combinacao[j]=combinacao[j] + 2**(k-1) #calcula o valor da combinação
            			est_aux.append(estacao[k]) #seleciona os c?dios das estações com dados
            	if not(combinacao[j] in combinacao_sel) : #verifica se a combinação já existe
            		combinacao_sel.append(combinacao[j]) #insere na lista de combinações uma nova combinação
            		comb_est.append(est_aux)#armazena a lista de estações para uma nova combina??o

            consulta=[0]*len(combinacao_sel)#inicializa a variável consulta

            for x in range(1,len(combinacao_sel)): #looping para cada combinação
            	aux='' #inicializa a variável aux
            	expressao='' #inicializa a variável expressao
            	for y in range(0,len(comb_est[x])): #looping pelas estações selecionadas para uma dada combinação
            		sel= '\"' + idGage + '\" =' + ' ' + str(int(comb_est[x][y])) + ' ' + 'OR ' #montagem da expressão SQL
            		aux=aux+sel
            	for i in range(len(aux)-4): #looping para eliminar o último OR da expressão
            		expressao=expressao+aux[i]

            	consulta[x] = expressao #armazena as respectivas consultas SQL para cada combinação
            #print(consulta)


            #funcoes de geoprocessamento para selecionar, criar os poligonos e clipar
            for x in range(1,len(consulta)):

                saida_thiessen = self.projectPath + r"/Thiessen_combinacao_" + str(combinacao_sel[x])+".shp"
                #QgsVectorLayer(self.projectPath + r"/Thiessen_combinacao_" + str(combinacao_sel[x])+".shp", '','ogr')#nome para os arquivos de poligonos de saida
                saida_thiessen_Clip = self.projectPath + r"/Thiessen_combinacao_" + str(combinacao_sel[x])+"_Clip.shp"
                #QgsVectorLayer(self.projectPath + r"/Thiessen_combinacao_" + str(combinacao_sel[x])+"_Clip.shp", '','ogr') #nome para os arquivos de poligonos clipados de saida

                expr = QgsExpression(consulta[x])
                layerGagesSelects = layerGages.materialize(QgsFeatureRequest(expr))
                gages = layerGagesSelects.getFeatures()


                feedback = QgsProcessingFeedback()
                context = QgsProcessingContext()
                context.setProject(QgsProject.instance())
                parameters={}
                parameters['layerGagesSelects'] = layerGagesSelects
                parameters['layerWatershed'] = layerWatershed
                parameters['box'] = [x_minimo,y_minimo,x_maximo,y_maximo]
                # parameters['layerThiessenTemp'] = saida_thiessen
                parameters['layerClipTemp'] = saida_thiessen_Clip
                results = thiessenClip.fThiessenClip(self, parameters, context, feedback)

                #layerThiessenClipStr = QgsProcessingUtils.mapLayerFromString(results['CLIP'], context)

                #QgsVectorFileWriter.writeAsVectorFormat( layerThiessenClipStr , saida_thiessen_Clip, "UTF-8", self.iface.activeLayer().crs() , "ESRI Shapefile")

                layerThiessenClip = QgsVectorLayer(results['CLIP'],'','ogr')

                pond_aux=[] #inicializa a variavel auxiliar para armazenar os ponderadores
                est_clip_aux=[]
                features = layerThiessenClip.getFeatures()
                layerThiessenClip.startEditing()

                for feature in features:
                #while pRow_Thiessen: # Loop pelas linhas da feature class poligonos thiessen clipados
                    area = feature.geometry().area()
                    feature['area'] = area
                    feature['pond'] = (area/vArea) #Calcula o ponderador
                    est_clip_aux.append(feature[idGage])
                    pond_aux.append((area/vArea)) # Inclui o ponderador calculado no fim do vetor auxiliar
                    layerThiessenClip.updateFeature(feature)

                layerThiessenClip.commitChanges()
                comb_pond.append(pond_aux) #armazena em um vetor os ponderadores
                comb_est_clip.append(est_clip_aux)

                layerThiessenClip = None
                layerGagesSelects = None
                #QgsVectorFileWriter.deleteShapeFile(saida_thiessen)
                QgsVectorFileWriter.deleteShapeFile(saida_thiessen_Clip)

            # Prepara o arquivo de dados para gravação dos ponderadores
            xlApp = Dispatch("Excel.Application")
            docxls = xlApp.Workbooks.Open(saida_xls)
            docxls.Sheets(1).Select()
            planilha = docxls.ActiveSheet
            # Gravacao dos ponderadores no xls de saida
            for data in range(1,len(combinacao)): # Looping pelas datas
                comb=combinacao[data] #Variável auxiliar para armazenar o valor da combinacao
                index_comb=combinacao_sel.index(comb) # Index para recuperar no vetor de estacoes e ponderadores
                est=comb_est_clip[index_comb] # Recupera as estacoes
                pond=comb_pond[index_comb] # Recupera o ponderador
                precx = 0
                for col in range(0,len(pond)): # Looping pelos ponderadores calculados para a data determinada
                    codigo_est=est[col] # Recupera a estação
                    ponderador=pond[col] # Recupera o ponderador
                    index_est=estacao.index(codigo_est) # Recupera a posicao da estação
                    planilha.Cells(data+1,(index_est+ncol+1)).Value = ponderador # Grava o ponderador na planilha de acordo com a posicao recuperada da estação
                    precx = precx + (ponderador * planilha.Cells(data+1,(index_est+1)).Value)
                planilha.Cells(data+1,((2*ncol) +2)).Value = precx
            # Salva o xls
            xlApp.Visible=1
            docxls.save
            # Termina a aplica??o xls
            xlApp.Quit()
            # Limpa a mem?ria
            del xlApp
            pass

    def runIDW(self):

        self.dlg = pMediaIDWDialog()
        self.projectPath = QgsProject.instance().homePath()
        # Create the dialog with elements (after translation) and keep reference
        # Only create GUI ONCE in callback, so that it will only load when the plugin is started
        if self.first_start == True:
            self.first_start = False

        self.dlg.cbLayerRainfall.setFilters(QgsMapLayerProxyModel.PointLayer)
        self.dlg.cbLayerWatershed.setFilters(QgsMapLayerProxyModel.PolygonLayer)

        self.dlg.cbCodeField.setLayer(self.dlg.cbLayerRainfall.currentLayer())
        #self.dlg.cbCodeField.setFilters(QgsFieldProxyModel.Numeric)

        self.dlg.cbIDField.setLayer(self.dlg.cbLayerWatershed.currentLayer())
        self.dlg.cbIDField.setFilters(QgsFieldProxyModel.Numeric)

        self.dlg.cbLayerRainfall.layerChanged.connect(self.rainfallChange)
        self.dlg.cbLayerWatershed.layerChanged.connect(self.watershedChange2)

        self.dlg.pathRainfallData.clear()
        #self.dlg.pathRainfallOutput.clear()
        self.dlg.pbRainfallData.clicked.connect(self.selecionar_xls)
        #self.dlg.pbRainfalOutput.clicked.connect(self.selecionar_output)
        # show the dialog
        self.dlg.show()
        # Run the dialog event loop
        result = self.dlg.exec_()
        if result:
            # Recebendo as variaveis de entrada
            idGage = self.dlg.cbCodeField.currentField()
            idWatershed = self.dlg.cbIDField.currentField()
            #print("field = " + areaField)
            fc_estacao = self.dlg.cbLayerRainfall.currentLayer().dataProvider().dataSourceUri().split('|')[0] # definindo a feature class das estacoes
            fc_bacia = self.dlg.cbLayerWatershed.currentLayer().dataProvider().dataSourceUri().split('|')[0] # definindo a feature class da área da bacia
            file = self.dlg.pathRainfallData.text()# arquivo de entrada de dados de pluviometria

            shutil.copy(file, file.replace('.xls','_output.xls'))

            saida_xls = file.replace('.xls','_output.xls') #self.dlg.pathRainfallOutput.text() # arquivo de saida de dados de pluviometria com ponderadores
            layerGages = QgsVectorLayer(fc_estacao, "Gages", "ogr")
            layerWatershed = QgsVectorLayer(fc_bacia, "Watersheds", "ogr")

            extGages = layerGages.extent()
            extWatersheds = layerWatershed.extent()

            #define o extent para os poligonos de thiessen
            ext_x=[]
            ext_y=[]

            ext_x.append(extGages.xMinimum())
            ext_x.append(extGages.xMaximum())
            ext_x.append(extWatersheds.xMaximum())
            ext_x.append(extWatersheds.xMinimum())
            ext_y.append(extGages.yMinimum())
            ext_y.append(extGages.yMaximum())
            ext_y.append(extWatersheds.yMinimum())
            ext_y.append(extWatersheds.yMaximum())
            x_maximo=max(ext_x)+10000
            x_minimo=min(ext_x)-10000
            y_maximo=max(ext_y)+10000
            y_minimo=min(ext_y)-10000
            Extent = QgsRectangle(x_minimo,y_minimo,x_maximo,y_maximo)


            dataP = layerGages.dataProvider()
            fieldsGages = dataP.fields()
            if 'area' not in fieldsGages.names():
                dataP.addAttributes([QgsField("area",QVariant.Double)])
                layerGages.updateFields()

            if 'pond' not in fieldsGages.names():
                dataP.addAttributes([QgsField("pond",QVariant.Double)])
                layerGages.updateFields()

            if 'prec' not in fieldsGages.names():
                dataP.addAttributes([QgsField("prec",QVariant.Double)])
                layerGages.updateFields()

            # Lendo os dados de entrada do xls
            xls = xlrd.open_workbook(file) # Abre o arquivo xls de dados de entrada
            dados=xls.sheets()[0] # Variável dados aponta para a 1ª planilha do xls
            ncol=dados.ncols # Número de colunas
            nrow=dados.nrows # Número de linhas
            linha=dados.row_values # Variável dados recebe os valores da planilha
            estacao=linha(0) # Vetor estacao recebe a 1ª linha da planilha

            # Prepara o arquivo de dados para gravação da PMédia
            nome_plans=[] # Vetor que receberá os identificadores das bacias para nomear as planilhas
            pbacias = layerWatershed.getFeatures() # Posciciona o cursor na feição de bacias
            for pbacia in pbacias: # Inicia o looping pelas rows
                nome_plans.append(pbacia[idWatershed]) # Captura os valores dos identificadores

            xlApp = Dispatch("Excel.Application") # Inicia o Excel
            wb = xlApp.Workbooks.Open(saida_xls) # Abre o arquivo de saída
            for i in range(wb.Sheets.Count, 1, -1): # Looping pelas planilhas
                wb.Sheets(i).Delete() # Deleta as planilhas exceto a 1ª
            for i in range(0, len(nome_plans)): # Looping pelos identificadores das bacias
                planilha = xlApp.ActiveWorkbook.Worksheets.Add() # Adiciona uma planilha
                planilha.Name=nome_plans[i] # Nomeia a planilha de acordo com os identificadores

            # Loop pelas linhas do arquivo de dados pluviometricos
            num=1 # Contador
            fileToDelete = []
            for j in range(1,nrow):
                sel =['']*ncol # Inicializa a variável sel
                expressao=''# Inicializa a variável espressao
                aux=''# Inicializa a variável auxiliar
                est_vet=[]# Inicializa a variável est_ve
                prec_vet=[] # Inicializa a variável prec_ve
                # Loop pelos postos pluviométricos
                for i in range(1,ncol):
                    if linha(j)[i]!= '': # Verifica se existe dado medido, caso sim copia o código para montar a expressção SQL
                        est_valor=(linha(0)[i]) # Lê o código da estação na primeira linha da planilha
                        est_vet.append(est_valor) # Armazena o código no vetor est_vet
                        prec_valor=(linha(j)[i])# Lê o valor de precipitação
                        prec_vet.append(prec_valor) # Armazena o valor de precipitação no vetor prec_vet
                        sel[i] = '\"' + idGage + '\" =' + ' ' + str(int(est_valor)) + ' ' + 'OR ' # Monta a expressão SQL para a referida estação

                    aux=aux+sel[i] # Monta a expressão SQL
                # Apaga o último OR da expressão SQL
                for t in range(len(aux)-4):
                    expressao=expressao+aux[t]
                print(expressao)
                # Nomes aos arquivos auxiliares
                selecaoName=r"\selecao_"+str(num)+'.shp'
                selecaoPath = self.projectPath + selecaoName
                rasterName=r"\IDW_"+str(num)+'.tif'
                rasterPath = self.projectPath + rasterName


                expr = QgsExpression(expressao)
                layerGagesSelects = layerGages.materialize(QgsFeatureRequest(expr))

                gages = layerGagesSelects.getFeatures()
                layerGagesSelects.startEditing()
                for gage in gages:
                    cod = gage[idGage] # Armazena o código da estação
                    prec_valor=prec_vet[(est_vet.index(cod))] # Copia o valor da precipitação do vetor prec_vet de acordo do index de posição do código da estação para o vetor prec_valor
                    gage['prec'] = prec_valor # Copia o valor da precipitação para a feature class de estações selecionadas
                    layerGagesSelects.updateFeature(gage)
                layerGagesSelects.commitChanges()

                #QgsVectorFileWriter.writeAsVectorFormat(layerGagesSelects, selecaoPath, "UTF-8", layerGages.crs() , "ESRI Shapefile")
                #gagesSelects = QgsVectorLayer(selecaoPath, "GagesSel", "ogr")
                gagesSelects = layerGagesSelects
                provider = gagesSelects.dataProvider()
                idxPrec = provider.fieldNameIndex('prec')

                idwData = QgsInterpolator.LayerData()
                idwData.source = gagesSelects
                idwData.zCoordInterpolation=False
                idwData.interpolationAttribute = idxPrec
                idwData.mInputType = 1

                IDW = QgsIDWInterpolator([idwData])

                res = 1000
                cols = int( ( x_maximo - x_minimo) / res )
                rows = int( (y_maximo - y_minimo ) / res)
                outputIDW = QgsGridFileWriter(IDW,rasterPath,Extent,cols,rows)
                outputIDW.writeFile()


                SHP = ogr.Open(fc_bacia)
                layerW = SHP.GetLayer()
                FeatureCount=layerW.GetFeatureCount()
                imagem = gdal.Open(rasterPath)
                stats = zonal_stats(fc_bacia, rasterPath, 1)
                x = stats[1]

                for iter in range(FeatureCount):
                    id=layerW.GetFeature(iter).GetField(idWatershed)
                    EstDescritivas=stats[0][iter]
                    media = EstDescritivas['mean']
                    plan=str(id) # Copia o valor do campo
                    planilha=wb.Worksheets(plan) # Ativa a planilha no Excel de acordo com o valor do campo
                    planilha.Cells(j,1).Value = media # Escreve o valor da média na planilha
                num=num+1 # Incrementa a variável num

                del idwData
                del IDW
                del provider
                del gagesSelects
                del layerGagesSelects
                del gage
                del gages

                fileToDelete.append(selecaoPath)

                del outputIDW
                imagem = None
                rasterio.shutil.delete(rasterPath)


            wb.save # Salva o Excel

            # Terminar aplicação
            xlApp.Quit()

            # Limpar a memória
            del xlApp

            pass
    def selecionar_csv(self, projectPath):
        arquivoCaminho = QFileDialog.getOpenFileName(self.dlg, "Select Rainfall Data File Input: ", self.projectPath, "*.csv")
        self.dlg.pathRainfallData.setText(arquivoCaminho[0])
        fileCSV = arquivoCaminho[0]
        dataFrame = pd.read_csv(fileCSV)
        colNames = list(dataFrame)
        self.dlg.comboBoxDate.addItems(colNames)

    def runFill(self):
        self.dlg = pFillDialog()
        self.projectPath = QgsProject.instance().homePath()
        # Create the dialog with elements (after translation) and keep reference
        # Only create GUI ONCE in callback, so that it will only load when the plugin is started
        if self.first_start == True:
            self.first_start = False

        self.dlg.cbLayerRainfall.setFilters(QgsMapLayerProxyModel.PointLayer)

        self.dlg.cbCodeField.setLayer(self.dlg.cbLayerRainfall.currentLayer())
        #self.dlg.cbCodeField.setFilters(QgsFieldProxyModel.Numeric)

        self.dlg.cbLayerRainfall.layerChanged.connect(self.rainfallChange)


        self.dlg.pathRainfallData.clear()
        #self.dlg.pathRainfallOutput.clear()
        self.dlg.pbRainfallData.clicked.connect(self.selecionar_csv)
        #self.dlg.pbRainfalOutput.clicked.connect(self.selecionar_output)
        # show the dialog

        #self.dlg.pathRainfallData.textChanged.connect(self.loadCombBox)


        self.dlg.show()
        # Run the dialog event loop
        result = self.dlg.exec_()

        if result:

            if self.dlg.rbRPM.isChecked()==True:
                method = "Mean"
            if self.dlg.rbRPC.isChecked()==True:
                method = "Correlation"
            if self.dlg.rbIDW.isChecked()==True:
                method = "InvDist"
            os.chdir(self.projectPath)

            fileCSV = self.dlg.pathRainfallData.text()
            fileSHP = self.dlg.cbLayerRainfall.currentLayer().dataProvider().dataSourceUri().split('|')[0]
            indexData = self.dlg.comboBoxDate.currentText()
            indexSHP = self.dlg.cbCodeField.currentField()
            dataPlu = pd.read_csv(fileCSV, index_col=indexData)
            gagePlu = gpd.read_file(fileSHP)


            gagePlu.set_index(indexSHP)
            indexList = dataPlu.index.values.tolist()
            srid = CRS(gagePlu.crs)

            if srid.coordinate_system.name == 'ellipsoidal':
                extent = gagePlu.total_bounds
                utm_crs_list = query_utm_crs_info( datum_name="WGS 84", area_of_interest=AreaOfInterest( west_lon_degree=extent[0], south_lat_degree=extent[1], east_lon_degree=extent[2], north_lat_degree=extent[3], ), )
                utm_crs = CRS.from_epsg(utm_crs_list[0].code)
                gagePlu = gagePlu.to_crs(utm_crs)

            means = dataPlu.mean()
            stds = dataPlu.std()
            colNames=list(dataPlu.columns)
            rows, columns = dataPlu.shape
            matrixCorr = dataPlu.corr()
            arrayCorr = matrixCorr.to_numpy()

            matrixDist = gagePlu.geometry.apply(lambda g: ((gagePlu.distance(g))))


            invMatrixDist = 1/matrixDist
            arrayDist = invMatrixDist.to_numpy()

            for col in range(columns):
                for row in range(rows):
                    if col==row:
                        invMatrixDist[col][row]=0


            matrixIndices = np.multiply(matrixCorr,invMatrixDist)
            arrayIndices = matrixIndices.to_numpy()
            arrayData = dataPlu.to_numpy()
            arrayDataP = np.copy(arrayData)

            arrayPosition=[]
            for col in range(columns):
                pMeans = means[col]
                pStds = stds[col]

                for row in range(rows):

                    if np.isnan(arrayData[row,col]):
                        arrayPIndicesCopy = np.empty(shape=columns)
                        rowData = np.empty(shape=columns)
                        arrayPMedia = np.empty(shape=columns)
                        arrayPStds = np.empty(shape=columns)
                        arrayPCorr = np.empty(shape=columns)
                        arrayPIndices = np.empty(shape=columns)
                        arrayPDist = np.empty(shape=columns)
                        arraySortIndices = np.empty(shape=columns)
                        arraySelecao = np.empty(shape=5)
                        precX = 0

                        for i in range(columns):
                            if not np.isnan(arrayData[row, i]):
                                rowData[i] = arrayData[row, i]
                                arrayPMedia[i] = means[i]
                                arrayPStds[i] = stds[i]
                                arrayPCorr[i] = arrayCorr[col, i]
                                arrayPIndices[i] = arrayIndices[col, i]
                                arrayPDist[i] = arrayDist[col, i]

                            else:
                                rowData[i] = 0
                                arrayPMedia[i] = 0
                                arrayPStds[i] = 0
                                arrayPCorr[i] = 0
                                arrayPIndices[i] = 0
                                arrayPDist[i] = 0

                        arrayPIndicesCopy = np.copy(arrayPIndices)
                        arraySortIndices = np.sort(arrayPIndicesCopy)
                        Cont = 0
                        for valor in arraySortIndices:
                            if valor > 0 and Cont <4:
                                Cont = Cont + 1
                                arraySelecao[Cont] = valor


                        arraySelecao = arraySortIndices[-5:]
                        #print(arraySelecao)
                        arrayPos=[]
                        for item in arraySelecao:
                            position = np.where(arrayPIndices == item)[0][0]

                            if method == "Mean":
                                precX = precX + ((1 / 5) * ((pMeans / arrayPMedia[position]) * rowData[position]))
                            if method == "Correlation":
                                precX = precX + ((pStds / 5) * (((rowData[position] - arrayPMedia[position]) / arrayPStds[position]) * arrayPCorr[position]))
                            if method == "InvDist":
                                somaDist=0
                                for item2 in arraySelecao:
                                    position2 = np.where(arrayPIndicesCopy == item2)[0][0]
                                    somaDist = somaDist + arrayPDist[position2]
                                precX = precX + ((arrayPDist[position]/somaDist) * rowData[position])

                        if method == "Mean":
                            arrayDataP[row,col] = precX
                        if method == "Correlation":
                            arrayDataP[row,col] = precX + pMedia
                        if method == "InvDist":
                            arrayDataP[row,col] = precX

                        del rowData
                        del arrayPMedia
                        del arrayPStds
                        del arrayPCorr
                        del arrayPIndices
                        del arrayPIndicesCopy
                        del arraySortIndices
                        del arraySelecao
            newName = "_Fill_" + method + ".csv"
            df = pd.DataFrame(arrayDataP, columns = colNames, index = indexList)
            df.to_csv(fileCSV.replace(".csv",newName))
            self.iface.messageBar().pushMessage(u'Filling Gaps Done!!', level=Qgis.Info, duration=5)


            pass