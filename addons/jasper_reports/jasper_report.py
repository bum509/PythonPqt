##############################################################################
import os
#import csv
#import copy
#import base64
import report
import pooler
from osv import orm, osv, fields
import tools
#import tempfile 
#import codecs
#import sql_db
import tempfile
import netsvc
import release
import logging

from report import report_sxw

from JasperReports import *

# Determines the port where the JasperServer process should listen with its XML-RPC server for incomming calls
tools.config['jasperport'] = tools.config.get('jasperport', 8090)

# Determines the file name where the process ID of the JasperServer process should be stored
tools.config['jasperpid'] = tools.config.get('jasperpid', 'openerp-jasper.pid')

# Determines if temporary files will be removed
tools.config['jasperunlink'] = tools.config.get('jasperunlink', True)

class Report(report_sxw.rml_parse):
    def __init__(self, name, cr, uid, ids, data, context):
        self.name = name
        self.cr = cr
        self.uid = uid
        self.ids = ids
        self.data = data
        self.model = self.data['model']
#		self.model = self.data.get('model', False) or context.get('active_model', False)
        self.context = context or {}
        self.pool = pooler.get_pool( self.cr.dbname )
        self.reportPath = None
        self.report = None
        self.temporaryFiles = []
#        self.outputFormat = 'pdf'
        if self.data.get('form'):#If is menu report type view
            self.outputFormat = str(self.data['form']['used_context']['format'])
        else:
            self.outputFormat = 'pdf'
    def addonsPath(self, path=False):
        if path:
            report_module = path.split(os.path.sep)[0]
            for addons_path in tools.config['addons_path'].split(','):
                if os.path.lexists(addons_path+os.path.sep+report_module):
                    return os.path.normpath( addons_path+os.path.sep+path )
                    
        return os.path.abspath(os.path.dirname(__file__))
    def execute(self):
        """
        If self.context contains "return_pages = True" it will return the number of pages
        of the generated report.
        """
        logger = netsvc.Logger()
#		logger = logging.getLogger(__name__)

        # * Get report path *
        # Not only do we search the report by name but also ensure that 'report_rml' field
        # has the '.jrxml' postfix. This is needed because adding reports using the <report/>
        # tag, doesn't remove the old report record if the id already existed (ie. we're trying
        # to override the 'purchase.order' report in a new module). As the previous record is
        # not removed, we end up with two records named 'purchase.order' so we need to destinguish
        # between the two by searching '.jrxml' in report_rml.
        ids = self.pool.get('ir.actions.report.xml').search(self.cr, self.uid, [('report_name', '=', self.name[7:]),('report_rml','ilike','.jrxml')], context=self.context)
        data = self.pool.get('ir.actions.report.xml').read(self.cr, self.uid, ids[0], ['report_rml','jasper_output'])
        #TT Get jasper_output from input screen(if nothing)
        if not self.outputFormat:
            if data['jasper_output']:
                self.outputFormat = data['jasper_output']
            else:
                self.outputFormat = 'pdf'
        self.reportPath = data['report_rml']     
        self.reportPath = self.reportPath.replace('jasper_reports/','')
        self.reportPath = self.reportPath.replace('/','\\')
        self.reportPath = os.path.join( self.addonsPath(path=False), self.reportPath )
#        if not os.path.lexists(self.reportPath):
#        	self.reportPath = self.addonsPath(path=data['report_rml'])

        # Get report information from the jrxml file
        logger.notifyChannel("jasper_reports", netsvc.LOG_INFO, "Requested report: '%s'" % self.reportPath)
#		logger.info("Requested report: '%s'" % self.reportPath)
        self.report = JasperReport( self.reportPath )

        # Create temporary input (XML) and output (PDF) files 
        fd, dataFile = tempfile.mkstemp()
        os.close(fd)
        fd, outputFile = tempfile.mkstemp()
        os.close(fd)
        self.temporaryFiles.append( dataFile )
        self.temporaryFiles.append( outputFile )
        logger.notifyChannel("jasper_reports", netsvc.LOG_INFO, "Temporary data file: '%s'" % dataFile)
#		logger.info("Temporary data file: '%s'" % dataFile)

        import time
        start = time.time()

        # If the language used is xpath create the xmlFile in dataFile.
        if self.report.language() == 'xpath':
            if self.data.get('data_source','model') == 'records':
                generator = CsvRecordDataGenerator(self.report, self.data['records'] )
            else:
                generator = CsvBrowseDataGenerator( self.report, self.model, self.pool, self.cr, self.uid, self.ids, self.context )
            generator.generate( dataFile )
            self.temporaryFiles += generator.temporaryFiles
        
        subreportDataFiles = []
        for subreportInfo in self.report.subreports():
            subreport = subreportInfo['report']
            if subreport.language() == 'xpath':
                message = 'Creating CSV '
                if subreportInfo['pathPrefix']:
                    message += 'with prefix %s ' % subreportInfo['pathPrefix']
                else:
                    message += 'without prefix '
                message += 'for file %s' % subreportInfo['filename']
                logger.notifyChannel("jasper_reports", netsvc.LOG_INFO, message)
#				logger.info("%s" % message)

                fd, subreportDataFile = tempfile.mkstemp()
                os.close(fd)
                subreportDataFiles.append({
                    'parameter': subreportInfo['parameter'],
                    'dataFile': subreportDataFile,
                    'jrxmlFile': subreportInfo['filename'],
                })
                self.temporaryFiles.append( subreportDataFile )

                if subreport.isHeader():
                    generator = CsvBrowseDataGenerator( subreport, 'res.users', self.pool, self.cr, self.uid, [self.uid], self.context )
                elif self.data.get('data_source','model') == 'records':
                    generator = CsvRecordDataGenerator( subreport, self.data['records'] )
                else:
                    generator = CsvBrowseDataGenerator( subreport, self.model, self.pool, self.cr, self.uid, self.ids, self.context )
                generator.generate( subreportDataFile )
                

        # Call the external java application that will generate the PDF file in outputFile
        pages = self.executeReport( dataFile, outputFile, subreportDataFiles )
        elapsed = (time.time() - start) / 60
        logger.notifyChannel("jasper_reports", netsvc.LOG_INFO, "ELAPSED: '%f'" % elapsed )
#		logger.info("ELAPSED: %f" % elapsed)

        # Read data from the generated file and return it
        f = open( outputFile, 'rb')
        try:
            data = f.read()
        finally:
            f.close()

        # Remove all temporary files created during the report
        if tools.config['jasperunlink']:
            for file in self.temporaryFiles:
                try:
                    os.unlink( file )
                except os.error, e:
#                    logger = netsvc.Logger()
                    logger.notifyChannel("jasper_reports", netsvc.LOG_WARNING, "Could not remove file '%s'." % file )
                    #logger.warning("Could not remove file '%s'." % file )
        self.temporaryFiles = []

        if self.context.get('return_pages'):
            return ( data, self.outputFormat, pages )
        else:
            return ( data, self.outputFormat )

    def path(self):
        return os.path.dirname( self.path() )
#    def addonsPath(self):
#	def addonsPath(self, path=False):
#		if path:
#			report_module = path.split(os.path.sep)[0]
#			for addons_path in tools.config['addons_path'].split(','):
#				if os.path.lexists(addons_path+os.path.sep+report_module):
#					return os.path.normpath( addons_path+os.path.sep+path )
#					
#        return os.path.abspath(os.path.dirname(__file__))

    def systemUserName(self):
        if os.name == 'nt':
            import win32api
            return win32api.GetUserName()
        else:
            import pwd
            return pwd.getpwuid(os.getuid())[0]

    def dsn(self):
        host = tools.config['db_host'] or 'localhost'
        port = tools.config['db_port'] or '5432'
        dbname = self.cr.dbname
        return 'jdbc:postgresql://%s:%s/%s' % ( host, port, dbname )
    
    def userName(self):
        return tools.config['db_user'] or self.systemUserName()

    def password(self):
        return tools.config['db_password'] or ''

    def executeReport(self, dataFile, outputFile, subreportDataFiles):    
        isNoPager = False
        strLocal = 'vi_VN'
        if self.outputFormat == 'xls':#TT Add to export
             isNoPager = True
             #strLocal = 'en_US'                              
        connectionParameters = {
            'output': self.outputFormat,
            #'xml': dataFile,
            'csv': dataFile,
            'dsn': self.dsn(),
            'user': self.userName(),
            'password': self.password(),
            'subreports': subreportDataFiles,
        }
        parameters = {
            'STANDARD_DIR': self.report.standardDirectory(),                        
            #'REPORT_LOCALE': 'vi_VN',
            'REPORT_LOCALE': strLocal,
            'IDS': self.ids,
            'IS_IGNORE_PAGINATION': isNoPager, #True,#28/6/2013(Add to disable view Header every pages)
        }
        if 'parameters' in self.data:
            parameters.update( self.data['parameters'] )        
        #TT Add parameters for report (Menu report type)                        
        if self.data.get('form'):   
            if 'used_context' in self.data['form']:
                company = self.pool.get('account.account').browse(self.cr, self.uid, self.data['form']['chart_account_id'], context=self.context).company_id
                com_id = company.id
                com_name = company.name
                com_add = company.street
                if company.partner_id.vat:
                    com_tax = company.partner_id.vat
                else:
                    com_tax = ''       
                com_city = company.city 
                com_web = company.website
                self.data['form']['used_context']['com_name'] = com_name
                self.data['form']['used_context']['com_add'] = com_add
                self.data['form']['used_context']['com_tax'] = com_tax
                self.data['form']['used_context']['com_phone'] = company.phone and company.phone or False
                self.data['form']['used_context']['com_fax'] = company.fax and company.fax or False
                self.data['form']['used_context']['com_web'] = company.website and company.website or False
                
                if 'acc_analytic_id' in self.data['form']['used_context']:
                    if self.data['form']['used_context']['acc_analytic_id'] != 0:
                        analytic = self.pool.get('account.analytic.account').browse(self.cr, self.uid, self.data['form']['used_context']['acc_analytic_id'] , context=self.context)
                        self.data['form']['used_context']['ala_name'] = analytic.name
                    else:
                        self.data['form']['used_context']['ala_name'] = 'All'                
                parameters.update( self.data['form']['used_context'] )
        #If report in form for print (document type view)
        #Case for list of report define
        #if self.name == 'report.Jasper.bank.PayReceive':if self.name == 'report.Jasper.bank.document':
        if self.data.get('id'):
            company = self.pool.get(self.model).browse(self.cr, self.uid, self.data.get('id'), context=self.context).company_id
            com_id = company.id
            com_name = company.name
            com_add = company.street
            if company.partner_id.vat:
                com_tax = company.partner_id.vat
            else:
                com_tax = ''        
            com_city = company.city
            com_web = company.website
            self.data['com_name'] = com_name
            self.data['com_add'] = com_add
            self.data['com_tax'] = com_tax
            strNumbertoText = ''
            switcher = {# Calculate sum of document
                'report.Jasper.arinvoice.document':'select SUM(out_debit) as SUM from rpt_pkt_ar_document(' + str(self.data.get('id')) + ')',
                'report.Jasper.arinvoice.invoice' :'select (SUM(out_amount) + SUM(out_tax_amount)) as SUM from rpt_pkt_ar_invoice(' + str(self.data.get('id')) + ')',
                'report.Jasper.apinvoice.document':'select SUM(out_debit) as SUM from rpt_pkt_ar_document(' + str(self.data.get('id')) + ')',
                'report.Jasper.arpayment.document':'select SUM(out_debit) as SUM from rpt_pkt_ar_payment_document(' + str(self.data.get('id')) + ')',
                'report.Jasper.appayment.document':'select SUM(out_debit) as SUM from rpt_pkt_ap_payment_document(' + str(self.data.get('id')) + ')',
                'report.Jasper.bank.document'     :'select SUM(out_debit) as SUM from rpt_pkt_bank_document(' + str(self.data.get('id')) + ')',
                'report.Jasper.bank.PayReceive'   :'select SUM(out_amount) as SUM from rpt_pkt_bank_pt_pc(' + str(self.data.get('id')) + ')',
                'report.Jasper.Expense.phieuchi'  :'select SUM(out_amount) as SUM from rpt_pkt_expense_pc(' + str(self.data.get('id')) + ')',
                'report.Jasper.Expense.document'  :'select SUM(out_debit) as SUM from rpt_pkt_expense_document(' + str(self.data.get('id')) + ')',
                'report.Jasper.inreceive.document':'select SUM(out_values) as SUM from rpt_pkt_in_phieunhap(' + str(self.data.get('id')) + ')',
                'report.Jasper.inissues.document' :'select SUM(out_values) as SUM from rpt_pkt_in_phieuxuat(' + str(self.data.get('id')) + ')',
                'report.Jasper.adjreceive.document':'select SUM(out_values) as SUM from rpt_pkt_in_phieuadj(' + str(self.data.get('id')) + ')',
                'report.Jasper.accountentries.document':'select SUM(out_debit) as SUM from rpt_pkt_gl_document(' + str(self.data.get('id')) + ')',
                }
            sql = switcher.get(self.name, "nothing")
            if sql!='nothing':
                self.cr.execute(sql)
                res = self.cr.fetchone()
                strNumbertoText = self.numbers_to_text(res[0])
            else:
                strNumbertoText='Không đồng'
            self.data['text_sum'] = strNumbertoText            
            parameters.update(self.data)    
        if 'date_from' not in parameters:
            strDate = str(time.strftime('%Y-01-01')) 
            date_from = strDate[8:10] + '/' + strDate[5:7] + '/' + strDate[0:4]
            parameters.update({'date_from':date_from})    
        if 'date_to' not in parameters:
            strDate = str(time.strftime('%Y-%m-%d')) 
            date_to = strDate[8:10] + '/' + strDate[5:7] + '/' + strDate[0:4]
            parameters.update({'date_to':date_to})  
        server = JasperServer( int( tools.config['jasperport'] ) )
        server.setPidFile( tools.config['jasperpid'] )
        return server.execute( connectionParameters, self.reportPath, outputFile, parameters )
    
    def numbers_to_text(cr, SoInput):
            KHONG_DONG = 'không đồng'
            SO_LON = 'số quá lớn'
    
            TRU = 'trừ'
            TRU_1 = 'âm'
            
            TRAM = 'trăm'
            MUOI = 'mươi'
            MUOI_1 = 'mười'
            GIDO = 'gì đó'
    
            NGANTY = 'ngàn tỷ'
            TY = 'tỷ'
            TRIEU = 'triệu'
            NGAN = 'ngàn'
            DONG = 'đồng'
            XU = 'xu'
    
            MOT = 'một'
            HAI = 'hai'
            BA = 'ba'
            BON = 'bốn'
            NAM = 'năm'
            SAU = 'sáu'
            BAY = 'bảy'
            TAM = 'tám'
            CHIN = 'chín'
    
            CHAN = 'chẵn'
            LE = 'lẽ'
            MUOIMOT = 'mười một'
            MUOIMOTT = 'mươi mốt'
            Space0 = ''
            Space1 = ' '
            Space2 = '  '
            Space3 = '   '
            KetQua = ''
            if long(SoInput) == 0:
                formula = KHONG_DONG
            else:
                if abs(long(SoInput)) >= 1000000000000000:
                    formula = SO_LON
                else:
                    if long(SoInput) <= 0:
                        KetQua = TRU_1 + Space1
    #                SoTien = ToText(abs(long(SoInput)), "####################0.00")
    #                SoTien = Right(Space(15) & SoTien, 18)
                    SoTien = str(abs(long(SoInput))) + '.00'
                    dodai = len(SoTien)
                    if dodai < 18:
                        i = 0
                        while i <= (18 - dodai):
                            i = i + 1
                            SoTien = Space1 + SoTien
                    Hang = [TRAM, MUOI, GIDO]
                    Doc = [NGANTY, TY, TRIEU, NGAN, DONG, XU]
                    Dem = [MOT, HAI, BA, BON, NAM, SAU, BAY, TAM, CHIN]
                    lap1 = [1,2,3,4,5,6]
                    lap2 = [1,2,3]
                    for i in lap1:
                        Nhom = SoTien[i * 3 - 2:i * 3 + 1]#Mid(SoTien, i * 3 - 2, 3)
                        if Nhom <> Space3:#If Nhom <> Space(3):
                            if Nhom == '000':
                                if i == 5:
                                    Chu = DONG + Space1#Chu = DONG & Space(1)
                                else:
                                    Chu = Space0#Chu = Space(0)                                
                            elif Nhom == '.00' or Nhom == ',00':
                                Chu = CHAN
                            else:
                                S1 = Nhom[0:1]#Left(Nhom, 1)
                                S2 = Nhom[1:2]#Mid(Nhom, 2, 1)
                                S3 = Nhom[len(Nhom)-1:len(Nhom)]#Right(Nhom, 1)
                                Chu = Space0#Chu = Space(0)
                                Hang[2] = Doc[i-1]#Hang[3] = Doc[i]                            
                                for j in lap2:
                                    Dich = Space0#Dich = Space(0)
                                    if Nhom[j-1:j] != Space1:
                                        S = long(Nhom[j-1:j])#S = Val(Mid(Nhom, j, 1))                                
                                        if S >0:#if S > 0:
                                            Dich = Dem[S-1] + Space1 + Hang[j-1] + Space1#Dich = Dem(S) & Space(1) & Hang(j) & Space(1)
                                        if j == 2:
                                            if S == 1:#if S == 1:                                            
                                                Dich = MUOI_1 + Space1#Dich = MUOI & Space(1)
                                            elif S == 0 and S3 != '0':#elif S == 0 and S3 != '0':
                                                if ((S1 >= '1') and (S1 <= '9')) or ((S1 == '0') and (i == 4)):
                                                    Dich = LE + Space1#Dich = LE & Space(1)
                                        if j == 3:
                                            if S == 0 and Nhom <> Space2 + '0':#if S = 0 and Nhom <> Space(2) & "0":
                                                Dich = Hang[j-1] + Space1#Dich = Hang(j) & Space(1)  
                                            if S == 5 and S2 <> Space1 and S2 <> '0':#if S = 5 and S2 <> Space(1) and S2 <> "0":
                                                Dich = 'l' + Dich[1:len(Dich)]#Dich = '1' & Mid(Dich, 2)
                                        Chu = Chu + Dich
                            #Vitri = InStr(1, Chu, MUOIMOT, 1)                        
                            KetQua = KetQua + Chu                        
            formula = (KetQua[0:1]).upper() + KetQua[1:len(KetQua)]                                       
            return formula  


class report_jasper(report.interface.report_int):
    def __init__(self, name, model, parser=None ):
        # Remove report name from list of services if it already
        # exists to avoid report_int's assert. We want to keep the 
        # automatic registration at login, but at the same time we 
        # need modules to be able to use a parser for certain reports.
        if release.major_version == '5.0':
            if name in netsvc.SERVICES:
                del netsvc.SERVICES[name]
        else:
            if name in netsvc.Service._services:
                del netsvc.Service._services[name]
        super(report_jasper, self).__init__(name)
        self.model = model
        self.parser = parser

    def create(self, cr, uid, ids, data, context):
        name = self.name
        if self.parser:
            d = self.parser( cr, uid, ids, data, context )
            ids = d.get( 'ids', ids )
            name = d.get( 'name', self.name )
            # Use model defined in report_jasper definition. Necesary for menu entries.
            data['model'] = d.get( 'model', self.model )
            data['records'] = d.get( 'records', [] )
            # data_source can be 'model' or 'records' and lets parser to return
            # an empty 'records' parameter while still executing using 'records'
            data['data_source'] = d.get( 'data_source', 'model' )
            data['parameters'] = d.get( 'parameters', {} )
        r = Report( name, cr, uid, ids, data, context )
        #return ( r.execute(), 'pdf' )
        return r.execute()

if release.major_version == '5.0':
    # Version 5.0 specific code

    # Ugly hack to avoid developers the need to register reports
    import pooler
    import report

    def register_jasper_report(name, model):
        name = 'report.%s' % name
        # Register only if it didn't exist another "jasper_report" with the same name
        # given that developers might prefer/need to register the reports themselves.
        # For example, if they need their own parser.
        if netsvc.service_exist( name ):
            if isinstance( netsvc.SERVICES[name], report_jasper ):
                return
            del netsvc.SERVICES[name]
        report_jasper( name, model )


    # This hack allows automatic registration of jrxml files without 
    # the need for developers to register them programatically.

    old_register_all = report.interface.register_all
    def new_register_all(db):
        value = old_register_all(db)

        cr = db.cursor()
        # Originally we had auto=true in the SQL filter but we will register all reports.
        cr.execute("SELECT * FROM ir_act_report_xml WHERE report_rml ilike '%.jrxml' ORDER BY id")
        records = cr.dictfetchall()
        cr.close()
        for record in records:
            register_jasper_report( record['report_name'], record['model'] )
        return value

    report.interface.register_all = new_register_all
else:
    # Version 6.0 and later

    def register_jasper_report(report_name, model_name):
        name = 'report.%s' % report_name
        # Register only if it didn't exist another "jasper_report" with the same name
        # given that developers might prefer/need to register the reports themselves.
        # For example, if they need their own parser.
        if name in netsvc.Service._services:
            if isinstance(netsvc.Service._services[name], report_jasper):
                return
            del netsvc.Service._services[name]
        report_jasper( name, model_name )

    class ir_actions_report_xml(osv.osv):
        _inherit = 'ir.actions.report.xml'

        def register_all(self, cr):
            # Originally we had auto=true in the SQL filter but we will register all reports.
            cr.execute("SELECT * FROM ir_act_report_xml WHERE report_rml ilike '%.jrxml' ORDER BY id")
            records = cr.dictfetchall()
            for record in records:
                register_jasper_report(record['report_name'], record['model'])
            return super(ir_actions_report_xml, self).register_all(cr)

    ir_actions_report_xml()

# vim:noexpandtab:smartindent:tabstop=8:softtabstop=8:shiftwidth=8:
