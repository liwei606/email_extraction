#!/usr/bin/python

import sys,os,re
import time
import urllib2
import MySQLdb
from BeautifulSoup import BeautifulSoup

class pipermailParser:
        def __init__(self):
                self.mailURL = "http://linux.intel.com/pipermail/"
                self.DBusername = "root"
                self.DBpassword = "cox-eb-b"
                self.DBname = "pipermail_data"
        def start(self,project,status):
                #Fetch Content
                print "Begin Sync MailList [" + str(project) + "] Status:" + str(status)
                HomePageContent = self.fetchHTML(self.mailURL + str(project) + "/")
                HPresults = self.AnalysisHomePageContent(HomePageContent,project,status)
                MailList = self.AnalysisMailPage(project,HPresults)
                print "Finish Sync MailList [" + str(project) + "] Status:" + str(status)
                return True
        def fetchHTML(self,HTMLURL):
                try:
                        HTMLcontent = urllib2.urlopen(HTMLURL).read()
                except:
                        print "Could not connect to Web Page. Try again!"
                        sys.exit(-1)
                return HTMLcontent
        def AnalysisHomePageContent(self,HTMLcontent,project,status):
                soup = BeautifulSoup(HTMLcontent)
                trData = soup.html.body.table.findAll('tr')
                LastDate = self.fetchLastDate(project)
                resultlist = []
                for trInfo in trData:
                        matchPar = re.compile('[A-Z][a-z]{2,10} \d{4}:')
                        match = matchPar.match(trInfo.td.string)
                        result = {}
                        if match:
                                Mlen = len(trInfo.td.string)
                                result['Date'] = trInfo.td.string[:(Mlen-1)]
                                result['URL'] = (trInfo.findAll('td')[1]).a['href']
                                if LastDate == "None" or status == "start":
                                        resultlist.append(result)
                                else:
                                        DateNow = time.strptime(result['Date'],'%B %Y')
                                        DateDB = time.strptime(LastDate,'%B %Y')
                                        if not DateDB > DateNow:
                                                resultlist.append(result)
                return resultlist
        def AnalysisMailPage(self,project,PageURLList):
                for PageURL in PageURLList:
                        MailPageUrl = self.mailURL + str(project) + "/" + PageURL['URL']
                        print MailPageUrl
                        MailPageContent = self.fetchHTML(MailPageUrl)
                        MailList = self.AnalysisMailPageContent(MailPageContent)
                        self.AnalysisMailList(MailPageUrl,MailList,project)
                        #print MailList
                return True
        def AnalysisMailPageContent(self,MailPageContent):
                soup = BeautifulSoup(MailPageContent)
                liData = soup.html.body.findAll('li')
                MailList = []
                for Data in liData:
                        if Data.a:
                                matchPar = re.compile('\d.*.html')
                                match = matchPar.match(Data.a['href'])
                                if match:
                                        MailInfo = {}
                                        MailInfo['Subject'] = Data.a.string.rstrip('\n')
                                        MailInfo['Sender'] = Data.i.string.rstrip('\n')
                                        MailInfo['URL'] = Data.a['href']
                                        MailList.append(MailInfo)
                return MailList
        def AnalysisMailList(self,MailPageUrl,MailList,project):
                for Mail in MailList:
                        MailContent = self.fetchHTML(MailPageUrl.replace("thread.html",Mail['URL']))
                        print "[" + MailPageUrl.replace("thread.html",Mail['URL']) + "]"
                        MailLink = MailPageUrl.replace("thread.html",Mail['URL'])
                        Result = self.AnalysisMailContent(MailContent)
                        SyncResult = self.SyncDB(Result,Mail,project,MailLink)
                return True
        def AnalysisMailContent(self,MailContent):
                soup = BeautifulSoup(MailContent)
                MailContentInfo = {}
                MailContentInfo['CreateOn'] = soup.html.body.i.string[:20] + soup.html.body.i.string[24:28]
                MailContentInfo['Text'] = soup.html.body.pre.text
                return MailContentInfo
        def SyncDB(self,Result,Mail,project,MailLink):
                try:
                        connection = MySQLdb.connect(user=self.DBusername, passwd=self.DBpassword,host="localhost",db=self.DBname)
                except:
                        print "Could not connect to MySQL server!"
                        sys.exit(-1)
                cursor = connection.cursor()
                TempSubject = Mail['Subject'].replace("'","\\'")
                TextT = Result['Text'].replace("'","\\'")
                CommitFlag = TextT.find('[This message was auto-generated]',0)
                if CommitFlag > 0:
                        TempText = TextT[:CommitFlag]
                else:
                        TempText = TextT
                TempProject = project.replace("-","")
                querySql = "select Subject,Sender,CreateOn from " + str(TempProject) + " where Subject = \
                '" + str(TempSubject) + "' and CreateOn = str_to_date('" + str(Result['CreateOn']) + "','%a %b %d %T %Y')"
                cursor.execute(querySql)
                queryresult = cursor.fetchone()
                if not queryresult:
                        print "Begin Sync Mail: " + str(Mail['Subject']) + " [" + str(Mail['Sender']) + "]"
                        sql = "insert into " + str(TempProject) + " (`Subject`,`Sender`,`CreateOn`,`Text`,`MailLink`) values \
                        ('" + str(TempSubject) + "','" + str(Mail['Sender']) + "',\
                        str_to_date('" + str(Result['CreateOn']) + "','%a %b %d %T %Y'),'" + str(TempText) + "',\
                        '" + str(MailLink) + "')"
                        try:
                                cursor.execute(sql)
                        except:
                                print sql
                                print "******************************"
                                print TempText
                                exit(-1)
                cursor.close()
                connection.commit()
                connection.close()
                return True
        def fetchLastDate(self,project):
                try:
                        connection = MySQLdb.connect(user=self.DBusername, passwd=self.DBpassword,host="localhost",db=self.DBname)
                except:
                        print "Could not connect to MySQL server!"
                        sys.exit(-1)
                cursor = connection.cursor()
                querySql = "select date_format(max(CreateOn),'%M %Y') from " + str(project.replace("-",""))
                cursor.execute(querySql)
                queryresult = cursor.fetchone()
                LastDate = queryresult[0]
                cursor.close()
                connection.commit()
                connection.close()
                return LastDate
if __name__ == '__main__':
        if len(sys.argv) == 1 or sys.argv[1] == "start":
                parser = pipermailParser()
                #Sync MailList Tzivi
                parser.start("tzivi","start")
                #Sync MailList automotive
                parser.start("automotive","start")
                #Sync MailList tizen-partners-commits
                parser.start("tizen-partners-commits","start")
                #Sync MailList tzmobile
                parser.start("tzmobile","start")
                #Sync MailList tzmobile
                parser.start("tizen-commits","start")
        elif sys.argv[1] == "quick":
                parser = pipermailParser()
                #Sync MailList Tzivi
                parser.start("tzivi","quick")
                #Sync MailList automotive
                parser.start("automotive","quick")
                #Sync MailList tizen-partners-commits
                parser.start("tizen-partners-commits","quick")
                #Sync MailList tzmobile
                parser.start("tzmobile","quick")
                #Sync MailList tzmobile
                parser.start("tizen-commits","quick")
        else:
                print "Usage: " + str(sys.argv[0]) + " [OPTION] ..."
                print "[OPTION]"
                print "\tstart\tSync a New MailList"
                print "\tquick\tSync an exist MailList"
