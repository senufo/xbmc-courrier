# -*- coding: utf-8 -*-
"""
xbmc script for read mail on IMAP/POP server
"""
#Script pour consulter ses mails
#Senufo, 2011 (c)
#
# Date : mercredi 30 novembre 2011, 19:14:13 (UTC+0100)
# $Author: Senufo $
##Modules xbmc
import xbmc, xbmcgui
import xbmcaddon
import os, re
from BeautifulSoup import *

#import BeautifulSoup
from re import compile as Pattern

__author__     = "Senufo"
__scriptid__   = "script.mail"
__scriptname__ = "Mail"

__addon__      = xbmcaddon.Addon(__scriptid__)

__cwd__        = __addon__.getAddonInfo('path')
__version__    = __addon__.getAddonInfo('version')
__language__   = __addon__.getLocalizedString

__profile__    = xbmc.translatePath( __addon__.getAddonInfo('profile') )
__resource__   = xbmc.translatePath( os.path.join( __cwd__, 'resources', 'lib' ) )


sys.path.append (__resource__)

import sys
import time
from time import gmtime, strftime
import poplib, imaplib
import string


import email
from email.Parser import Parser as EmailParser
from email.utils import parseaddr
from email.Header import decode_header
#import mimetypes

#Script html2text.py dans resources/lib
from html2text import *
#Utilise le fichier de configuration de service notifier si il existe
try:
    Addon = xbmcaddon.Addon('service.notifier')
    #On vérifie que le fichier de configuration existe
    #si il n'existe pas on charge le fichier de config de mail
    if not (Addon.getSetting( 'name1' )):
        Addon = xbmcaddon.Addon('script.mail')
except:
    Addon = xbmcaddon.Addon('script.mail')
#Pour les messages traduit on utilise ceux de script.mail
Addon_traduc = xbmcaddon.Addon('script.mail')
#get actioncodes from keymap.xml/ keys.h
ACTION_PREVIOUS_MENU = 10
ACTION_SELECT_ITEM = 7
ACTION_MOVE_LEFT       =   1
ACTION_MOVE_RIGHT      =   2
ACTION_MOVE_UP         =   3
ACTION_MOVE_DOWN       =   4
ACTION_PAGE_UP         =   5
ACTION_PAGE_DOWN       =   6
ACTION_NUMBER1         =   59
ACTION_NUMBER2         =   60
ACTION_VOLUME_UP       =   88
ACTION_VOLUME_DOWN     =   89
ACTION_REWIND          =   78
ACTION_FASTFORWARD     =   77

#ID des boutons dans script-mail-main.xml
STATUS_LABEL = 100
NX_MAIL      = 101
MSG_BODY     = 102
EMAIL_LIST   = 120
SCROLL_BAR   = 121
MSG_BODY     = 102
SERVER1      = 1001
SERVER2      = 1002
SERVER3      = 1003
QUIT         = 1004
FILE_ATT     = 1005
MAX_SIZE_MSG = int(Addon.getSetting( 'max_msg_size' ))
SEARCH_PARAM = Addon.getSetting( 'search_param' )

class MailWindow(xbmcgui.WindowXML):
    """
    Display main window for read mail
    """
    def __init__(self, *args, **kwargs):

        #variable pour position dans le msg
        self.position = 0

    def onInit( self ):
        """
        Initialize parameters
        """

        self.getControl( EMAIL_LIST ).reset()
        for i in [1, 2, 3]:
            id = 'name' + str(i)
            NOM =  Addon.getSetting( id )
            Button_Name = 1000 + i
            if NOM:
                self.getControl( Button_Name ).setLabel( NOM )
            else:
                self.getControl( Button_Name ).setEnabled(False)
        self.checkEmail(Addon.getSetting( 'name1' ))



#Verifie les mails et affiche les sujets et expediteurs
#Alias etant le nom du serveur POP ou IMAP
    def checkEmail(self, alias):
        """
        Check mail on POP or IMAP server
        """
        #print 'ALIAS = %s ' % alias
        self.getControl( STATUS_LABEL ).setLabel( '%s ...' % alias )

        #Vide la liste des sudjet des messages
        #self.listControl.reset()
        self.USER = ''
        self.NOM = ''
        self.SERVER = ''
        self.PASSWORD = ''
        self.PORT = ''
        self.SSL = ''
        self.TYPE = ''
        self.FOLDER = ''
        for i in [1, 2, 3]:
            NOM = Addon.getSetting( 'name%i' % i )
            USER = Addon.getSetting( 'user%i' % i )
            NOM =  Addon.getSetting( 'name%i' % i )
            SERVER = Addon.getSetting( 'server%i' % i )
            PASSWORD =  Addon.getSetting( 'pass%i' % i )
            PORT =  Addon.getSetting( 'port%i' % i )
            SSL = Addon.getSetting( 'ssl%i' % i ) == "true"
            TYPE = Addon.getSetting( 'type%i' % i )
            FOLDER = Addon.getSetting( 'folder%i' % i )
            #On cherche le serveur selectionne
            if (alias == NOM):
                self.NOM = NOM
                self.SERVER = SERVER
                self.USER = USER
                self.PORT = PORT
                self.PASSWORD = PASSWORD
                self.TYPE = TYPE
                self.SSL = SSL
                self.FOLDER = FOLDER
                try:
                    #Partie POP3
                    if '0' in self.TYPE:  #'POP'
                        self.getPopMails()
                    if '1' in self.TYPE: #IMAP
                        self.getImapMails()
                except Exception, e:
                    print str( e )
                    dialog = xbmcgui.DialogProgress()
                    dialog.create(Addon_traduc.getLocalizedString(id=614),
                          Addon_traduc.getLocalizedString(id=620) % self.SERVER)
                    time.sleep(5)
                    dialog.close()

    def getPopMails(self):
        """
        Get mail on POP server with poplib
        """
        dialog = xbmcgui.DialogProgress()
        dialog.create(Addon_traduc.getLocalizedString(id=614),
                      Addon_traduc.getLocalizedString(id=610))#Inbox, Logging in
        if self.SSL:
            mail = poplib.POP3_SSL(str(self.SERVER), int(self.PORT))
        else:  #'POP3'
            mail = poplib.POP3(str(self.SERVER), int(self.PORT))
        mail.user(str(self.USER))
        mail.pass_(str(self.PASSWORD))
        numEmails = mail.stat()[0]

        print "You have", numEmails, "emails"
        #Affiche le nombre de msg
        self.getControl( NX_MAIL ).setLabel( '%d msg(s)' % numEmails )
        dialog.close()
        if numEmails == 0:
            dialogOK = xbmcgui.Dialog()
            dialogOK.ok("%s" % self.NOM,
                        Addon_traduc.getLocalizedString(id=612)) #no mail
            self.getControl( EMAIL_LIST ).reset()
        else:             #Inbox                           #You have
                                #emails
            dialog.create(Addon_traduc.getLocalizedString(id=613),
                Addon_traduc.getLocalizedString(id=615) + str(numEmails) + Addon_traduc.getLocalizedString(id=616))
            ##Retrieve list of mails
            resp, items, octets = mail.list()
            print "resp %s, %s " % (resp, items)
            dialog.close()
            #On recupere tous les messages pour les afficher
            progressDialog = xbmcgui.DialogProgress()
                              #Message(s)                       #Get mail
            progressDialog.create(Addon_traduc.getLocalizedString(id=617),
                                  Addon_traduc.getLocalizedString(id=618))
            i = 0
            #Mise a zero de la ListBox msg
            self.getControl( EMAIL_LIST ).reset()
            self.emails = []
            for item in items:
                i = i + 1
                id, size = string.split(item)
                up = (i*100)/numEmails    #Get mail             Please wait
                progressDialog.update(up,
                                      Addon_traduc.getLocalizedString(id=618),
                                      Addon_traduc.getLocalizedString(id=619))

                #Si dépasse la taille max on télécharge que 50 lignes
                if (MAX_SIZE_MSG == 0) or (size < MAX_SIZE_MSG):
                    resp, text, octets = mail.retr(id)
                else:
                    resp, text, octets = mail.top(id, 300)
                att_file = ':'
                text = string.join(text, "\n")
                self.processMails(text, att_file)
            progressDialog.close()
            #Affiche le 1er mail de la liste
            self.getControl( EMAIL_LIST ).selectItem(0)

    def processMails(self, text, att_file):
        """
        Parse mail for display in XBMC
        """
        myemail = email.message_from_string(text)
        p = EmailParser()
        msgobj = p.parsestr(text)
        if msgobj['Subject'] is not None:
            decodefrag = decode_header(msgobj['Subject'])
            subj_fragments = []
            for s , enc in decodefrag:
                if enc:
                    s = unicode(s , enc).encode('utf8','replace')
                subj_fragments.append(s)
            subject = ''.join(subj_fragments)
        else:
            subject = None
        if msgobj['Date'] is not None:
            date = msgobj['Date']
        else:
            date = '--'
        Sujet = subject
        realname = parseaddr(msgobj.get('From'))[1]

        body = None
        html = None
        for part in msgobj.walk():
            content_disposition = part.get("Content-Disposition", None)
            prog = re.compile('attachment')
            #Retrouve le nom des fichiers attaches
            if prog.search(str(content_disposition)):
                file_att = str(content_disposition)

                pattern = Pattern(r"\"(.+)\"")
                att_file +=  str(pattern.findall(file_att))

            if part.get_content_type() == "text/plain":
                if body is None:
                    body = ""
                try :
                    #Si pas de charset défini
                    if (part.get_content_charset() is None):
                        body +=  part.get_payload(decode=True)
                    else:
                        body += unicode(
                           part.get_payload(decode=True),
                           part.get_content_charset(),
                           'replace'
                           ).encode('utf8','replace')
                except Exception, e:
                    body += "Erreur unicode"
                    print "BODY = %s " % body
            elif part.get_content_type() == "text/html":
                if html is None:
                    html = ""
                try :
                    unicode_coded_entities_html = unicode(BeautifulStoneSoup(html,
                            convertEntities=BeautifulStoneSoup.HTML_ENTITIES))

                    html += unicode_coded_entities_html
                    html = html2text(html)
                except Exception, e:
                    html += "Erreur unicode html"
                    #print "HTML = %s " % html
            realname = parseaddr(msgobj.get('From'))[1]
        Sujet = subject
        description = ' '
        if (body):
            description = str(body)
        else:
            try:
                html = html.encode('ascii','replace')
                description = str(html)
            except Exception, e:
                print str(e)
        #Nb de lignes du msg pour permettre le scroll text
        self.nb_lignes = description.count("\n")

        listitem = xbmcgui.ListItem( label2=realname, label=Sujet)
        listitem.setProperty( "realname", realname )
        date += att_file
        listitem.setProperty( "date", date )
        listitem.setProperty( "message", description )
        self.getControl( EMAIL_LIST ).addItem( listitem )

    def getImapMails(self):
        """
        Get amil form IMAP server
        """
        dialog = xbmcgui.DialogProgress()
        dialog.create(Addon_traduc.getLocalizedString(id=614),
                      Addon_traduc.getLocalizedString(id=610))#Inbox,Logging in
        #Mise a zero de la ListBox msg
        #self.getControl( EMAIL_LIST ).reset()
        self.emails = []
        try:
            if self.SSL:
                imap = imaplib.IMAP4_SSL(str(self.SERVER), int(self.PORT))
            else:
                imap = imaplib.IMAP4(str(self.SERVER), int(self.PORT))
            att_file = ':'
            imap.login(self.USER, self.PASSWORD)
            imap.select(self.FOLDER)
            #numEmails = len(imap.search(None, 'UnSeen')[1][0].split())
            numEmails = len(imap.search(None, SEARCH_PARAM )[1][0].split())
            #print "You have", numEmails, "emails IMAP"
            #Affiche le nombre de msg
            self.getControl( NX_MAIL ).setLabel( '%d msg(s)' % numEmails )
            dialog.close()
            if numEmails == 0:
                dialogOK = xbmcgui.Dialog()
                dialogOK.ok("%s" % self.NOM,
                            Addon_traduc.getLocalizedString(id=612)) #no mail
                self.getControl( EMAIL_LIST ).reset()
            else:
                progressDialog2 = xbmcgui.DialogProgress()
                                  #Message(s)                       #Get mail
                progressDialog2.create(Addon_traduc.getLocalizedString(id=617),
                                       Addon_traduc.getLocalizedString(id=618))
                i = 0
        ##Retrieve list of mails
                typ, data = imap.search(None, SEARCH_PARAM)
                for num in data[0].split():
                    i = i + 1
                    typ, data = imap.fetch(num, '(RFC822)')
                    up = (i*100)/numEmails    #Get mail              Please wait
                    progressDialog2.update(up,
                                       Addon_traduc.getLocalizedString(id=618),
                                       Addon_traduc.getLocalizedString(id=619))
                    print "UP = %d " % up
                    text = data[0][1].strip()
                    self.processMails(text, att_file)
                progressDialog2.close()
            #Affiche le 1er mail de la liste
            self.getControl( EMAIL_LIST ).selectItem(0)
            imap.logout
        except Exception, e:
            print str( e )
            print 'IMAP exception'


    def onAction(self, action):
        #print "ID Action %d" % action.getId()
        #print "Code Action %d" % action.getButtonCode()
        if action == ACTION_PREVIOUS_MENU:
            self.close()
        if action == ACTION_MOVE_UP:
            controlId = action.getId()
        if action == ACTION_MOVE_DOWN:
            controlId = action.getButtonCode()
        if action == ACTION_FASTFORWARD: #PageUp
            if (self.position > 0):
                self.position = self.position - 1
            self.getControl( MSG_BODY ).scroll(self.position)
            print "Position F = %d " % self.position
        if (action == ACTION_REWIND): #PageUp
            if (self.position <= self.nb_lignes):
                self.position = self.position + 1
            self.getControl( MSG_BODY ).scroll(self.position)
            print "Position R = %d " % self.position

    def onClick( self, controlId ):
        #print "onClick controId = %d " % controlId
        if (controlId in [SERVER1, SERVER2, SERVER3]):
            label = self.getControl( controlId ).getLabel()
            self.checkEmail(label)
        elif (controlId == QUIT):
            self.close()

mydisplay = MailWindow( "script-mail-main.xml" , __cwd__, "Default")
mydisplay .doModal()
del mydisplay
