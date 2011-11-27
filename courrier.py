# -*- coding: utf-8 -*-
#Script pour consulter ses mails
#Senufo, 2011 (c)
#Version 0.0.6
#
# $Date: 2011-11-13 22:23:30 +0100 (dim. 13 nov. 2011) $
# $Author: Senufo $
##Modules xbmc
import xbmc, xbmcgui
import xbmcaddon
import os, re

__author__     = "Senufo"
__scriptid__   = "script.mail"
__scriptname__ = "Mail"

__addon__      = xbmcaddon.Addon(id=__scriptid__)

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
import string, random
import StringIO, rfc822


import email
from email.Parser import Parser as EmailParser
from email.utils import parseaddr
from email.Header import decode_header

import errno
import mimetypes

#Script html2text.py dans resources/lib
from html2text import *

Addon = xbmcaddon.Addon('script.mail')

#get actioncodes from keymap.xml/ keys.h
ACTION_PREVIOUS_MENU = 10
ACTION_SELECT_ITEM = 7
#ACTION_PARENT_DIR = 9
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
ACTION_REWIND          =   77
ACTION_FASTFORWARD     =   78

#ID des boutons dans myWin.xml
STATUS_LABEL   	= 100
NX_MAIL       	= 101
MSG_BODY	= 102
EMAIL_LIST     	= 120
SCROLL_BAR	= 121
MSG_BODY	= 102
SERVER1		= 1001
SERVER2		= 1002
SERVER3		= 1003
QUIT		= 1004
MAX_SIZE_MSG = 10000
class MailWindow(xbmcgui.WindowXML):
   
  def __init__(self, *args, **kwargs):

    #variable pour position dans le msg
    self.position = 0

  def onInit( self ):
    print "Branch  EXPERIMENTAL"

    self.getControl( EMAIL_LIST ).reset()
    for i in [1,2,3]:
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
    print 'ALIAS = %s ' % alias
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
    for i in [1,2,3]:
        USER = Addon.getSetting( 'user%i' % i )
        NOM =  Addon.getSetting( 'name%i' % i )
        SERVER = Addon.getSetting( 'server%i' % i )
        PASSWORD =  Addon.getSetting( 'pass%i' % i )
        PORT =  Addon.getSetting( 'port%i' % i )
        SSL = Addon.getSetting( 'ssl%i' % i ) == "true"
        TYPE = Addon.getSetting( 'type%i' % i )
        FOLDER = Addon.getSetting( 'folder%i' % i )
        #print "SERVER = %s, PORT = %s, USER = %s, password = %s, SSL = %s, TYPE = %s" % (SERVER,PORT,USER, PASSWORD, SSL,TYPE)
        #On cherche le serveur selectionne
        if (alias == NOM):
            self.SERVER = SERVER 
            self.USER = USER
            self.PORT = PORT
            self.PASSWORD = PASSWORD
            self.TYPE = TYPE
            self.SSL = SSL
            print "self.SERVER = %s " % self.SERVER
            dialog = xbmcgui.DialogProgress()
            dialog.create(Addon.getLocalizedString(id=614), Addon.getLocalizedString(id=610))  #Inbox,  Logging in...
            try:
                #Partie POP3
                if  '0' in self.TYPE:  #'POP' 
                    if self.SSL:
                        mail = poplib.POP3_SSL(str(self.SERVER),int(self.PORT))
                    else:  #'POP3'
                        mail = poplib.POP3(str(self.SERVER),int(self.PORT))
                    mail.user(str(self.USER))
                    mail.pass_(str(self.PASSWORD))
                    numEmails = mail.stat()[0]
#       #if '1' in self.TYPE: #IMAP
#       #   if self.SSL:
#	       imap = imaplib.IMAP4_SSL(self.SERVER, int(self.PORT))
#	  else:
#	       imap = imaplib.IMAP4(self.SERVER, int(self.PORT))
#	  imap.login(self.USER, self.PASSWORD)
#          id = 'folder' + str(i)
#	  FOLDER = Addon.getSetting( id )
#	  imap.select(FOLDER)
#	  #  numEmails = 1 
#	  numEmails = len(imap.search(None, 'UnSeen')[1][0].split()) 
#	  print "IMAP numEmails = %d " % numEmails
#
#	  self.getImapMails(imap)
#	  return

                print "You have", numEmails, "emails"
                #Affiche le nombre de msg
                self.getControl( NX_MAIL ).setLabel( '%d msg(s)' % numEmails )
                dialog.close()
                if numEmails == 0:
                    dialogOK = xbmcgui.Dialog()
                    dialogOK.ok("%s" % NOM ,Addon.getLocalizedString(id=612)) #no mail 
                    self.getControl( EMAIL_LIST ).reset()
                else:              #Inbox                           #You have                                           #emails
                    dialog.create(Addon.getLocalizedString(id=613),Addon.getLocalizedString(id=615) + str(numEmails) + Addon.getLocalizedString(id=616))
	                ##Retrieve list of mails
                    resp, items, octets = mail.list()
                    dialog.close()
                    #On recupere tous les messages pour les afficher
                    progressDialog = xbmcgui.DialogProgress()
                                          #Message(s)                       #Get mail
                    progressDialog.create(Addon.getLocalizedString(id=617), Addon.getLocalizedString(id=618))
                    i = 0
                    #Mise a zero de la ListBox msg
                    self.getControl( EMAIL_LIST ).reset()
                    self.emails = []
                    for item in items:
                        i = i + 1
                        #print "item %s" % item
                        id, size = string.split(item)
                        #progressDialog.update((id*100)/numEmails)
                        up = (i*100)/numEmails    #Get mail                         Please wait
                        progressDialog.update(up, Addon.getLocalizedString(id=618), Addon.getLocalizedString(id=619))

                        #Si dépasse la taille max on télécharge que 50 lignes
                        if (MAX_SIZE_MSG == 0) or (size < MAX_SIZE_MSG):
                            resp, text, octets = mail.retr(id)
                        else: 
                            resp, text, octets = mail.top(id,300)

                        print "size = %s " % size
		                #resp, text, octets = mail.top(id,300)
                        #resp, text, octets = mail.retr(id)
                        text = string.join(text, "\n")
                        myemail = email.message_from_string(text)
                        p = EmailParser()
                        msgobj = p.parsestr(text)
		                #print "res = % s, text = %s, size = %d" % (resp, text, octets)
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

                        attachments = []
                        body = None
                        html = None
                        for part in msgobj.walk():
                        #content_disposition = part.get("Content-Disposition", None)        
                        #print "content-disp =", content_disposition
                            if part.get_content_type() == "text/plain":
                                if body is None:
                                    body = ""
                                try :
                                    body += unicode(
                                        part.get_payload(decode=True),
                                        part.get_content_charset(),
                                        'replace'
                                        ).encode('utf8','replace')
                                except Exception ,e:
                                    print "UNICODE ERROR text/plain"
                                    print str(e)
        	                        #print "####> %s " % part.get_payload()
                                    body += "Erreur unicode"
                            elif part.get_content_type() == "text/html":
                                if html is None:
                                    html = ""
                                try :
                                    print "Charset = %s " % part.get_content_charset()
                                    #html += unicode(part.get_payload(decode=True),part.get_content_charset(),'replace').encode('utf8','replace')
                                    html += unicode(part.get_payload(decode=True),part.get_content_charset(),'replace').encode('ascii','replace')
                                    html = html2text(html)
			                        #print "HTML => %s" % html
                                except Exception ,e:
                                    print "UNICODE ERROR text/html"
                                    print str(e)
                                    body += "Erreur unicode html"
                            realname = parseaddr(msgobj.get('From'))[1]
                        Sujet = subject 
                        description = ' '
                        if (body):
                            description = str(body)
                        else:
                            html = html.encode('ascii','replace')
                            description = str(html)
		                #Nb de lignes du msg pour permettre le scroll text
                        self.nb_lignes = description.count("\n")
 
                        listitem = xbmcgui.ListItem( label2=realname, label=Sujet) 
                        listitem.setProperty( "realname", realname )    
                        listitem.setProperty( "date", date )   
                        listitem.setProperty( "message", description )
                        self.getControl( EMAIL_LIST ).addItem( listitem )
                    progressDialog.close()
                    #Affiche le 1er mail de la liste
                    self.getControl( EMAIL_LIST ).selectItem(0)

            except Exception, e:
                print "=============>"
                print str( e )
                dialog.close() #"Inbox"                         "Problem connecting to server : %s" 
                dialog.create(Addon.getLocalizedString(id=614),Addon.getLocalizedString(id=620) % self.SERVER)
                time.sleep(5)
                dialog.close()
  
  def getImapMails(self, imap):
    print "getImapMails"
    #Mise a zero de la ListBox msg
    self.getControl( EMAIL_LIST ).reset()
    self.emails = []
 

  def onAction(self, action):
        #print "ID Action %d" % action.getId()
        #print "Code Action %d" % action.getButtonCode()
        if action == ACTION_PREVIOUS_MENU:
            self.close()
        if action == ACTION_MOVE_UP:
            controlId = action.getId()
            print "ACTION _MOVE_UP"
        if action == ACTION_MOVE_DOWN:
            controlId = action.getButtonCode()
            print "ACTION _MOVE_DOWN"
        if action == ACTION_FASTFORWARD: #PageUp
            if (self.position > 0):
                self.position = self.position - 1
            self.getControl( MSG_BODY ).scroll(self.position)
        if (action == ACTION_REWIND): #PageUp
            if (self.position <= self.nb_lignes):
                self.position = self.position + 1
            self.getControl( MSG_BODY ).scroll(self.position)
																       
  def onClick( self, controlId ):
        print "onClick controId = %d " % controlId
        if (controlId in [SERVER1,SERVER2,SERVER3]):
            print "onClick controId bis = %d " % controlId
            label = self.getControl( controlId ).getLabel()
            self.checkEmail(label)
        elif (controlId == QUIT):
            self.close()

mydisplay = MailWindow( "WinCourrier.xml" , __cwd__, "Default")
mydisplay .doModal()
del mydisplay
