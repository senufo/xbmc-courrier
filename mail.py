#Script pour consulter ses mails
#Senufo, 2011 (c)
#Version 0.0.3
#
# $Date: 2011-11-13 22:23:30 +0100 (dim. 13 nov. 2011) $
# $Author: Senufo $
#
import os, re
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
#Traitmennt des fichiers xml
import xml.dom.minidom

#Modules xbmc
import xbmc, xbmcgui
import xbmcaddon

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

class MailWindow(xbmcgui.WindowXML):
   
  def __init__(self, *args, **kwargs):

    #Variablew du systeme
    MAX_SIZE_MSG = 51200
    SMART_HTML = True
    #ouvre le fichier des comptes mail
    dirHome = __cwd__
    #Creation des boutons et fenetres
    #variable pour position dans le msg
    self.position = 0
    
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
    for i in [1,2,3]:
     id = 'user' + str(i)
     USER = Addon.getSetting( id )
     id = 'name' + str(i)
     NOM =  Addon.getSetting( id )
     id = 'server' + str(i)
     SERVER = Addon.getSetting( id )
     id = 'pass' + str(i)
     PASSWORD =  Addon.getSetting( id )
     id = 'port' + str(i)
     PORT =  Addon.getSetting( id )
     id = 'ssl' + str(i)
     SSL = Addon.getSetting( id ) == "true"
     id = 'type' + str(i)
     TYPE = Addon.getSetting( id )
     print "SERVER = %s, PORT = %s, USER = %s, password = %s, SSL = %s, TYPE = %s" % (SERVER,PORT,USER, PASSWORD, SSL,TYPE)
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
    dialog.create("Inbox","Logging in...")
    try:
       #Partie POP3
       print "183 = %s, %s ,%s, %s " % (self.SERVER,self.PORT , self.USER, self.TYPE)
       if  '0' in self.TYPE:  #'POP' 
	  if self.SSL:
	       mail = poplib.POP3_SSL(str(self.SERVER),int(self.PORT))
	  else:  #'POP3'
	       print "190"
	       mail = poplib.POP3(str(self.SERVER),int(self.PORT))
	  print "191"
          mail.user(str(self.USER))
          mail.pass_(str(self.PASSWORD))
          numEmails = mail.stat()[0]
	  print "194"
       if '1' in self.TYPE: #IMAP
          if self.SSL:
	       imap = imaplib.IMAP4_SSL(self.SERVER, int(self.PORT))
	  else:
	       imap = imaplib.IMAP4(self.SERVER, int(self.PORT))
	  imap.login(self.USER, self.PASSWORD)
          id = 'folder' + str(i)
	  FOLDER = Addon.getSetting( id )
	  imap.select(FOLDER)
	  #  numEmails = 1 
	  numEmails = len(imap.search(None, 'UnSeen')[1][0].split()) 
	  print "IMAP numEmails = %d " % numEmails
	  return

       print "You have", numEmails, "emails"
       #Affiche le nombre de msg
       self.getControl( NX_MAIL ).setLabel( '%d msg(s)' % numEmails )
       dialog.close()
       if numEmails == 0:
	    dialog.create("Courriels","Pas de courriels")
            time.sleep(5)
	    dialog.close()
       else:
            dialog.create("Inbox","You have " + str(numEmails) + " emails")
            ##Retrieve list of mails
            resp, items, octets = mail.list()
            print "resp = % s" % resp
            print "items ", items
            dialog.close()
            #result = resp.find('+OK')
            #On recupere tous les messages pour les afficher
            progressDialog = xbmcgui.DialogProgress()
            progressDialog.create('Message(s)', 'Get mail')
            i = 0
            #Mise a zero de la ListBox msg
            self.getControl( EMAIL_LIST ).reset()
            self.emails = []
            for item in items:
                i = i + 1
                #print "item %s" % item
	        id, size = string.split(item)
		#progressDialog.update((id*100)/numEmails)
                up = (i*100)/numEmails
                progressDialog.update(up, 'Get mail', 'Please wait')
		#resp, text, octets = mail.top(id,300)
		resp, text, octets = mail.retr(id)
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
			print "Date = %s" %  msgobj['Date']
		else:
			date = '--'
			print "Pas de date"
                #print "Sujet = %s " % subject
                Sujet = subject
                realname = parseaddr(msgobj.get('From'))[1]
                #print "FROM = ", realname
                listitem = xbmcgui.ListItem( label2=realname, label=Sujet) 
                listitem.setProperty( "summary", realname )    
                listitem.setProperty( "updated", date )    
		self.getControl( EMAIL_LIST ).addItem( listitem )
            progressDialog.close()
            #Affiche le 1er mail de la liste
            self.processEmail(0)
            self.getControl( EMAIL_LIST ).selectItem(0)

    except:
        dialog.close()
	dialog.create("Inbox","Problem connecting to server : %s" % self.SERVER)
        time.sleep(5)
        dialog.close()
        #self.setFocus(self.cmButton)
     #else : print "Nom = %s "% NOM

  def onInit( self ):
    #self.getControl( STATUS_LABEL ).setLabel( 'Serveur LIBERTYSURF ...' )
    self.getControl( EMAIL_LIST ).reset()
    for i in [1,2,3]:
	id = 'name' + str(i)
    	NOM =  Addon.getSetting( id )
	Button_Name = 1000 + i 
	self.getControl( Button_Name ).setLabel( NOM )
    self.checkEmail(Addon.getSetting( 'name1' ))

#Recupere le texte dans un tag xml
  def getText(self, nodelist):
    rc = ""
    for node in nodelist:
       if node is not None:
         if node.nodeType == node.TEXT_NODE:
            rc = rc + node.data
    return rc


  def onAction(self, action):
    print "ID Action %d" % action.getId()
   #print "Code Action %d" % action.getButtonCode()
    if action == ACTION_PREVIOUS_MENU:
       self.close()
    if action == ACTION_MOVE_UP:
       controlId = action.getId()
       print "ACTION _MOVE_UP"
       #if (controlId == EMAIL_LIST ):
       self.processEmail(self.getControl( EMAIL_LIST ).getSelectedPosition())
    if action == ACTION_MOVE_DOWN:
       controlId = action.getButtonCode()
       print "ACTION _MOVE_DOWN"
       #if (controlId == EMAIL_LIST ):
       self.processEmail(self.getControl( EMAIL_LIST ).getSelectedPosition())
    if action == ACTION_FASTFORWARD: #PageUp
       if (self.position > 0):
             self.position = self.position - 1
       self.getControl( MSG_BODY ).scroll(self.position)
    if (action == ACTION_REWIND): #PageUp
       if (self.position <= self.nb_lignes):
             self.position = self.position + 1
       self.getControl( MSG_BODY ).scroll(self.position)
																       
#  def onControl(self, control):
    #print "onCOntrol %s " % control
    #Identifie le bouton selectione
    #for bouton in self.btnevent:    
#	label = bouton.getLabel()
#	print "Label = %s " % label
#	if control == bouton: 
#		print "%s select" % label
#		self.checkEmail(label)
 #   if control == self.listControl:
  #      self.processEmail(self.listControl.getSelectedPosition())
   # elif control == self.msgbody:
#	    print "MSGBODY"
 #   if control == self.vaButton:
  #      self.close()

  #def onFocus(self, control):
    #print "onFocus"
    #if control == self.cmButton:
    #    self.checkEmail()
    #elif control == self.listControl:
        #self.fsButton.setLabel("Fullscreen", "font14", "ffffffff")
        #self.cmButton.controlDown(self.fsButton)
    #    self.processEmail(self.listControl.getSelectedPosition())
    #if control == self.vaButton:
    #    self.close()
   
  def onClick( self, controlId ):
    print "onClick controId = %d " % controlId
    if (controlId == EMAIL_LIST ):
	#dialog = xbmcgui.Dialog()
	#ok = dialog.ok('Sujet', 'Here the messages')
        self.processEmail(self.getControl( controlId ).getSelectedPosition())
    elif (controlId in [SERVER1,SERVER2,SERVER3]):
	label = self.getControl( controlId ).getLabel()
        self.checkEmail(label)
    elif (controlId == QUIT):
    	self.close()

  def processEmail(self, position):
    self.position = 0	
    #Efface le texte deja present 
    #self.msgbody.reset() 
    print "Position = %d " % position
    position = position + 1
    
    mail = poplib.POP3(str(self.SERVER))
    mail.user(str(self.USER))
    mail.pass_(str(self.PASSWORD))
       
    #resp, text, octets = mail.top(position,500)
    resp, text, octets = mail.retr(position)
    text = string.join(text, "\n")
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
    print "Sujet = %s " % subject
    if msgobj['Date'] is not None:
	date = msgobj['Date']
	print "Date = %s" %  msgobj['Date']
    else:
	date = '--'
	print "Pas de date"

    attachments = []
    body = None
    html = None
    for part in msgobj.walk():
     #content_disposition = part.get("Content-Disposition", None)        
     #print "content-disp =", content_disposition
     if part.get_content_type() == "text/plain":
           if body is None:
               body = ""
           body += unicode(
                part.get_payload(decode=True),
                part.get_content_charset(),
                'replace'
                ).encode('utf8','replace')
     elif part.get_content_type() == "text/html":
           if html is None:
               html = ""
           html += unicode(part.get_payload(decode=True),part.get_content_charset(),'replace').encode('utf8','replace')
     
     realname = parseaddr(msgobj.get('From'))[1]
     print "FROM = %s " % realname
    Sujet = subject 
    #print "Sujet = ", Sujet
    description = 'De :' + realname + '\nSujet :' + Sujet + "\nDate :" + date
    description = description + '\n__________________________________________________________________\n\n'
    description = ' '
    if (body):
       description = description + str(body)
    else:
       description = description + str(html)
    description = description + str(body)
    self.nb_lignes = description.count("\n")
    self.getControl( MSG_BODY ).setVisible(True)
    self.getControl( MSG_BODY ).setText(description)



mydisplay = MailWindow( "myWin.xml" , __cwd__, "Default")
mydisplay .doModal()
del mydisplay

        
