#!/usr/bin/env python

## @mainpage Email Notification Documentation
#
#  @section intro_sec Introduction
#  This documentation is created to give usage details of Email Notification.
#  The Email Notification module provides an interface for sending a notifiaction
#  via email to tester when test is finished.
#  \n Somtimes the time of testing will be longer,
#  so tester will wastes much time to wait for the result from testing application.
#  If there is notification can be sent to the tester automatically when test is finished,
#  it will be convevient.
#  \n This module send a notification by email.
#  This email can not only be sent to the tester but also CC or BCC be sent to other people.
#  The files of test result report in detail can be also attahced in this email.
#
#  @section reference_sec Reference
#  More detailed inforamtion about SMTP and email can be found in
#  the native Python module 'smtplib' and 'email'.
#
#  @section design_sec Design
#  This module is wrapped as a class with Python module 'smtplib' and 'email'.
#  User can use a series of simple APIs to send notification
#  without realizing the structure and operations of SMTP and email in detail.
#
#  @section install_sec Installation
#   The non-native Python module 'validate_email' must be installed before using this module.
#   This module 'validate_email' can be found in GitHUB and be installed by a file 'setup.py'.
#   \n To facilitate and ensure this compatibility the installation task is entrusted to the
#   Python distribution facility.
#   A setup.py file of module 'Email' is used to invoke the module installation mechanism
#   in the directory of this module run the following command:
#   @em "python setup.py install" @em
#
#  @section license_sec License
#  Copyright 2015 @em Logitech @em



import os
import smtplib
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from mimetypes import guess_type
from email.encoders import encode_base64
from getpass import getpass
from validate_email import validate_email



## A email object  allowing the sending of emails
class Email(object):
    """Email Notification Package"""
    __COMMASPACE = ', '
    toAddrLst = []
    ccAddrLst = []
    bccAddrLst = []



    ##Constructor
    #\verbatim
    #   Email Email(
    #       smptSvrIp = "smtp.gmail.com",
    #       port = 587,
    #       accountEmailAddr = "mbgpqa@gmail.com",
    #       passWord = "mbgpqa1981"
    #   )
    #\endverbatim
    # @brief New a instance of class 'Email''.
    # @param smtpSvrIp: DNS Name of the smtp server
    # @param port: Port at which the smtp server listen
    # @param accountEmailAddr: Sender email address.
    #        \n Sender may be same as receiver.
    #        Some SMTPs server will check the current location of user
    #        according to the current IP address.
    #        If user accesses the server between difference locations
    #        in a long diatance and a short time,
    #        server will reject the access for security of account.
    # @pram passWord: password for the server email
    # @return instance of calss 'Email'.
    def __init__(self, smtpSvrIp = "smtp.gmail.com", port = 587, accountEmailAddr = "mbgpqa@gmail.com", passWord = "mbgpqa1981"):
        try:
            self.__sender = accountEmailAddr
            self.smtpSvrIp = smtpSvrIp
            self.port = port
            self.passWord = passWord

        except:
            print("Email Notification Module Initialization Error: SMTPException: ")


        self.toAddrLst.clear()
        self.ccAddrLst.clear()
        self.bccAddrLst.clear()
        self.logger = logging.getLogger(__name__)

        return



    ## Destructor
    #
    def __del__(self):
        self.__smtpObj.close()
        return


    ## Connect
    #
    def Connect(self):
        self.__smtpObj = smtplib.SMTP(self.smtpSvrIp, self.port)
        self.__smtpObj.starttls()
        self.__smtpObj.login(self.__sender, self.passWord)
        self.__smtpObj.set_debuglevel(1)
        return



    ## Disconnect
    #
    def Disconnect(self):
        self.__smtpObj.quit()
        return


    ##RegisterTo
    #\verbatim
    #   RegisterTo(toAddress)
    #\endverbatim
    # @brief Register "To" address to be notified
    # @param toAddress: wellformed email address to be added to the TO list
    # @return True = Success
    #       \n False = Fail
    def RegisterTo(self, toAddress):
        if validate_email(toAddress) == True:
            self.toAddrLst.append(toAddress)
            return True
        else:
            self.logger.error("email address is not valid")
            return False



    ##RegisterCc
    #\verbatim
    #   RegisterCc(ccAddress)
    #\endverbatim
    # @brief Register "Cc" address to be notified
    # @param ccAddress: wellformed email address to be added to the CC list
    # @return True = Success
    #       \n False = Fail
    def RegisterCc(self, ccAddress):
        if validate_email(ccAddress) == True:
            self.ccAddrLst.append(ccAddress)
            return True
        else:
            self.logger.error("email address is not valid")
            return False



    ##RegisterBc
    #\verbatim
    #   RegisterBc(bccAddress)
    #\endverbatim
    # @brief Register "Bcc" address to be notified
    # @param bccAddress: wellformed email address to be added to the BCC list
    # @return True = Success
    #       \n False = Fail
    def RegisterBc(self, bccAddress):
        if validate_email(bccAddress) == True:
            self.bccAddrLst.append(bccAddress)
            return True
        else:
            self.logger.error("email address is not valid")
            return False



    ##UnregisterTo
    #\verbatim
    #   UnregisterTo(toAddress)
    #\endverbatim
    # @brief Unregister "To" address to be notified
    # @param toAddress: wellformed email address to be removed from the TO list
    # @return True = Success
    #       \n False = Fail
    def UnregisterTo(self, toAddress):
        try:
            idx = self.toAddrLst.index(toAddress)
        except:
            self.logger.error("item %s is not in the list %s" % (toAddress, self.toAddrLst))
            return False
        self.toAddrLst.remove(toAddress)
        return True



    ##UnregisterCc
    #\verbatim
    #   UnregisterCc(ccAddress)
    #\endverbatim
    # @brief Unregister "CC" address to be notified
    # @param ccAddress: wellformed email address to be removed from the CC list
    # @return True = Success
    #       \n False = Fail
    def UnregisterCc(self, ccAddress):
        try:
            idx = self.ccAddrLst.index(ccAddress)
        except:
            self.logger.error("item %s is not in the list %s" % (ccAddress, self.ccAddrLst))
            return False
        self.ccAddrLst.remove(ccAddress)
        return True



    ##UnregisterBc
    #\verbatim
    #   UnregisterBc(bccAddress)
    #\endverbatim
    # @brief Unregister "BCC" address to be notified
    # @param bccAddress: wellformed email address to be removed from the BCC list
    # @return True = Success
    #       \n False = Fail
    def UnregisterBc(self, bccAddress):
        try:
            idx = self.bccAddrLst.index(bccAddress)
        except:
            self.logger.error("item %s is not in the list %s" % (bccAddress, self.bccAddrLst))
            return False
        self.bccAddrLst.remove(bccAddress)
        return True



    ##MailSend
    #\verbatim
    #   MailSend(
    #       toAddrLst,
    #       message,
    #       attachedFiles = None,
    #       ccAddrLst = None,
    #       bccAddrLst = None,
    #       subject = "Test Automation Notification"
    #   )
    #\endverbatim
    # @brief Send Mail
    # @param toAddressLst: a list containing the intended recipients of the email
    # @param message: a text to be included in the email as the body thereof
    # @param attachedFiles: A list of filenames to be attached to this email
    # @param ccAddrLst: a list of  emails addresses to be copied on this email
    # @param bccAddrLst: a list of emails addresses to be blind copied on this email
    # @param subject: the purpose of this email
    def MailSend(
        self, toAddressLst, message,
        attachedFiles = None, ccAddrLst = None, bccAddrLst = None, subject = "Test Automation Notification"):

        # Open a session with the server
        self.Connect()

        finalToAddrLst = []
        finalToAddrLst += toAddressLst

        mail = MIMEMultipart()
        mail['From'] = self.__sender
        mail['To'] = Email.__COMMASPACE.join(toAddressLst)
        mail['Subject'] = subject

        if ccAddrLst is not None:
            finalToAddrLst += ccAddrLst
            mail['Cc'] = Email.__COMMASPACE.join(ccAddrLst)

        if bccAddrLst is not None:
            finalToAddrLst += bccAddrLst
            mail['Bcc'] = Email.__COMMASPACE.join(bccAddrLst)

        text = MIMEText(message, 'plain', 'us-ascii')
        mail.attach(text)


        if attachedFiles is not None:
            for filename in attachedFiles:
                mimetype, encoding = guess_type(filename)
                if mimetype is None:
                    mimetype = 'application/octet-stream'

                mimetype = mimetype.split('/', 1)
                with open(filename, 'rb') as fp:
                    attachment = MIMEBase(mimetype[0], mimetype[1])
                    attachment.set_payload(fp.read())
                    fp.close()
                encode_base64(attachment)
                attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(filename))
                mail.attach(attachment)


        # SMTPHeloError          The server didn't reply properly to
        #                       the helo greeting.
        # SMTPRecipientsRefused  The server rejected ALL recipients
        #                        (no mail was sent).
        # SMTPSenderRefused      The server didn't accept the from_addr.
        # SMTPDataError          The server replied with an unexpected
        #                       error code (other than a refusal of
        #                       a recipient).
        try:
            self.__smtpObj.sendmail(self.__sender, finalToAddrLst, mail.as_string())
        except:
            self.logger.error("Error: unable to send email, SMTPException: ")

        self.Disconnect()

        return

