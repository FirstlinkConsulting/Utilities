# Email Module Interface validation
# This module is a validation of the Email Module Interface
# It uses the unittest module to provide the test infrastructure.

from ctypes import *
from progressbar import ProgressBar, SimpleProgress
from logging.config import dictConfig
import unittest
import getopt
import configparser
import time
import os
import platform
import re
import sys
import inspect
import binascii
import logging
import logging.config
import logging.handlers
import Timer
import shutil
import subprocess
import datetime



class SetUp(object):
    _instance = None

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super(SetUp, cls).__new__(cls, *args, **kwargs)
        return cls._instance

    def __init__(self, *args, **kwargs):
        global total_numbers_of_test_cases
        global current_progress

        self.logger = logging.getLogger(__name__)
        current_progress += 1
        self.logger.debug('init singleton')

        if total_numbers_of_test_cases > 0:
            self.logger.debug("\r\n\r\n****** Test Cases Progress => %s of %s ******" %
                (current_progress - 1, total_numbers_of_test_cases))
        return

    def __del__(self):
        pass


class EmailValidationTestCase(unittest.TestCase):

    def setUp(self):
        self.s = SetUp()
        return


    def tearDown(self):
        #print("tearDown()")
        pass

    @classmethod
    def setUpClass(cls):
        #print("setUpClass()")
        from Email import Email
        cls.email = Email(accountEmailAddr = "mbgpqa@gmail.com", passWord = "mbgpqa1981")

    @classmethod
    def tearDownClass(cls):
        #print("tearDownClass()")
        pass

    """ A Test Verifying function 'RegisterTo' """
    def test_RegisterTo(self):
        global pbar
        global test_index

        test_index += 1
        pbar.update(test_index)


        self.s.logger.info("test_RegisterTo")
        userInputMsg = 'jsainfeld@logitech.com'

        result = self.email.RegisterTo(userInputMsg)
        #self.logger.debug("return = %s" % result)
        self.assertEqual(result, True)
        return

    """ A Test Verifying function 'RegisterCc' """
    def test_RegisterCc(self):
        global pbar
        global test_index

        test_index += 1
        pbar.update(test_index)
        #self.logger.debug("test_RegisterCc")

        userInputMsg = 'ephraim.schoenfeld@shamirgems.com'

        result = self.email.RegisterCc(userInputMsg)
        #self.logger.debug("return = %s" % result)
        self.assertEqual(result, True)
        return

    """ A Test Verifying function 'RegisterBc' """
    def test_RegisterBc(self):
        global pbar
        global test_index

        test_index += 1
        pbar.update(test_index)

        #self.logger.debug("test_RegisterBc")

        userInputMsg = 'mbgpqa@gmail.com'

        result = self.email.RegisterBc(userInputMsg)
        #self.logger.debug("return = %s" % result)
        self.assertEqual(result, True)
        return

    """ A Test Verifying function 'MailSend' """
    def test_MailSend(self):
        global pbar
        global test_index
        test_index += 1
        pbar.update(test_index)
        attachements = []
        #self.logger.debug("test_MailSend")
        msg = 'this is a message to be sent as email.'
        msg += 'We can read this message from some file'
        attachements.append('attachement.txt')
        attachements.append('attachement1.txt')
        subject = 'This is the new subject not the default one'
        for i in range(2):
            self.email.MailSend(self.email.toAddrLst,
                                msg,
                                attachements,
                                self.email.ccAddrLst,
                                self.email.bccAddrLst,
                                subject
                                )
        return


""" Email Test Suite """
def EmailValidationTestSuite(fn):
    global pbar
    global test_index

    logger = logging.getLogger(__name__)
    nb_test_steps = 0
    suite = unittest.TestSuite()
    if fn != None:
        logger.debug('Process the specified test suite from the file %s' %fn)
        if os.path.exists(fn):
            lines = [line.rstrip('\n') for line in open(fn)]
            for i in range(len(lines)):
                l = (lines[i])
                if l != '':
                    logger.debug(l)
                    try:
                        testcase, teststep, state, repeatCount = l.split(',', 4)
                        if state == 'enable':
                            logger.debug('teststep %s is enabled' % teststep)
                            for j in range(int(repeatCount)):
                                logger.debug('add this test to the suite')
                                x = 'suite.addTest(' + testcase + '(\''+teststep +'\'))'
                                nb_test_steps += 1
                                logger.debug(x)
                                exec(x)
                    except:
                        logger.debug('line %s did not split correctly' % l)
                        continue
    else:

        suite = unittest.TestSuite()
        suite.addTest(EmailValidationTestCase("test_RegisterTo"))
        nb_test_steps += 1
        suite.addTest(EmailValidationTestCase("test_RegisterCc"))
        nb_test_steps += 1
        suite.addTest(EmailValidationTestCase("test_RegisterBc"))
        nb_test_steps += 1
        suite.addTest(EmailValidationTestCase("test_MailSend"))
        nb_test_steps += 1

    pbar = ProgressBar(widgets=[SimpleProgress()], maxval=nb_test_steps).start()
    test_index = 0
    return suite
'''
Displays Usage in case of help or command line error
'''
def usage():
    logger = logging.getLogger(__name__)
    logger.info('This is the usage function ')
    logger.info('Usage: ' + sys.argv[0] + ' [OPTIONS]')
    logger.info('Mandatory arguments to long options are mandatory for short options too.')
    logger.info('')
    logger.info('\t\t -h,  --help')
    logger.info('\t\t\t\t display the various options whith their usage')
    logger.info('\t\t\t\t if this option is specified with others it will take precendence')
    logger.info('\t\t -c,  --config=<PLATFORM_DESCRIPTION>')
    logger.info('\t\t\t\t specifiy the path of the platform description file')
    logger.info('\t\t\t\t if this option is specified the file is a mandatory parameter')
    logger.info('\t\t -o,  --output=<PATH_OF_OUTPUT_FILE>')
    logger.info('\t\t\t\t specifiy the path of the output file')
    logger.info('\t\t\t\t if this option is specified the file is a mandatory parameter')
    logger.info('\t\t -v,  --verbose')
    logger.info('\t\t\t\t indicator of verbosity in the output')
    logger.info('\t\t -i,  --interactive')
    logger.info('\t\t\t\t indicator of interactive mode')
    logger.info('\t\t -r,  --report=<PATH_OF_REPORT_FILE>')
    logger.info('\t\t\t\t specifiy the path of the report file')
    logger.info('\t\t\t\t if this option is specified the file is a mandatory parameter')
    logger.info('\t\t -m,  --mode=<MODE>')
    logger.info('\t\t\t\t the specified mode may be either HTML|TEXT')
    logger.info('\t\t\t\t if this option is specified the file is a mandatory parameter')
    logger.info('\t\t -x, --repeat=<REPEAT_COUNT')
    logger.info('\t\t\t\t Test suite excution will be repeated the specified number of times')
    logger.info('\t\t\t\t the specified mode may be either HTML|TEXT')
    logger.info('\t\t -s, --suite=<TEST_SUITE_FILENAME>')
    logger.info('\t\t\t\t the specified test suite will be executed instead of the inline one')
    logger.info('\t\t\t\t if this option is specified the file is a mandatory parameter')
    logger.info('\t\t -n, --notification=<NOTIFICATION_FILENAME>')
    logger.info('\t\t\t\t the person(s) specified in the notification  will be sent message of completion')
    logger.info('\t\t\t\t if this option is specified the file is a mandatory parameter')
    logger.info('End of Usage')
    return

def main(argv):
    global current_progress
    global total_numbers_of_test_cases

    total_numbers_of_test_cases = 0
    current_progress = 0
    lcfgPath = os.path.abspath(os.path.dirname(sys.argv[0])) + '\logging.cfg'
    print(lcfgPath)
    logging.config.fileConfig(lcfgPath, defaults=None, disable_existing_loggers=True)
    logger = logging.getLogger(__name__)
    logger.debug('Email Validation Log')

    # When this module is executed from the command line,  run all its tests
    try:
        opts, args = getopt.getopt(sys.argv[1:], 'hc:v:m:x:s:n:',
								   ['help', 'config=', 'verbose=', 'mode=',
								    'repeat=', 'suite=', 'notification=' ])
    except getopt.GetoptError as err:
        # print help information and exit:
        logger.error (str(err)) # will print something like 'option -a not recognized'
        logger.error('option not recognized')
        usage()
        sys.exit(2)


    config = None
    mode = None
    verbose_level = 'NOTSET'
    repeat = 1
    suite_fn = None
    config_fn = None
    notification_fn = None

    for o, a in opts:
        if o in ( '-v', '--verbose'):
            valid_levels = [ 'NOTSET',
                             'DEBUG',
                             'INFO',
                             'WARNING',
                             'ERROR',
                             'CRITICAL'
                             ]

            logger_levels = [logging.NOTSET,
                             logging.DEBUG,
                             logging.INFO,
                             logging.WARNING,
                             logging.ERROR,
                             logging.CRITICAL
                            ]
            verbose_level = a
            if verbose_level in valid_levels:
                i = valid_levels.index(verbose_level)
                logger.setLevel(logger_levels[i])
            else:
                usage()
                sys.exit()
        elif o in ('-c', '--config'):
            config_fn = a
            if os.path.exists(config_fn) is False:
                logger.error('The specified config file %s does not exist' % config_fn)
                sys.exit(3)
        elif o in ('-s', '--suite'):
            suite_fn = a
            if os.path.exists(suite_fn) is False:
                logger.error('The specified test suite file %s does not exist' % suite_fn)
                sys.exit(3)
        elif o in ('-n', '--notification'):
            notification_fn = a
            if os.path.exists(notification_fn) is False:
                logger.error('The specified test suite file %s does not exist' % notification_fn)
                sys.exit(3)
        elif o in ('-m', '--mode'):
            mode = a
        elif o in ('-x', '--repeat'):
            repeat = int(a)


    #
    # handle notfication option if specified on command line
    #
    if notification_fn:
        logger.debug('Handle notification file')
        from Email import Email
        notif = configparser.ConfigParser()
        try:
            notif.read(notification_fn)
        except:
            logger.exception('We failed reading the notification file ')
        logger.debug(notif)
        logger.debug(notif.sections())
        for n in notif.sections():
            logger.debug(n)

        account = notif['DEFAULT']['accountEmailAddr']
        password = notif['DEFAULT']['password']
        email = Email(accountEmailAddr = account, passWord = password)


        to_register_list = notif['TORECIPIENT']['register_list']
        tl = to_register_list.split(',')
        for to in tl:
            try:
                email.RegisterTo(to)
            except:
                pass

        to_unregister_list = notif['TORECIPIENT']['unregister_list']
        tl = to_unregister_list.split(',')
        for to in tl:
            try:
                email.UnregisterTo(to)
            except:
                pass

        cc_register_list = notif['CCRECIPIENT']['register_list']
        cl = cc_register_list.split(',')
        for cc in cl:
            try:
                email.RegisterCc(cc)
            except:
                pass

        cc_unregister_list = notif['CCRECIPIENT']['unregister_list']
        cl = cc_unregister_list.split(',')
        for cc in cl:
            try:
                email.UnregisterCc(cc)
            except:
                pass

        bcc_register_list = notif['BCCRECIPIENT']['register_list']
        bcl = bcc_register_list.split(',')
        for bcc in bcl:
            try:
                email.RegisterBc(bcc)
            except:
                pass

        bcc_unregister_list = notif['BCCRECIPIENT']['unregister_list']
        bcl = bcc_unregister_list.split(',')
        for bcc in bcl:
            try:
                email.UnregisterBc(bcc)
            except:
                pass

    test_suite_outcome = 'FAIL'
     # build test suite and count all of test cases
    test_suite = EmailValidationTestSuite(suite_fn)

    total_numbers_of_test_cases = test_suite.countTestCases() * repeat

    if mode == 'HTML':

        logger.debug('Use the TestReportGen module to generate HTML reports')
        import TestRunner
        with open('result.html', 'w', encoding='utf-8') as f:
            runner = TestRunner.TestRunner( stream=f, verbosity = 19,  title='Email Validation ', description='Email Validation HTML report')
            for i in range(repeat):
                logger.info('TestSuite [ %s ] run # %i of %i' % (suite_fn, i, repeat))
                result = runner.run(EmailValidationTestSuite(suite_fn))
                if result.wasSuccessful() is True:
                    test_suite_outcome='PASS'
                if result.wasSuccessful() is False:
                    test_suite_outcome='FAIL'

            f.close()

    else:
        # do report in text mode with the eventual report file name option

        logger.debug('Use the text Report ')
        with open('result.txt', 'w', encoding='utf-8') as f:
            f.write('Running Email Validation test\n')
            f.write('=========================================\n')
            runner = unittest.TextTestRunner(stream=f, verbosity=19)
            for i in range(repeat):
                logger.info('TestSuite [ %s ] run # %i of %i' % (suite_fn, i, repeat))
                result = runner.run(EmailValidationTestSuite(suite_fn))
                if result.wasSuccessful() is True:
                    test_suite_outcome='PASS'
                if result.wasSuccessful() is False:
                    test_suite_outcome='FAIL'

            f.close()




       ##  Handle Notification tasks
    #
    if notification_fn:
        attachements = []
        logger.debug("Sending Notification mail to registered Users")
        msg = 'This notification is to inform you that:\n\n'
        msg += 'The Email Validation '
        msg += ' is now completed.'
        msg += '\n\n'
        msg += 'The attached report file: '
        msg += '#### Not yet supplied ####'
        msg += ' can be consulted for details of the outcome.'
        msg += '\n\n'
        msg += 'We ran: '
        msg += str(total_numbers_of_test_cases)
        msg += ' test cases.'

        attachements.append('notification.txt')
        subject = 'Test Completion Notification for '
        subject += 'Email'
        subject += ' ( '
        subject += test_suite_outcome
        subject += ' ) '

        email.MailSend(email.toAddrLst,
                            msg,
                            attachements,
                            email.ccAddrLst,
                            email.bccAddrLst,
                            subject
                            )


if __name__ == '__main__':
        main(sys.argv[1:])
