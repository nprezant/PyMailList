#-*- coding: utf-8 -*-

import sys
import time
import traceback
from pathlib import Path
from collections import OrderedDict

#import subprocess
#subprocess.call(['python', '-m', 'PyQt5.uic.pyuic', '-x', 'design\\mainwindow.ui', '-o', 'design\\mainwindow.py'])
# python -m PyQt5.uic.pyuic -x mainwindow.ui -o mainwindow.py

from design.mainwindow import Ui_MainWindow
from gmail import Authenicator, Emailer

from PyQt5.QtCore import pyqtSignal, QObject, QRunnable, pyqtSlot, QThreadPool, pyqtSlot, QFile, QTextStream
from PyQt5.QtWidgets import QApplication, QMainWindow, QDialog, QMessageBox, QTextEdit
from PyQt5.QtGui import QIcon, QPixmap

class WorkerSignals(QObject):
    '''
    Defines the signals available from a running worker thread.

    Supported signals are:

    finished
        No data

    error
        `tuple` (exctype, value, traceback.format_exc() )

    result
        `object` data returned from processing, anything

    progress
        `int` indicating % progress
    '''
    finished = pyqtSignal()
    error = pyqtSignal(tuple)
    result = pyqtSignal(object)
    progress = pyqtSignal('PyQt_PyObject')


class Worker(QRunnable):
    '''
    Worker thread

    Inherits from QRunnable to handler worker thread setup, signals and wrap-up.

    :param callback: The function callback to run on this worker thread. Supplied args and 
                     kwargs will be passed through to the runner.
    :type callback: function
    :param args: Arguments to pass to the callback function
    :param kwargs: Keywords to pass to the callback function
    '''

    def __init__(self, fn, *args, **kwargs):
        super().__init__()

        # Store constructor arguments (re-used for processing)
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()

        #pythoncom.CoInitialize()

    @pyqtSlot()
    def run(self):
        '''
        Initialise the runner function with passed args, kwargs.
        '''

        # Retrieve args/kwargs here; and fire processing using them
        try:
            result = self.fn(*self.args, **self.kwargs)
        except:
            traceback.print_exc()
            exctype, value = sys.exc_info()[:2]
            self.signals.error.emit((exctype, value, traceback.format_exc()))
        else:
            self.signals.result.emit(result)
        finally:
            self.signals.finished.emit()


class ApplicationWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self._setupUi_extra()

        self.auth = Authenicator('credentials.json', 'token.json')
        self.email = Emailer(service=None)

        # note that we are not yet authorized, nor are we sending anything
        self._sender_thread_is_running = False
        self._auth_thread_is_running = False
        self._authorized = False

        # init thread pool
        self.threadpool = QThreadPool()
        self.console_log('{} threads available or multi-threading'.format(self.threadpool.maxThreadCount()))

        # try to authorize
        self.start_authorize_thread()

        # connect actions
        self.ui.action_send.triggered.connect(self.send)
        self.ui.action_authorize.triggered.connect(self.force_authorize)
        self.ui.action_clear_fields.triggered.connect(self.clear)
        self.ui.action_toggle_theme.triggered.connect(self.cycle_stylesheet)

        # connect buttons
        self.ui.clear_console_button.clicked.connect(self.clear_console)
        self.ui.toggle_console_wrap_button.clicked.connect(lambda: self.toggle_word_wrap(self.ui.console))
        self.ui.reset_progress_bar_button.clicked.connect(self.reset_progress_bar)


    def _setupUi_extra(self):
        '''
        does the extra Ui setup stuff like icons because
        I am dumb and can't figure out the resource browser
        '''
        icon = QIcon()
        icon.addPixmap(QPixmap('design/icons/running_stick.png'), QIcon.Normal, QIcon.Off)
        self.setWindowIcon(icon)

        self.themes = OrderedDict()
        self.themes['dark'] = 'design/qss/dark.qss'
        self.themes['light'] = 'design/qss/light.qss'
        self.themes['default'] = ':/invalid_path3243'

        # initialize style
        self.set_stylesheet('dark')


    def toggle_word_wrap(self, text_edit):
        '''
        toggles the word wrap in a text edit object
        '''
        if text_edit.lineWrapMode() == QTextEdit.NoWrap:
            text_edit.setLineWrapMode(QTextEdit.WidgetWidth)
        else:
            text_edit.setLineWrapMode(QTextEdit.NoWrap)


    def reset_progress_bar(self):
        '''
        resets the progress bar so long the emailer isn't being used
        '''
        if not self._sender_thread_is_running:
            self._init_progress_bar(self.ui.progress_bar)
        else:
            self.console_log('Honey, you know I can\'t reset the progress '
                             'bar while sending emails')


    def clear_console(self):
        '''
        clears the gui console
        '''
        self.ui.console.clear()


    def force_authorize(self):
        '''
        remove the current authorization
        to require a re-sign in from user
        '''
        self.console_log('Re-authorizing . . .')
        self.auth.remove()
        success = self.start_authorize_thread() 
        
        
    def _auth_already_started_msg(self):
        '''
        the credentials file was not found; 
        ask user to please visit a website
        to get API credentials
        :return OkCancel: int of QMessageBox StandardButtons enum
        '''
        mb = QMessageBox()
        mb.setIcon(QMessageBox.Warning)
        mb.setWindowTitle('Authorization Error')
        mb.setText(
            'The Google "Sign In" window should already be open in your browser! '
            'You are currently using '
            f'{self.threadpool.activeThreadCount()}/{self.threadpool.maxThreadCount()} '
            'of the available threads.\n\n'
            'Are you sure you want to add another?\n\n')
        mb.setDetailedText(
            'The authorization thread is already running, and I unfortunately'
            'cannot quit the thread from the outside, nor can I detect an'
            '"authorization cancelled" event from the OAuth2 service.\n\n'
            'This is a bummer.\n\n'
            'It\'s also a known issue: '
            'https://github.com/googleapis/google-api-dotnet-client/issues/968')
        mb.setStandardButtons(QMessageBox.Ok | QMessageBox.Cancel)
        mb.setDefaultButton(QMessageBox.Cancel)
        OkCancel = mb.exec()
        return OkCancel


    def start_authorize_thread(self):
        '''
        call the authorize function from another thread
        this should clear up the gui thread
        (prompt user to restart thread if it's already running)
        '''
        if self._auth_thread_is_running:
            try_again = self._auth_already_started_msg()
            if try_again == QMessageBox.Ok:
                pass
            else:
                self.console_log('Re-authorization cancelled.')
                return

        else:
            self.console_log('Starting authorization thread . . .')

        self._auth_thread_is_running = True

        worker = Worker(self.authorize)
        worker.signals.result.connect(self._authorize_result)
        worker.signals.finished.connect(self._authorize_thread_complete)
        worker.signals.error.connect(self._handle_thread_error)

        self._auth_worker = worker
        self.threadpool.start(worker)


    def _authorize_result(self, success):
        '''
        :success: boolean for whether the authorization was successful or not
        this gets the result of the authorize thread
        will be either true or false
        depending on whether it worked or not
        '''
        self._authorized = success
        if self._authorized:
            self.console_log('Authorized as {}'.format(self.auth.profile['emailAddress']))
            self.setWindowTitle('PyMailList - ' + self.auth.profile['emailAddress'])
        else:
            self.console_log('Authorization failed')

    def _authorize_thread_complete(self):
        '''
        let the people know the the authorize thread completed
        '''
        self.console_log('Authorization thread completed')
        self._auth_thread_is_running = False
    

    def authorize(self):
        '''
        safely try to authorize
        if the files can't be found a message is displayed
        (credentials file, namely)
        :return: authenticated service
        '''
        try:
            self.auth.start()
        except FileNotFoundError:
            self._credential_error()
            success = False
        else:
            self.email.service = self.auth.service
            success = True
        return success


    def _credential_error(self):
        '''
        the credentials file was not found; 
        ask user to please visit a website
        to get API credentials
        '''
        mb = QMessageBox()
        mb.setIcon(QMessageBox.Warning)
        mb.setWindowTitle('Authorization Error')
        mb.setText(
            'Authorization cancelled!\n\n'
            'It is likely your API credentials file ("credentials.json") was not found!\n\n'
            'Please visit https://developers.google.com/gmail/api/quickstart/python to get your credentials file.')
        mb.setDefaultButton(QMessageBox.Ok)
        mb.exec()


    def _form_complete_error(self):
        '''
        displays a notification to the user
        that their form is not yet complete
        '''
        mb = QMessageBox()
        mb.setIcon(QMessageBox.Warning)
        mb.setWindowTitle('Warning')
        mb.setText('Sending cancelled. Please ensure the form is complete.')
        mb.setDefaultButton(QMessageBox.Ok)
        mb.exec()


    def send(self):
        '''
        send the message in the gui
        to the contacts in the gui
        '''

        # we can't send if we are already sending
        if self._sender_thread_is_running:
            self.console_log('You\'re already sending messages, silly. Hang on a sec.')
            return

        # try to authorize yo-self
        if not self._authorized:
            self.console_log('You\'re not authorized yet! Press the "Authorize" button and try again')
            #auth_success = self.authorize()
            #if not auth_success: return
            return

        # ensure the form is complete
        if not self._form_complete():
            self._form_complete_error()            
            return

        # make the email from the GUI inputs
        self._make_email()

        # initialize progress bar
        self._init_progress_bar(pb = self.ui.progress_bar, 
                                max_val = len(self.contacts))

        # start email sender thread
        self._start_sender_thread()


    def _start_sender_thread(self):
        '''
        starts the email sender thread
        '''
        worker = Worker(self.send_runner,
                        email=self.email,
                        contacts=self.contacts)
        worker.kwargs['progress_callback'] = worker.signals.progress

        worker.signals.result.connect(self.console_log)
        worker.signals.finished.connect(self._send_thread_complete)
        worker.signals.error.connect(self._handle_thread_error)
        worker.signals.progress.connect(self._log_progress)

        self._sender_thread_is_running = True
        self.threadpool.start(worker)


    def _init_progress_bar(self, pb, max_val=100):
        '''
        initialize progress bar to zero
        '''
        pb.setMinimum(0)
        pb.setMaximum(max_val)
        pb.setValue(0)


    def _log_progress(self, log=None):
        '''
        logs the progress of a thread
        increments progress bar
        and logs a message
        '''
        self._increment_progress_bar()
        self.console_log(log)


    def send_runner(self, email:Emailer, contacts:list, progress_callback):
        '''
        sends email message individually
        to each contact in contacts list

        intended for use with threads so
        :param progress_callback: can indicate
        progress
        '''
        for address in contacts:
            email.message.to = address
            email.message.recreate()
            sent, e = email.send()

            if sent:
                status = 'done'
            else:
                status = f'Error! {str(e)}'
            progress_callback.emit(f'Just gonna send it to {address} . . . {status}')


    def _handle_thread_error(self, e):
        '''
        takes the worker thread error and
        processes it for the gui
        '''
        self.console_log('ERROR: ' + str(e[1]))


    def _send_thread_complete(self):
        '''
        display message once the thread that sends all the
        messages finishes up
        '''
        self._sender_thread_is_running = False
        self.console_log('Email sender thread completed.')


    def console_log(self, s):
        '''
        :param s: object to log (must have __str__ attribute)
        logs whatever is passed to it
        to the gui console
        '''
        if s is not None: self.ui.console.append(str(s))


    def _increment_progress_bar(self):
        '''
        increments progress bar by 1
        '''
        self.ui.progress_bar.setValue(self.ui.progress_bar.value() + 1)


    def _form_complete(self):
        '''
        returns true if all fields
        in form are filled out
        otherwise, returns false
        '''
        form_complete:bool = True

        if len(self.subject)==0:
            form_complete = False
        if len(self.body)==0:
            form_complete = False
        if len(self.contacts)==0:
            form_complete = False

        return form_complete


    def _make_email(self):
        '''
        creates the message based on
        the gui form fields
        '''
        self.email.message.create(
            to='',
            sender='me',
            subject=self.subject,
            body=self.body,
            body_type=self.body_type
            )
        return self.email.message


    @property
    def body_type(self):
        '''
        returns the radio box selection
        for the text type
        '''
        if self.ui.plain_text_radio_button.isChecked():
            return 'plain'
        elif self.ui.HTML_radio_button.isChecked():
            return 'html'
        else:
            # you'll want to turn this into a GUI message at some point I guess
            raise NotImplementedError('Please select a radio button for the message type')
        
    @property
    def subject(self):
        '''
        returns the subject of
        the message in the gui
        '''
        return self.ui.subject_line_edit.text()


    @property
    def body(self):
        '''
        returns the body of the
        message in the gui
        '''
        return self.ui.message_text_edit.toPlainText()


    @property
    def contacts(self):
        '''
        returns a list of the
        contacts in the gui
        '''
        return self.ui.contacts_text_edit.toPlainText().split('\n')


    def clear(self):
        self.ui.message_text_edit.clear()
        self.ui.contacts_text_edit.clear()
        self.ui.subject_line_edit.clear()
        self.console_log('Fields Cleared')


    def cycle_stylesheet(self):
        '''
        cycles through stylesheets for gui
        '''

        current_sheet = self._stylesheet
        sheets = self.themes

        # find the index of the current sheet
        keys = list(sheets.keys())
        index = keys.index(current_sheet)

        # go to the next index, or back to the start of the list
        if index == len(keys) - 1:
            new_index = 0
        else:
            new_index = index + 1

        self.set_stylesheet(keys[new_index])
        

    def set_stylesheet(self, theme_name):
        '''
        Toggle the stylesheet to use the desired path in the Qt resource
        system (prefixed by `:/`) or generically (a path to a file on
        system).

        :path:      A full path to a resource or file on system
        '''

        # get the QApplication instance,  or crash if not set
        app = QApplication.instance()
        if app is None:
            raise RuntimeError("No Qt Application found.")

        self._stylesheet = theme_name

        try:
            path = self.themes[theme_name]
        except:
            self.console_log(f'ERROR: Bad theme name "{theme_name}"')
            return

        self.console_log(f'Setting GUI theme to "{theme_name}"')

        file = QFile(path)
        file.open(QFile.ReadOnly | QFile.Text)
        stream = QTextStream(file)
        app.setStyleSheet(stream.readAll())


def main():
    app = QApplication(sys.argv)

    application = ApplicationWindow()
    application.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()