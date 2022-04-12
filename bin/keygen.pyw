import os
import secrets
import string
import subprocess
from delphivcl import *
from winreg import HKEY_CURRENT_USER, QueryValueEx, OpenKey

# Base directory of the project
BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),'..')

class MainForm(Form):
    def __init__(self, owner):
        self.lblTitle = None
        self.pnlEdit = None
        self.lblLength = None
        self.edtLength = None
        self.chkSymbols = None
        self.chkNumbers = None
        self.chkLetters = None
        self.pnlAction = None
        self.bttKeygen = None
        self.bttCopy = None
        self.lblAbout = None
        self.edtKeygen = None
        # Load GUI components properties
        self.LoadProps(os.path.join(BASE_DIR,'etc','gui','components.pydfm'))
        # Events
        self.bttKeygen.OnClick = self.__gen_key
        self.bttCopy.OnClick = self.__copy_to_clipboard
        self.lblAbout.OnClick = self.__open_link
        self.chkLetters.OnClick = self.__no_focus
        self.chkNumbers.OnClick = self.__no_focus
        self.chkSymbols.OnClick = self.__no_focus

        # Style actions
        self.__sm = StyleManager()
        self.__load_style()
        # Other actions
        self.OnClose = self.__on_form_close

    def __on_form_close(self, sender, action):
        """Exit application"""
        action.Value = caFree

    def __gen_key(self, sender):
        """Generate random key"""
        length = int(self.edtLength.Value)
        if length > 0 and any([self.chkLetters.Checked, self.chkNumbers.Checked, self.chkSymbols.Checked]):
            char = (string.ascii_letters if self.chkLetters.Checked else '')+(
                          string.digits if self.chkNumbers.Checked else '')+(
                          string.punctuation if self.chkSymbols.Checked else '')
            key = ''.join(secrets.choice(char) for i in range(length))
            self.edtKeygen.Text = key

    def __copy_to_clipboard(self, sender):
        """Copy key to clipboard"""
        subprocess.run('clip', input=self.edtKeygen.Text, universal_newlines=True, creationflags=subprocess.CREATE_NO_WINDOW)

    def __open_link(self, sender):
        """Open about link"""
        os.startfile('https://github.com/ariel-mn')

    def __no_focus(self, sender):
        """Disable on click focus"""
        self.ActiveControl = None

    def __load_style(self):
        """Auto-load style theme from file"""
        try:
            Key = OpenKey(HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize")
            UseLightTheme = QueryValueEx(Key, "AppsUseLightTheme")[0]
            if not UseLightTheme:
                self.__sm.LoadFromFile(os.path.join(BASE_DIR,'etc','styles','Windows11_Modern_Dark.vsf'))
            else:
                self.__sm.LoadFromFile(os.path.join(BASE_DIR,'etc','styles','Windows11_Modern_Light.vsf'))
            self.__sm.SetStyle(self.__sm.StyleNames[1])
        except FileNotFoundError:
            return

def main():
    """Launch application"""
    Application.Initialize()
    Application.Title = 'Keygen App'
    Application.MainFormOnTaskBar = True # Allows minimization of the application
    app = MainForm(Application)
    app.Show()
    # Comment FreeConsole() for debugging
    #FreeConsole()
    Application.Run()
    app.Destroy()

if __name__ == '__main__':
    main()