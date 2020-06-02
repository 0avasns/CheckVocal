from distutils.core import setup
import py2exe

# We need to import the glob module to search for all files.
import glob

class Target:
    def __init__(self, **kw):
        self.__dict__.update(kw)
        self.version = "3.0.2.0"
        self.company_name = "Athanassios Protopapas"
        self.copyright = "(c) 2020 -- June 2nd"
        self.name = "CheckVocal"

CV = Target(
    description = "DMDX Vocal Response & RT Check Utility",
    version = "3.0.2.0",
    script = "CheckVocal.pyw",
    icon_resources = [(1, "cv.ico")],
    dest_base = "CheckVocal")

CF = Target(
    description = "Audio File RT Check Utility",
    version = "3.0.2.0",
    script = "CheckVocal.pyw",
    icon_resources = [(1, "cf.ico")],
    dest_base = "CheckFiles")

AZT = Target(
    description = "DMDX output converter: .azk to .txt format",
    version = "1.3.0.0",
    script = "azk2txt.pyw",
    icon_resources = [(1, "a2t.ico")],
    dest_base = "azk2txt")

setup(
      options={ "py2exe":
               {"packages": ["encodings"],
                "compressed": 1,
                "optimize": 2}
              },
      windows=[CF,CF,AZT,CV], # makes no sense but whatever is first does not get an icon on the exe, so I duplicated it ...
      data_files=[
                  (r'tcl\snacklib', glob.glob(r'C:\python27\tcl\snack2.2\*')),
                  ("icons",["cv.ico","cf.ico","a2t.ico"]),
                  (".",["README-CheckVocal.txt"])
                 ],
     )
