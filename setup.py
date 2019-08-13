from cx_Freeze import setup, Executable
setup(name="Compare DNA",
      version="0.2",
      description="compares sequences of dna and color matches",
      executables = [Executable("MultiDNAcolorMatch.py")])

'''
In order to actually make the executable you
need to first change the directory in the cmd
to the location of the folder e.g. on windows
cd /path/to/file

which should have both this file, which you can
call the setup file, and also the file itself
that we are trying to turn into an exectutable,
in this case "distme.py".

once we are in the file directory type:
setup.py build (in this case we type setup.py,
because that is the name of this file)

it should then make the executable
'''
