
import os
import tempfile

filename = tempfile.mktemp(".txt")

open (filename , "w").write ("How to Print Hard Copies in Python")

os.startfile(filename, "print")