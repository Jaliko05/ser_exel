import os

from get_txt import get_routs
from pathlib import Path

rout_aplication = str(Path(__file__).parent.absolute())# ruta SIIFNET
print(rout_aplication)
rout = get_routs(rout_aplication,"24")
print(rout)