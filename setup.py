from distutils.core import setup
import py2exe
from glob import *

setup(
    console=[{"script": "ttAuditor.py", "icon_resources": [(1, "tt.ico")]}],
    options={
            "py2exe":{
                    "skip_archive": True,
                    "unbuffered": True,
                    "optimize": 2,
					"bundle_files": 1, 
            }
    },
    data_files = [(r'.', glob(r'C:\Python34\selenium\webdriver\remote\getAttribute.js')),(r'.', glob(r'C:\Python34\selenium\webdriver\remote\isDisplayed.js'))],

)