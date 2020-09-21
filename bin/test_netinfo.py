"""Test netinfo file
"""
from netinfo import *

__author__ = "help@castellanidavide.it"
__version__ = "1.0 2020-9-21"

def test():
	"""Tests the netinfo function in the netinfo class
	Write here all test you want to do.
	REMEMBER to test your programm you can't use __init__ function
	"""
	assert netinfo.netinfo() == "netinfo", "test failed"
	#assert netinfo.<function>(<values>) == <the result(s) you would like to have>, "<the fail message>"
	
if __name__ == "__main__":
	test()
