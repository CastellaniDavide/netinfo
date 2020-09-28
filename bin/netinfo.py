"""netinfo
"""
import os
import wmi
import sqlite3
from datetime import datetime

__author__ = "help@castellanidavide.it"
__version__ = "02.01 2020-09-28"

class netinfo:
	def __init__ (self, debug=False, db=True):
		"""Where it all begins
		"""
		base_dir = "." if debug else ".." # the project "root" in Visual studio it is different

		log = open(os.path.join(base_dir, "log", "trace.log"), "a+")
		csv_names = open(os.path.join(base_dir, "flussi", "computers.csv"), "r+")
		csv_netinfo_history = open(os.path.join(base_dir, "flussi", "netinfo_history.csv"), "a+")
		csv_netinfo = open(os.path.join(base_dir, "flussi", "netinfo.csv"), "w+")
		csv_unchecked = open(os.path.join(base_dir, "flussi", "unchecked_PC.csv"), "r+").read()
		csv_unchecked2 = open(os.path.join(base_dir, "flussi", "unchecked_PC.csv"), "w+")
		if db : db_netinfo = sqlite3.connect(os.path.join(base_dir, "flussi", "netinfo.db"))
		
		netinfo.log(log, f"Opened all files{' and database connected' if db else ' opened'}")
		
		start_time = datetime.now()
		netinfo.log(log, f"Start time: {start_time}")
		netinfo.log(log, "Running: netinfo.py")
		
		# Init files and database
		intestation = "Caption,Description,Status,Manufacturer,Name,GuaranteesDelivery,GuaranteesSequencing,MaximumAddressSize,MaximumMessageSize,SupportsConnectData,SupportsEncryption,SupportsGracefulClosing,SupportsGuaranteedBandwidth,SupportsQualityofService,Date_local,Date_universal_microsecond"
		intestation_unchecked = '"names","fail_reach","total_search"'
		netinfo.init_csv(csv_netinfo, intestation, log)
		csv_netinfo_history.seek(0)
		if csv_netinfo_history.read() == "" : netinfo.init_csv(csv_netinfo_history, intestation, log)
		if db : netinfo.init_db(db_netinfo.cursor(), intestation, intestation_unchecked, log, db)

		# Core
		fail = netinfo.print_all(log, csv_names, csv_netinfo, csv_netinfo_history , db_netinfo.cursor() if db else None, debug, db)
		netinfo.update_unchecked_PC(fail, csv_unchecked, csv_unchecked2, db_netinfo.cursor() if db else None, log, debug, db)
		
		if db and debug : netinfo.print_db(db_netinfo.cursor(), intestation, intestation_unchecked, log, db)

		# End
		csv_names.close()
		csv_netinfo.close()
		csv_netinfo_history.close()
		csv_unchecked2.close()
		if db : db_netinfo.commit()
		netinfo.log(log, f"End time: {datetime.now()}\nTotal time: {datetime.now() - start_time}")
		netinfo.log(log, "")
		log.close()

	def print_all(log, csv_names, csv_netinfo, csv_netinfo_history, db_netinfo, debug=False, db=True):
		"""Prints the infos by Win32_NetworkClient & Win32_NetworkProtocol
		"""
		fail = []
		for PC_name in ("My PC, debug option",) if debug else csv_names.read().split("\n")[1:]:

			if netinfo.check_PC("localhost" if debug else PC_name):			
				conn = wmi.WMI("." if debug else PC_name)
				netinfo.print_and_log(log, f" - {PC_name}", debug)

				for network_client, network_protocol in zip(conn.Win32_NetworkClient(["Caption", "Description", "Status", "Manufacturer", "Name"]), conn.Win32_NetworkProtocol(["GuaranteesDelivery", "GuaranteesSequencing", "MaximumAddressSize", "MaximumMessageSize", "SupportsConnectData", "SupportsEncryption", "SupportsEncryption", "SupportsGracefulClosing", "SupportsGuaranteedBandwidth", "SupportsQualityofService"])):
					netinfo.print_and_log(log, "   - Istructions: Win32_NetworkClient && Win32_NetworkProtocol", debug)
					data = f"'{network_client.Caption}','{network_client.Description}','{network_client.Status}','{network_client.Manufacturer}','{network_client.Name}','{network_protocol.GuaranteesDelivery}','{network_protocol.GuaranteesSequencing}','{network_protocol.MaximumAddressSize}','{network_protocol.MaximumMessageSize}','{network_protocol.SupportsConnectData}','{network_protocol.SupportsEncryption}','{network_protocol.SupportsGracefulClosing}','{network_protocol.SupportsGuaranteedBandwidth}','{network_protocol.SupportsQualityofService}','{datetime.now()}','{int(datetime.utcnow().timestamp() * 10 ** 6)}'"
				
					csv_netinfo.write(f"""{netinfo.make_csv_standart(data).replace("'", '"')}\n""")
					csv_netinfo_history.write(f"""{netinfo.make_csv_standart(data).replace("'", '"')}\n""")
					if db : db_netinfo.execute(f"INSERT INTO netinfo VALUES ({data})")
					if db : db_netinfo.execute(f"INSERT INTO netinfo_history VALUES ({data})")
			else:
				fail.append(PC_name)

		return fail

	def check_PC(PC_name, n_test=2):
		"""Checks if PC is disponible
		"""
		respond = os.popen(f"PING -n {n_test} {PC_name}").read()	# Windows
		respond += os.popen(f"PING -i {n_test} {PC_name}").read()	# Linux

		return f"Received = {n_test}" in respond

	def make_csv_standart(data):
		"""Convert my text in csv standard to prevent errors
		Reference: https://tools.ietf.org/html/rfc4180
		"""
		data.replace(",", "%x2C").replace("}%x2C{", "},{")	# make sure extra commas
		data.replace("\"", "%x22").replace("%x22", "\"", 1).replace("%x22", "\"") # make sure extra double commas
		data.replace("'", '"') # Use " and not ', as csv standard
		data.encode("ASCII")
		return data

	def log(file, item):
		"""Writes a line in the log.log file
		"""
		file.write(f"{item}\n")

	def print_and_log(file, item, debug):
		"""Writes on the screen and in the log file
		"""
		if debug : print(item)
		netinfo.log(file, item)

	def init_db(db_netinfo, intestation, intestation_unchecked, log, db=True):
		"""Init the database
		"""
		if db : db_netinfo.execute(f'''CREATE TABLE IF NOT EXISTS netinfo ({intestation})''')
		if db : db_netinfo.execute(f'''DELETE FROM netinfo''')

		if db : db_netinfo.execute(f'''CREATE TABLE IF NOT EXISTS netinfo_history ({intestation})''')

		if db : db_netinfo.execute(f'''CREATE TABLE IF NOT EXISTS unchecked_PC ({intestation_unchecked})''')
		if db : db_netinfo.execute(f'''DELETE FROM unchecked_PC''')
		
		netinfo.log(log, "database now initialized")

	def init_csv(flie, intestation, log):
		"""Init the csv files
		"""
		flie.write(f"{intestation}\n")
		netinfo.log(log, "csv now initialized")

	def print_db(db_file, intestation, intestation_unchecked, log, debug, tablename_main="netinfo", tablename_history="netinfo_history", tablename_unchecked="unchecked_PC"):
		"""Checks the database content
		"""
		# Main content
		if debug : print("\nDatabase main content:")
		if debug : print(intestation)

		for row in db_file.execute(f"SELECT * FROM {tablename_main}").fetchall():
			if debug : print(str(row)[1:-1].replace("'", ""))

		# History content
		if debug : print("\n\nDatabase history content:")
		if debug : print(intestation)

		for row in db_file.execute(f"SELECT * FROM {tablename_history}").fetchall():
			if debug : print(str(row)[1:-1].replace("'", ""))

		# Unchecked content
		if debug : print("\n\nDatabase unchecked content:")
		if debug : print(intestation_unchecked)

		for row in db_file.execute(f"SELECT * FROM {tablename_unchecked}").fetchall():
			if debug : print(str(row)[1:-1].replace("'", ""))

	def update_unchecked_PC(unchecked_PCs, csv_unchecked_PC, csv_out, db_unchecked_PC, log, debug, db):
		""" Update unchecked PC
		"""
		# Make my updated array
		last = netinfo.csv2array(csv_unchecked_PC[:-1:])
		array = []

		array.append(last[0]) # Add intestation

		for item in last[1::]:
			if item[0] in unchecked_PCs:
				array.append([item[0], item[1] + 1, item[2] + 1])
				unchecked_PCs.remove(item[0])
			else:
				array.append([item[0], item[1], item[2] + 1])

		for item in unchecked_PCs:
			array.append([item, 1, 1])

		# Update csv
		csv_out.write(netinfo.array2csv(array))

		# Update db
		if db : db_unchecked_PC.executemany('INSERT INTO unchecked_PC VALUES (?,?,?)', array[1::])

		netinfo.log(log, "Unchecked updated")

	def csv2array(csv):
		""" Converts csv file to a py array
		"""
		array = []

		for line in csv.split("\n"):
			temp = []
			for item in line.replace(",", "','").split("','"):
				try:
					temp.append(int(item.replace('"', "")))
				except:
					temp.append(item.replace('"', ""))
			array.append(temp)

		return array

	def array2csv(array):
		""" Converts py array to a csv file
		"""
		text = ""

		for line in array:
			for item in line:
				text += f'"{item}",'
			text = text[:-1:] + "\n"

		return text

if __name__ == "__main__":
	# debug/visual studio flag
	debug = False

	# database flag
	db = True

	netinfo(debug, db)
