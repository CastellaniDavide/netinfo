"""netinfo
"""
import os, wmi, sqlite3
from datetime import datetime

__author__ = "help@castellanidavide.it"
__version__ = "01.01 2020-09-21"

class netinfo:
	def __init__ (self, debug=False):
		"""The core of my project
		"""
		base_dir = "." if debug else ".." # the project "root" in Visual studio it is different

		log = open(os.path.join(base_dir, "log", "trace.log"), "a")
		csv_names = open(os.path.join(base_dir, "flussi", "computers.csv"), "r")
		csv_netinfo = open(os.path.join(base_dir, "flussi", "netinfo.csv"), "w+")
		db_netinfo = sqlite3.connect(os.path.join(base_dir, "flussi", "netinfo.db")).cursor()

		netinfo.log(log, "Opened all files and database connected")

		intestation = "Caption,Description,Status,Manufacturer,Name,GuaranteesDelivery,GuaranteesSequencing,MaximumAddressSize,MaximumMessageSize,SupportsConnectData,SupportsEncryption,SupportsGracefulClosing,SupportsGuaranteedBandwidth,SupportsQualityofService"
		netinfo.init_csv(csv_netinfo,intestation, log)
		
		start_time = datetime.now()
		netinfo.log(log, f"Start time: {start_time}")
		netinfo.log(log, "Running: netinfo.py")

		netinfo.print_all(log, csv_names, csv_netinfo, db_netinfo, debug)
		netinfo.check_db(db_netinfo, intestation, log)

		netinfo.log(log, f"End time: {datetime.now()}\nTotal time: {datetime.now() - start_time}")
		netinfo.log(log, "")
		log.close()

	def print_all(log, csv_names, csv_netinfo, db_netinfo, debug=False):
		"""Prints the infos by Win32_NetworkClient & Win32_NetworkProtocol
		"""
		for PC_name in ("My PC, debug option",) if debug else csv_names.read().split("\n")[1:]:
			
			conn = wmi.WMI("" if debug else PC_name)
			netinfo.print_and_log(log, f" - {PC_name}")

			for network_client, network_protocol in zip(conn.Win32_NetworkClient(["Caption", "Description", "Status", "Manufacturer", "Name"]), conn.Win32_NetworkProtocol(["GuaranteesDelivery", "GuaranteesSequencing", "MaximumAddressSize", "MaximumMessageSize", "SupportsConnectData", "SupportsEncryption", "SupportsEncryption", "SupportsGracefulClosing", "SupportsGuaranteedBandwidth", "SupportsQualityofService"])):
				netinfo.print_and_log(log, "   - Istructions: Win32_NetworkClient && Win32_NetworkProtocol")
				data = f"'{network_client.Caption}','{network_client.Description}','{network_client.Status}','{network_client.Manufacturer}','{network_client.Name}','{network_protocol.GuaranteesDelivery}','{network_protocol.GuaranteesSequencing}','{network_protocol.MaximumAddressSize}','{network_protocol.MaximumMessageSize}','{network_protocol.SupportsConnectData}','{network_protocol.SupportsEncryption}','{network_protocol.SupportsGracefulClosing}','{network_protocol.SupportsGuaranteedBandwidth}','{network_protocol.SupportsQualityofService}'"
				csv_netinfo.write(data.replace("'", ""))
				db_netinfo.execute(f"INSERT INTO netinfo VALUES ({data})")

	def log(file, item):
		"""Writes a line in the log.log file
		"""
		file.write(f"{item}\n")

	def print_and_log(file, item):
		"""Writes on the screen and in the log file
		"""
		print(item)
		netinfo.log(file, item)

	def init_db(db_netinfo, intestation, log):
		"""Init the database
		"""
		try:
			netinfo.init_db(db_netinfo, intestation, log)
		except:
			db_netinfo.execute("DROP TABLE netinfo")
			netinfo.init_db(db_netinfo,intestation, log)

		db_netinfo.execute(f'''CREATE TABLE netinfo ({intestation})''')
		netinfo.log(log, "database now initialized")

	def init_csv(csv_netinfo, intestation, log):
		"""Init the csv files
		"""
		csv_netinfo.write(f"{intestation}\n")
		netinfo.log(log, "csv now initialized")

	def check_db(db_file, intestation, log, tablename="netinfo"):
		"""Checks the database content
		"""
		print("\nDatabase content:")
		print(f"{intestation}")

		for row in db_file.execute(f"SELECT * FROM {tablename}"):
			print(str(row)[1:-1].replace("'", ""))

if __name__ == "__main__":
	# degub flag
	debug = True

	netinfo(debug)
